r"""答案生成器（识别侧）——**2026-09-20 按重写后的 3.6.2 识别表从零重建**。

按《工具操作逻辑》3.6.2「文档的识别」的字面判据，算出每一段"应该被识别成
什么"——供写法层、结构层、期望对拍当**独立的标准答案**，与真实引擎逐段对拍。

新表把识别拆成三个识别器，本文件就照这个结构实现（编号是本文件自己编的，
用户不需要；报告里用它指回表里哪一行）：

    F1~F4  前向识别器·四个标签（摘要/关键词/英文摘要/英文关键词）
    T1~T4  前向识别器·题目与副标题（中文正标题/中文副标题/英文正标题/英文副标题）
    B1~B3  后向识别器·三个后置标签（参考文献/附录/致谢）
    H1~H6  后向识别器·标题（一～四级标题/引言/结论）
    S1~S4  补充识别器（题注/目录三清单标签/目录三清单域/附录序号题名）
    G      归属划分（其余对象归入上方最近的标签部件，没有归属的按普通正文）

铁律（沿用）：
  * 判据**只许抄表里的文字**，不许照 `essay_tool/engine/` 里的识别实现写；
  * 表里没写清的地方一律按**常理**推断，并登记进 `ASSUMPTIONS`；
  * 自动编号、目录域这些"结构事实"本文件自己扫 XML，不借引擎的探测函数；
  * 编号（SEQ 域）一个字不碰——那属于加工侧。

用法（运行方式与重建前一致）：
    .venv\Scripts\python.exe "Adversarial Automatic Test\aat_expect.py" reference\民总.docx
"""

from __future__ import annotations

import argparse
import dataclasses
import pathlib
import re
import shutil
import sys
import zipfile
from dataclasses import dataclass, field

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent
for entry in (str(ROOT), str(HERE)):
    if entry not in sys.path:
        sys.path.insert(0, entry)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import aat_diff                                             # noqa: E402
from lxml import etree                                      # noqa: E402

from aat_docx import (DocShape, Para, attrs_of, local_name,  # noqa: E402
                      qn, read_docx)
from aat_engine_api import expect_index                      # noqa: E402
from essay_tool.config.preset_io import load_preset          # noqa: E402
from essay_tool.engine import pipeline                       # noqa: E402
from essay_tool.engine.placeholders import (
    is_english_subtitle_placeholder, is_english_title_placeholder,
    is_placeholder, is_subtitle_placeholder, is_title_placeholder)

PRESET = ROOT / "presets" / "中国政法大学-本科毕业论文.json"
OUT = HERE / "_out" / "期望"

# ---------------------------------------------------------------- 按常理推断的
# 表里没写清、本实现不得不选的东西。**不许偷偷拿引擎的做法**；每一条都写下来，
# 报告里跟着期望结果一起打给用户看。
ASSUMPTIONS: list[str] = [
    "「句末标点」表里没列举：按项目既有口径取一套——。！？；，、与半角 .!?;,。"
    "省略号、右括号、右引号、书名号不算（“三、或许……”这类以省略号收尾的仍算"
    "标题）。三级/四级原写“标点”是用户 2026-09-20 确认的漏写，已把表改回"
    "“句末标点”，全层级统一。",
    "“不以（句末）标点结尾”查的是**编号之后那一截标题文字**；没有标题文字的"
    "（如“1.”“（一）”）不算以标点结尾。",
    "中文数字的承认范围：表只说了“百、千不能作首字”“廿、卅不被承认”。按常理"
    "裁成 13 个常用字“零一二三四五六七八九十百千”——两、万、〇、亿 不列（否则"
    "“万一、”“两、”这类正文开头的句子会被吸成标题）。",
    "60/40/120 字的上限按**去掉全部空格**（半角+全角）后的长度量。",
    "Word 自动编号的“反算”表里没写做法：本文件自己读 word/numbering.xml，按多级"
    "列表常识算（同级 +1、更深层清零、lvlText 的 %1~%9 按各级格式填充）——不调"
    "引擎的编号模块（共用实现会一起错，差分看不见）。反算不出的段落登记“未判定”。",
    "「目录域区间」（域指令所在段之后、域结束之前）按字面取**严格夹在中间**的"
    "段落：带 begin 的指令段、带 end 的结束段都不算条目。嵌套域（条目自带的"
    "PAGEREF 等）净深度为零，不影响区间。",
    "「以 TOC 开头的域」= 域指令去掉前导空白后以 TOC 起头、不限大小写；「收集"
    "标题的」= 指令里有 \\o 参数；「收集题注的」= \\c \"图\" / \\c \"表\" / "
    "\\t \"图题,1\" / \\t \"表题,1\" 这些写法（表只说“收集”，没写参数长什么样）。",
    "「题注功能插入的 SEQ 域」= 段落域指令里出现 SEQ 且标签精确是 图/表/fig/"
    "figure/table 之一（“图表”“公式”不算）。",
    "「题注」「Caption」「toc 1/2/3」「table of figures」都按**样式名**比（忽略"
    "大小写）。",
    "前后扫描器都**只看正文段落**：表格单元格里的段落、内容控件里没被拆出来的"
    "段落不在识别范围内（与既有管线的可见性口径一致；`prepare()` 已先拆包）。",
    "停止条件②「前向识别器已经认出的部件」把前向认出的题目与副标题也算在内。",
    "重复出现的后置标签（参考文献/附录/致谢/结论）**不新开段位、仍然触发“退回"
    "pending 标题”的副作用**（表说“每识别到一个后置标签…一律退回”，未把重复者"
    "排除）；被退回的标题不再参与④的比对（cjk 表一并清空）。",
    "停止条件④“汉字部分”= 只保留汉字（含中文数字）后的文字；“全同”按这个规范形"
    "比；“连续扫描的两段”= 倒序扫描里相邻的两个标题候选（中间不能隔着别的候选）。"
    "触发后不再往前看：更靠前的段落交给归属划分，按普通正文处理。",
    "题目/副标题/英文正标题/英文副标题**都不把工具自插的占位语当候选**——只有表上"
    "明写的“英文题目占位文字”例外（T3 行里写了“工具自己插入的英文题目占位文字"
    "同样会被认出”）。占位语的样子按产品自己的定义（`placeholders.py`，它是"
    "产品事实、不是识别判据）。",
    "T1 那句“位于所有已识别标签和标题之前”：**英文题目算“已识别标题”**——用户"
    "2026-09-20 裁决，英文行排在中文题目之前时，后面的中文题目/副标题不再认"
    "（与引擎一致；T3/T4 因此在 T1/T2 之前先算）。",
    "用户 2026-09-20 裁决：**中文副标题不能是全英文的**——全英文行只能被 T3"
    "（英文题目）收走（`zmarkA` 这类短英文段即此例）。",
    "“全英文”= 不含 CJK 统一表意文字（U+4E00~U+9FFF）**且含至少一个 ASCII 字母**"
    "（用户 2026-09-20 裁决：只由下划线/标点构成的行不算全英文，可以当题目/副标题，"
    "照引擎现状）；“拉丁字母”数按 ASCII A-Za-z 计。",
    "段内软换行（w:br，Shift+Enter）**不认标题**——用户 2026-09-20 裁决保持现状；"
    "aat_docx 把 w:br 读成换行符，标题正则不跨行，与引擎一致。",
    "归属划分：标签部件共 8 个（摘要/关键词/英文摘要/英文关键词/引言/结论/参考"
    "文献/附录/致谢）；目录与三清单标签之后**不归属**（表里明说）。“下一个已"
    "识别的身份”按全部已识别对象算（标签、题目、各级标题、题注）；域生成的"
    "条目、占位语与附录序号题名**不算身份、也不打断归属**（后两种读法下角色"
    "都一样——锚点部件的正文本来也报“正文”）。",
    "“紧邻”= 段落下标正好差 1（中间隔着空段就不算）——表里明说副标题“中间不得"
    "隔有空段”。",
    "英文题注弱判据的“上一段（跳过空段）必须是中文题注”：按**已识别**的题注算"
    "（不是“看起来像题注”），且该上一段是中文行（续行也是中文题注）。",
    "续图/续表第三档“在数字之后紧跟”：数字编号段之后（允许空格）直接是“续图/"
    "续表”字样才算（如“表 1 续表”）。",
    "中文序号层级的三级标题**点号后不接数字**（`1.1`/`1.4` 这类不认）——用户"
    "2026-09-20 裁决“不存在问题”，照引擎现状收口；阿拉伯序号层级下“1.4 xx”仍是"
    "二级标题。",
    "阿拉伯序号层级：行 1 的“只允许半角数字和半角点号；每一段数字最多包含两"
    "位”与“点号后须跟空格，或直接跟汉字、字母”按**全级**口径收口——一~四级都"
    "按“1~4 段、每段 1~2 位数字，编号后可选一个半角点号，点号后须跟空格或直接"
    "跟汉字/字母，题名不得为空”实现（否则“1.111 标题”会被粘成“1.11”+题名"
    "“1 标题”）。",
    "括号类标题（二级/四级）的左右括号**都要在**：用户 2026-09-20 口径，“不必成对"
    "使用”指两边可各用半角/全角，不是可以缺一边。四级“加点号”仍按宽松收"
    "（“（1.”也认）。",
    "附录序号题名（S4）：“字母后可跟空格、顿号、点号、冒号”按**允许**理解、不强制；"
    "只查“附录+字母”前缀、长度、句末标点（“附录AB”“附录A调查问卷”都不额外否决）。",
    "标题长度上限（60 字）被超过时不算标题，登记为“疑似标题”（A2）；中文序号层级"
    "只对三、四级适用，阿拉伯数字序号层级一至四级都适用。",
]


# ---------------------------------------------------------------- 标点与常量

SENTENCE_PUNCT = "。！？；，、.!?;,"
ANY_PUNCT = SENTENCE_PUNCT + "：:"
CN_NUMERAL = "零一二三四五六七八九十百千"
CN_FORBIDDEN_INITIAL = {"百", "千"}
TITLE_MAX = 60                 # 标题长度上限（去空格后）
CAPTION_MAX = 60               # 题注长度上限（表：不多于六十个字；去空格量）
APPENDIX_MAX = 40              # 附录序号题名（不超过 40 个字；去空格量）
EN_TITLE_MIN_LETTERS = 5
EN_TITLE_MAX_LEN = 120

CJK_RE = re.compile(r"[\u4e00-\u9fff]")
LATIN_RE = re.compile(r"[A-Za-z]")


def squeezed(text: str) -> str:
    r"""去掉全部空白（含半角/全角空格与换行）——60/40/120 字上限按这个量。

    与引擎 `_squeeze_len` 同口径（`\s` + 全角空格）；`w:br` 读成 `\n` 之后
    这里要一起收，否则带软换行的段落量出来会多 1。
    """
    return re.sub(r"\s|　", "", text)


def ends_with(text: str, chars: str) -> bool:
    stripped = text.rstrip()
    return bool(stripped) and stripped[-1] in chars


def ends_sentence(text: str) -> bool:
    return ends_with(text, SENTENCE_PUNCT)


def ends_any_punct(text: str) -> bool:
    return ends_with(text, ANY_PUNCT)


def is_all_english(text: str) -> bool:
    """“全英文”= 不含汉字、**且含至少一个拉丁字母**（题目判据里用它）。

    与引擎 `_find_title_subtitle.all_english` 同口径；用户 2026-09-20 裁决：
    只由下划线/标点构成的行（无汉字也无字母）**不算**“全英文”，可以当题目/
    副标题（照引擎现状）。
    """
    return not CJK_RE.search(text) and bool(LATIN_RE.search(text))


# ---------------------------------------------------------------- 字面标签

F_LABELS: dict[str, str] = {           # 部件名 → 字面量（F1~F4 的顺序就是这里的顺序）
    "摘要": "摘要", "关键词": "关键词",
    "英文摘要": "ABSTRACT", "英文关键词": "KEYWORD",
}
B_LABELS: dict[str, str] = {"参考文献": "参考文献", "附录": "附录", "致谢": "致谢"}
TOC_LABELS: tuple[str, ...] = ("目录", "插图清单", "附表清单", "插图和附表清单")
INTRO_WORDS: tuple[str, ...] = ("引言", "绪论", "序言", "前言", "导论", "导言",
                                "引论", "绪言")
CONCLUSION_WORD = "结论"
ANCHORS: tuple[str, ...] = ("引言", "结论", "参考文献", "附录", "致谢")
OWNER_PARTS: tuple[str, ...] = ("摘要", "关键词", "英文摘要", "英文关键词",
                                "引言", "结论", "参考文献", "附录", "致谢")

_OPT_OPEN = r"[【\[]?"
_OPT_CLOSE = r"[】\]]?"


def _spaced(word: str) -> str:
    """字与字之间允许半角/全角空格。"""
    return r"[ 　]*".join(re.escape(char) for char in word)


def match_cjk_label(text: str, word: str) -> bool:
    """“包括且只包括 X 的段落”：字间可有空格、外面可有【】或[]（单边不超过一个）。"""
    return bool(re.fullmatch(rf"{_OPT_OPEN}{_spaced(word)}{_OPT_CLOSE}", text.strip()))


def match_ascii_label(text: str, word: str) -> bool:
    return bool(re.fullmatch(rf"{_OPT_OPEN}{re.escape(word)}{_OPT_CLOSE}",
                             text.strip(), re.IGNORECASE))


def match_keywords_label(text: str) -> bool:
    r"""英文关键词：KEYWORD / KEYWORDS（大小写不敏感）；允许拆开写作 Key Word(s)。"""
    return bool(re.fullmatch(rf"{_OPT_OPEN}key[ 　]*words?{_OPT_CLOSE}",
                             text.strip(), re.IGNORECASE))


def match_inline_keyword(text: str, part: str) -> bool:
    r"""关键词的“同段写法”（规则表 3.6.2 第 2、4 行第二款）：关键词三字（英文：
    KEYWORD(S) / Key Word(s)）带可选的【】或[]，其后紧跟“】”“]”、冒号（: ：）
    或空白，且后面还有内容。三字后没有任何分隔（“关键词是……”）不算。
    判据按规则表字面写，不抄引擎的正则。"""
    if part == "关键词":
        word = _spaced("关键词")
    elif part == "英文关键词":
        word = r"key[ \u3000]*words?"
    else:
        return False
    flags = re.IGNORECASE | re.DOTALL
    match = re.fullmatch(rf"\s*{_OPT_OPEN}{word}(?P<tail>[】\]]?[:：]?[ \u3000\t]*)(?P<rest>\S.*)",
                         text, flags)
    return bool(match and match.group("tail"))


def match_f_label(text: str, part: str) -> bool:
    if part == "英文摘要":
        return match_ascii_label(text, "ABSTRACT")
    if part == "英文关键词":
        return match_keywords_label(text) or match_inline_keyword(text, part)
    if part == "关键词":
        return match_cjk_label(text, F_LABELS[part]) or match_inline_keyword(text, part)
    return match_cjk_label(text, F_LABELS[part])


def _strip_arabic_prefix(text: str) -> str:
    """阿拉伯序号层级下，引言/结论允许“附加阿拉伯数字编号”：剥掉 '1 ' / '1.' 前缀。"""
    match = re.match(r"^\d{1,2}\s*[.．、]?\s*(.+)$", text)
    return match.group(1).strip() if match else text


def match_intro(text: str, *, arabic: bool) -> str | None:
    body = _strip_arabic_prefix(text) if arabic else text
    for word in INTRO_WORDS:
        if match_cjk_label(body, word):
            return word
    return None


def match_conclusion(text: str, *, arabic: bool) -> bool:
    body = _strip_arabic_prefix(text) if arabic else text
    return match_cjk_label(body, CONCLUSION_WORD)


# ---------------------------------------------------------------- 标题判据
# H1~H4。返回 (级别, 规则 ID, 编号原文, 标题文字) 或 None。


def _cn_number_ok(chars: str) -> bool:
    return bool(chars) and chars[0] not in CN_FORBIDDEN_INITIAL


def heading_match(text: str, scheme: str, *,
                  ignore_length: bool = False) -> tuple[int, str, str, str] | None:
    r"""表 H1~H4 的标题判据（H5/H6 的标签匹配在 `match_intro`/`match_conclusion`）。

    长度上限（60）按表：中文序号层级只对三、四级；阿拉伯数字序号层级一至四级。
    `ignore_length=True` 给“疑似标题”探形用（形状与标点照查，只是不因长度否决）。
    """
    stripped = text.strip()
    if not stripped:
        return None

    def finish(level: int, raw: str, title: str, *, sentence: bool,
               capped: bool) -> tuple[int, str, str, str] | None:
        if capped and not ignore_length and len(squeezed(stripped)) > TITLE_MAX:
            return None
        if title and (ends_sentence(title) if sentence else ends_any_punct(title)):
            return None
        return level, f"H{level}", raw, title

    if scheme == "dotted":
        # 阿拉伯数字序号层级：先长后短地试（x.y.z.f → x.y.z → x.y → x）。
        # 表（行 1）：每一段数字最多两位；编号后可选点号，随后须跟空格、或直接
        # 跟汉字/字母——这两句按同一口径套到一~四级（见 ASSUMPTIONS）。少了
        # “跟空格或汉字/字母”这道闸，"1.111"会被粘成"1.11"+题名"1 标题"。
        for level, segments in ((4, 4), (3, 3), (2, 2), (1, 1)):
            pattern = (r"^(\d{1,2}(?:\.\d{1,2}){%d})\.?"
                       r"(?:[ \u3000]+(?=.)|(?=[\u4e00-\u9fffA-Za-z]))(.*)$"
                       % (segments - 1))
            match = re.match(pattern, stripped)
            if match:
                return finish(level, match.group(1), match.group(2).strip(),
                              sentence=True, capped=True)
        # 四级：也可用中文序号层级的四级写法（左右括号都要在）
        match = re.match(r"^[（(](\d+)[.．、]?[）)][ \u3000]*(.*)$", stripped)
        if match:
            return finish(4, match.group(1), match.group(2).strip(),
                          sentence=True, capped=True)
        return None

    # 中文序号层级
    match = re.match(rf"^([{CN_NUMERAL}]+)、[ \u3000]*(.*)$", stripped)
    if match and _cn_number_ok(match.group(1)):
        return finish(1, match.group(1), match.group(2).strip(),
                      sentence=True, capped=False)
    # 二级：「括号括中文数字」——**左右括号都要在**（用户 2026-09-20 口径：
    # “不必成对使用”指两边可各用半角/全角，不是可以缺一边）。
    match = re.match(rf"^[（(]([{CN_NUMERAL}]+)[）)][ \u3000]*(.*)$", stripped)
    if match and _cn_number_ok(match.group(1)):
        return finish(2, match.group(1), match.group(2).strip(),
                      sentence=True, capped=False)
    # 三、四级：表原写“不以标点结尾”，用户 2026-09-20 裁决系漏写、已改回
    # “不以句末标点结尾”——全层级统一用句末标点集。另：**点号后不接数字**
    # （用户 2026-09-20 裁决 `1.1`/`1.4` 这类不认，与引擎一致）。
    match = re.match(r"^(\d+)[.．、](?!\d)[ \u3000]*(.*)$", stripped)
    if match:
        return finish(3, match.group(1), match.group(2).strip(),
                      sentence=True, capped=True)
    match = re.match(r"^[（(](\d+)[.．、]?[）)][ \u3000]*(.*)$", stripped)
    if match:
        return finish(4, match.group(1), match.group(2).strip(),
                      sentence=True, capped=True)
    return None


def long_suspect(text: str, scheme: str) -> tuple[int, str, str, str] | None:
    """被 60 字上限刷掉的标题形状（给“疑似标题”登记）。

    形状与标点都用同一套判据，只是长度不参与否决；再按表收窄到“受上限管”的
    级别（中文层级三、四级；阿拉伯数字层级一至四级）。
    """
    if len(squeezed(text.strip())) <= TITLE_MAX:
        return None
    found = heading_match(text, scheme, ignore_length=True)
    if found is None:
        return None
    return found if (scheme == "dotted" or found[0] >= 3) else None


# ---------------------------------------------------------------- 题注（S1）

CAPTION_CN_RE = re.compile(r"^([图表])[ 　]*(\d+(?:[.\-]\d+)*)")
CAPTION_EN_RE = re.compile(r"^(fig(?:ure)?|table)\.?[ \u3000]+(\d+(?:[.\-]\d+)*)",
                           re.IGNORECASE)
CAPTION_EN_SOFT_RE = re.compile(r"^(fig(?:ure)?|table)\.?[ \u3000]", re.IGNORECASE)
CAPTION_CN_SOFT_RE = re.compile(r"^[图表][ 　]*")
CONTINUATION_RE = re.compile(r"[（(][ 　]*续[图表]?表?[ 　]*[）)]")
CONTINUATION_HEAD_RE = re.compile(r"^续[图表][ 　]*\d")


def seq_labels(para: Para) -> list[str]:
    """段落里所有 SEQ 域用的标签（域指令形如 `SEQ 图 \\* ARABIC`）。"""
    found: list[str] = []
    for instruction in para.fields:
        for match in re.finditer(r"SEQ\s+([^\s\\]+)", instruction or "", re.IGNORECASE):
            found.append(match.group(1).strip('"“”'))
    return found


def seq_kind(labels: list[str]) -> str | None:
    """SEQ 标签 → 图/表；精确匹配（图表、公式一律不算）。"""
    for label in labels:
        low = label.strip().lower()
        if low in ("图", "fig", "figure"):
            return "图"
        if low in ("表", "table"):
            return "表"
    return None


def is_caption_style(para: Para) -> bool:
    return str(para.style_name).strip().lower() in ("caption", "题注")


def is_continuation(text: str, title_part: str) -> bool:
    """续图/续表三档（表 S1）：括号续注（位置不限）、以“续图/续表+数字”开头、
    或数字之后紧跟“续图/续表”。"""
    if CONTINUATION_RE.search(text):
        return True
    if CONTINUATION_HEAD_RE.match(text.strip()):
        return True
    return title_part.strip().startswith(("续图", "续表"))


# ---------------------------------------------------------------- 目录域（S3）

_TOC_INSTR_RE = re.compile(r"^\s*TOC(\s|$)", re.IGNORECASE)
_LIST_FIELD_KIND_RE = re.compile(r"\\c\s+\"?([^\"\s\\]+)\"?")
_LIST_FIELD_STYLE_RE = re.compile(r'\\t\s+"([^"]*)"')
_STYLE_NAME_KINDS = {"图题": "图", "表题": "表"}


def has_toc_instr(para: Para) -> bool:
    return any(_TOC_INSTR_RE.match(code or "") for code in para.fields)


def list_kinds_in_instr(code: str) -> set[str]:
    kinds = {match.group(1).strip() for match in _LIST_FIELD_KIND_RE.finditer(code)}
    for match in _LIST_FIELD_STYLE_RE.finditer(code):
        for item in match.group(1).split(","):
            kind = _STYLE_NAME_KINDS.get(item.strip())
            if kind is not None:
                kinds.add(kind)
    return kinds


def is_entry_style(para: Para) -> bool:
    name = str(para.style_name).strip().lower()
    return name.startswith("toc ") or name == "table of figures"


def toc_domain(path: pathlib.Path) -> tuple[set[int], set[int]]:
    r"""扫 word/document.xml，返回 `(域指令段, 条目段)`。

    * 域指令段：段落里至少有一条以 TOC 开头的域指令；
    * 条目段：**域指令所在段之后、域结束之前**的段落（表 S3 的字面）。
      嵌套域不影响（条目的 PAGEREF 自开自合，净深度为零）。
    """
    with zipfile.ZipFile(path) as package:
        root = etree.fromstring(package.read("word/document.xml"))
    body = root.find(qn("w:body"))
    if body is None:
        return set(), set()
    instructions: set[int] = set()
    entries: set[int] = set()
    depth = 0
    open_index: int | None = None
    paragraph = -1
    for element in body:
        if not isinstance(element.tag, str) or element.tag != qn("w:p"):
            continue
        paragraph += 1
        opens = closes = 0
        codes: list[str] = []
        for node in element.iter():
            if not isinstance(node.tag, str):
                continue
            name = local_name(node)
            if name == "fldChar":
                kind = attrs_of(node).get("fldCharType", "")
                if kind == "begin":
                    opens += 1
                elif kind == "end":
                    closes += 1
            elif name == "instrText" and node.text:
                codes.append(node.text)
            elif name == "fldSimple":
                codes.append(attrs_of(node).get("instr", ""))
        here_toc = any(_TOC_INSTR_RE.match(code or "") for code in codes)
        if here_toc:
            instructions.add(paragraph)
        # “域结束之前”：本段自己把域收回去（处理完这一段深度就归零）⇒ 不算条目。
        # ⚠️ 要用 depth + opens - closes（条目段自带的嵌套域是先 begin 后 end，
        # 只看 closes 会把每个条目段都误判成收尾段——2026-09-20 在民总目录上实拍）。
        carries_end = bool(closes and depth > 0 and depth + opens - closes <= 0)
        if open_index is not None and paragraph > open_index and depth > 0 \
                and not carries_end:
            entries.add(paragraph)
        depth += opens - closes
        if depth <= 0:
            depth = 0
            open_index = None
        if here_toc and opens:
            open_index = paragraph
    return instructions, entries


# ---------------------------------------------------------------- 自动编号反算
# 表前总述：识别时先把 Word 自动编号按编号定义反算成实际显示的编号文字，再依
# 表判断。本实现自己读 numbering.xml（不借引擎的编号模块）。

_CN_SIMPLE = "零一二三四五六七八九"


def _chinese(value: int) -> str:
    if value < 0:
        return str(value)
    if value < 10:
        return _CN_SIMPLE[value]
    if value < 20:
        return "十" + (_CN_SIMPLE[value - 10] if value > 10 else "")
    if value < 100:
        tens, ones = divmod(value, 10)
        return _CN_SIMPLE[tens] + "十" + (_CN_SIMPLE[ones] if ones else "")
    if value < 1000:
        hundreds, rest = divmod(value, 100)
        head = _CN_SIMPLE[hundreds] + "百"
        if not rest:
            return head
        return head + ("零" if rest < 10 else "") + _chinese(rest)
    thousands, rest = divmod(value, 1000)
    head = _CN_SIMPLE[thousands] + "千"
    if not rest:
        return head
    return head + ("零" if rest < 100 else "") + _chinese(rest)


def _roman(value: int) -> str:
    if value <= 0:
        return str(value)
    table = ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"),
             (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"),
             (5, "V"), (4, "IV"), (1, "I"))
    out = []
    for amount, text in table:
        while value >= amount:
            out.append(text)
            value -= amount
    return "".join(out)


def _letter(value: int) -> str:
    if value <= 0:
        return str(value)
    out = ""
    while value > 0:
        value, remainder = divmod(value - 1, 26)
        out = chr(ord("A") + remainder) + out
    return out


def format_number(value: int, number_format: str) -> str:
    """按编号格式把计数写成看得见的文字（表没列举格式，常见的一网打尽）。"""
    name = (number_format or "decimal").strip()
    if name == "decimal":
        return str(value)
    if name in ("chineseCounting", "chineseCountingThousand", "japaneseCounting"):
        return _chinese(value)
    if name == "lowerRoman":
        return _roman(value).lower()
    if name == "upperRoman":
        return _roman(value)
    if name == "lowerLetter":
        return _letter(value).lower()
    if name == "upperLetter":
        return _letter(value)
    if name == "decimalEnclosedCircle":
        circled = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
        return circled[value - 1] if 1 <= value <= len(circled) else str(value)
    if name in ("decimalEnclosedParen", "decimalEnclosedFullstop"):
        return f"({value})" if name.endswith("Paren") else f"{value}."
    return str(value)


def read_numbering(path: pathlib.Path):
    """读 word/numbering.xml → (每级定义, numId→abstractNumId, 局部覆盖)。"""
    name = "word/numbering.xml"
    with zipfile.ZipFile(path) as package:
        if name not in package.namelist():
            return {}, {}, {}
        root = etree.fromstring(package.read(name))
    abstracts: dict[str, dict[int, dict]] = {}
    numbers: dict[str, str] = {}
    overrides: dict[str, dict[int, dict]] = {}
    for abstract in root.findall(qn("w:abstractNum")):
        abstract_id = abstract.get(qn("w:abstractNumId")) or ""
        levels: dict[int, dict] = {}
        for level in abstract.findall(qn("w:lvl")):
            index = int(level.get(qn("w:ilvl")) or 0)
            levels[index] = _level_info(level)
        abstracts[abstract_id] = levels
    for number in root.findall(qn("w:num")):
        num_id = number.get(qn("w:numId")) or ""
        holder = number.find(qn("w:abstractNumId"))
        if holder is not None:
            numbers[num_id] = holder.get(qn("w:val")) or ""
        local: dict[int, dict] = {}
        for override in number.findall(qn("w:lvlOverride")):
            index = int(override.get(qn("w:ilvl")) or 0)
            level = override.find(qn("w:lvl"))
            if level is not None:
                local[index] = _level_info(level)
        if local:
            overrides[num_id] = local
    return abstracts, numbers, overrides


def _level_info(level) -> dict:
    start = level.find(qn("w:start"))
    numfmt = level.find(qn("w:numFmt"))
    text = level.find(qn("w:lvlText"))
    return {
        "start": int(attrs_of(start).get("val") or 1) if start is not None else 1,
        "numFmt": attrs_of(numfmt).get("val", "decimal") if numfmt is not None else "decimal",
        "lvlText": attrs_of(text).get("val", "%1") if text is not None else "%1",
    }


def reverse_numbering(paras, path: pathlib.Path) -> dict[int, str]:
    """按文档顺序反算每个"用了自动编号"的段落实际显示的编号文字。"""
    abstracts, numbers, overrides = read_numbering(path)
    counters: dict[str, dict[int, int]] = {}
    visible: dict[int, str] = {}
    for para in paras:
        if para.index is None or para.num_id is None:
            continue
        levels = dict(abstracts.get(numbers.get(para.num_id, ""), {}))
        levels.update(overrides.get(para.num_id, {}))
        level = para.num_level
        info = levels.get(level)
        if info is None:
            continue
        state = counters.setdefault(para.num_id, {})
        state[level] = state.get(level, info["start"] - 1) + 1
        for deeper in [key for key in state if key > level]:
            del state[deeper]
        text = info["lvlText"]
        for slot in range(1, 10):
            placeholder = f"%{slot}"
            if placeholder not in text:
                continue
            value = state.get(slot - 1, levels.get(slot - 1, {}).get("start", 1))
            fmt = levels.get(slot - 1, {}).get("numFmt", "decimal")
            text = text.replace(placeholder, format_number(value, fmt))
        visible[para.index] = text
    return visible


def toc_field_paragraphs(path: pathlib.Path) -> set[int]:
    """兼容旧调用：落在目录族域区间里的条目段（= `toc_domain()` 的第二项）。"""
    _instructions, entries = toc_domain(path)
    return entries


# ---------------------------------------------------------------- 期望结果


@dataclass
class Expectation:
    roles: dict[int, str] = field(default_factory=dict)
    rules: dict[int, str] = field(default_factory=dict)
    levels: dict[int, int] = field(default_factory=dict)
    raws: dict[int, str] = field(default_factory=dict)
    titles: dict[int, str] = field(default_factory=dict)
    undecided: dict[int, str] = field(default_factory=dict)
    captions: dict[int, str] = field(default_factory=dict)
    english_captions: set[int] = field(default_factory=set)
    auto_numbered: set[int] = field(default_factory=set)
    english_title: int | None = None
    english_subtitle: int | None = None
    skipped: dict[int, str] = field(default_factory=dict)
    duplicates: dict[int, str] = field(default_factory=dict)      # 未采信的重复前置标签
    long_suspects: set[int] = field(default_factory=set)          # 超长疑似标题
    appendix_entries: set[int] = field(default_factory=set)
    inline_keywords: set[int] = field(default_factory=set)        # 关键词同段写法的段落

    def note(self, index: int, role: str, rule: str) -> None:
        self.roles[index] = role
        self.rules[index] = rule


def _placeholder_name(text: str) -> str | None:
    """工具自插占位语的名字（样子按产品自己的定义——产品事实，不是识别判据）。"""
    if is_title_placeholder(text):
        return "题目占位语"
    if is_subtitle_placeholder(text):
        return "副标题占位语"
    if is_english_title_placeholder(text):
        return "英文题目占位语"
    if is_english_subtitle_placeholder(text):
        return "英文副标题占位语"
    if is_placeholder(text):
        return "占位语"
    return None


def _previous_nonempty(paras: list[Para], index: int, start: int) -> int | None:
    for previous in range(index - 1, start - 1, -1):
        if paras[previous].text.strip():
            return previous
    return None


# ---------------------------------------------------------------- 识别本体


def expect(shape: DocShape, *, cover_end: int | None = None,
           generated_entries: set[int] | None = None,
           visible_numbers: dict[int, str] | None = None,
           scheme: str = "chinese") -> Expectation:
    """按 3.6.2 识别表算每一段的期望身份。

    `generated_entries`：目录/清单域区间里的条目段（`toc_field_paragraphs()` 的
    产物；表 S3 说它们不参与任何识别）。
    `visible_numbers`：`reverse_numbering()` 反算出的自动编号（表前总述）。
    `scheme`：原文标题序号方案，"chinese"（默认）或 "dotted"。

    执行顺序照表：前向识别器 → 后向识别器 → 题目与副标题（“前向扫描器停止后”
    按已认出的部件算）→ 补充识别器（S2/S1/S4）→ 归属划分。
    """
    start = 0 if cover_end is None else cover_end + 1
    result = Expectation()
    numbers = visible_numbers or {}
    arabic = scheme == "dotted"

    paragraphs = [(para.index, para) for para in shape.paras
                  if para.index is not None and para.index >= start]

    def visible(index: int, para: Para) -> str:
        body = para.text.strip()
        prefix = numbers.get(index, "")
        return (prefix + body) if prefix else body

    # 不参与识别的两类：域生成的条目（S3）、工具自插占位语（表里没这个对象；
    # 只有 T3 明说“英文题目占位文字会被认出”，那一处单独处理）。
    domain_entries: set[int] = set(generated_entries or set())
    placeholders: dict[int, str] = {}
    for index, para in paragraphs:
        name = _placeholder_name(para.text)
        if name is not None:
            placeholders[index] = name
    for index in sorted(domain_entries):
        result.skipped[index] = "目录/清单域生成的条目（S3）：不参与任何识别"
    for index in sorted(placeholders):
        result.skipped[index] = f"{placeholders[index]}（产品自插，表里没有这个对象）"

    def invisible(index: int) -> bool:
        return index in domain_entries or index in placeholders

    claimed: set[int] = set()

    # ---------------------------------------------------------- ① 前向识别器
    # 从前往后扫；四个标签各自只认第一次；撞到停止清单里的任何身份就停。
    forward_range_end = len(shape.paras)
    forward_found: dict[str, int] = {}
    for index, para in paragraphs:
        if invisible(index):
            continue
        text = visible(index, para)
        if not text:
            continue
        # 停止清单：引言家族 / 结论、参考文献、附录、致谢 / 目录与三清单 /
        # 当前方案下的标题（都按后向识别器部分的判据认，只看形状）。
        if (match_intro(text, arabic=arabic) is not None
                or match_conclusion(text, arabic=arabic)
                or any(match_cjk_label(text, word) for word in B_LABELS)
                or any(match_cjk_label(text, word) for word in TOC_LABELS)
                or heading_match(text, scheme) is not None):
            forward_range_end = index
            break
        for position, part in enumerate(F_LABELS, start=1):
            if part in forward_found:
                continue
            if match_f_label(text, part):
                forward_found[part] = index
                result.note(index, f"{part}标签", f"F{position}")
                if match_inline_keyword(text, part):
                    result.inline_keywords.add(index)
                claimed.add(index)
    # 同一标签在**前置区**内再出现的：识别到但不采信（3.7“重复的部件”），登记位置。
    for index, para in paragraphs:
        if index >= forward_range_end:
            break
        if index in claimed or invisible(index):
            continue
        text = visible(index, para)
        for part in F_LABELS:
            if part in forward_found and match_f_label(text, part):
                result.duplicates[index] = part

    # ---------------------------------------------------------- ② 后向识别器
    # 从文档尾向前；四个停止条件；后置标签回收它后面的 pending 标题。
    back_found: dict[str, int] = {}
    pending: dict[int, tuple[int, str, str, str]] = {}   # 标题候选（未定身份）
    cjk_seen: set[str] = set()
    streak_index: int | None = None
    streak_len = 0
    backward_stop = start
    i = len(shape.paras) - 1
    while i >= start:
        para = shape.paras[i]
        text = visible(i, para)
        # 停止②：前向已认出的部件 / 目录三清单标签 / 带 TOC 域的段 / toc 条目样式段
        if (i in claimed
                or has_toc_instr(para)
                or is_entry_style(para)
                or (text and any(match_cjk_label(text, word) for word in TOC_LABELS))):
            backward_stop = i + 1
            break
        if invisible(i):
            i -= 1
            continue
        if not text:
            i -= 1
            continue
        # 停止①：引言家族
        if match_intro(text, arabic=arabic) is not None:
            result.note(i, "引言", "H5")
            claimed.add(i)
            backward_stop = i
            break
        # 后置标签（参考文献/附录/致谢）与“结论”：只登记第一次碰到的；
        # 重复出现不再登记，但同样把 pending 退回（表：每识别到一个后置标签…）。
        back_word = next((word for word in B_LABELS if match_cjk_label(text, word)), None)
        if back_word is not None or match_conclusion(text, arabic=arabic):
            name = back_word or CONCLUSION_WORD
            if name not in back_found:
                back_found[name] = i
                rule = {"参考文献": "B1", "附录": "B2", "致谢": "B3"}.get(name, "H6")
                result.note(i, name, rule)
                claimed.add(i)
            for index in list(pending):
                result.note(index, "正文", "G")       # 一律退回为普通正文
                pending.pop(index)
            cjk_seen.clear()
            streak_index, streak_len = None, 0
            i -= 1
            continue
        # 标题候选
        heading = heading_match(text, scheme)
        if heading is not None:
            level, rule_id, raw, title_text = heading
            cjk = "".join(ch for ch in text if "一" <= ch <= "\u9fff")
            duplicate = bool(cjk) and cjk in cjk_seen
            if duplicate and streak_index == i + 1:
                streak_len += 1
            elif duplicate:
                streak_len = 1
            else:
                streak_len = 0
            streak_index = i
            if duplicate and streak_len >= 2:
                # 停止④：连续两段都撞上已扫描过的标题汉字部分 ⇒ 退回并停止
                for offset in (i, i + 1):
                    if offset in pending:
                        result.note(offset, "正文", "G")
                        pending.pop(offset)
                backward_stop = i
                break
            pending[i] = (level, rule_id, raw, title_text)
            if cjk:
                cjk_seen.add(cjk)
            i -= 1
            continue
        # 超长疑似（表：不识别为标题、按正文处理，同时登记给“疑似标题”）
        if long_suspect(text, scheme) is not None:
            result.long_suspects.add(i)
        streak_index, streak_len = None, 0
        i -= 1

    # pending 里活下来的就是正文标题（表：后向识别器记录每一个有效标题）
    for index, (level, rule_id, raw, title_text) in pending.items():
        result.note(index, f"{['', '一', '二', '三', '四'][level]}级标题", rule_id)
        result.levels[index] = level
        result.raws[index] = raw
        result.titles[index] = title_text
        if index in numbers:
            result.auto_numbered.add(index)
        claimed.add(index)

    # ---------------------------------------------------------- T3/T4 英文题名
    # 2026-09-20 用户裁决：英文题目算“已识别标题”——它排在中文题目之前时，
    # 后面的中文题目/副标题不再认（引擎行为即口径）。所以 T3/T4 比 T1/T2 先算。
    # T3 判据：前置区内第一个不含汉字、不以句末标点结尾、≥5 拉丁字母、≤120 字的段落。
    english_title = None
    for index, para in paragraphs:
        if index >= forward_range_end:
            break
        if index in claimed or index in domain_entries:
            continue
        text = visible(index, para)
        if not text:
            continue
        if placeholders.get(index) == "英文题目占位语":
            english_title = index
            break
        if index in placeholders:
            continue
        if any(match_cjk_label(text, word) for word in TOC_LABELS):
            continue
        if not is_all_english(text) or ends_sentence(text):
            continue
        if len(LATIN_RE.findall(text)) < EN_TITLE_MIN_LETTERS:
            continue
        if len(text) > EN_TITLE_MAX_LEN:
            continue
        english_title = index
        break
    if english_title is not None:
        result.english_title = english_title
        result.note(english_title, "英文题目", "T3")
        claimed.add(english_title)
        follow = english_title + 1
        if follow < len(shape.paras) and follow not in claimed and not invisible(follow):
            para = shape.paras[follow]
            text = visible(follow, para)
            if (text and not ends_sentence(text) and is_all_english(text)
                    and len(LATIN_RE.findall(text)) >= 1):
                result.english_subtitle = follow
                result.note(follow, "英文副标题", "T4")
                claimed.add(follow)

    # ---------------------------------------------------------- T：题目与副标题
    # T1/T2 的边界 = 所有已识别标签和标题里最靠前的位置（含目录/三清单标签——
    # 它们虽然在补充识别器里才登记，但“已识别标签”按结果算；2026-09-20 起
    # 英文题目/英文副标题也算在内）。
    boundary_positions = [index for index, role in result.roles.items()
                          if role.endswith("标签")
                          or role in ("引言", "结论", "参考文献", "附录", "致谢")
                          or role in ("英文题目", "英文副标题")
                          or role in ("一级标题", "二级标题", "三级标题", "四级标题")]
    for index, para in paragraphs:
        text = visible(index, para)
        if text and any(match_cjk_label(text, word) for word in TOC_LABELS):
            boundary_positions.append(index)
    boundary = min(boundary_positions, default=len(shape.paras))

    title_index: int | None = None
    for index, para in paragraphs:
        if index >= boundary:
            break
        if invisible(index) or index in claimed:
            continue
        text = visible(index, para)
        if not text or ends_sentence(text) or is_all_english(text):
            continue
        if is_entry_style(para) or is_caption_style(para):
            continue
        title_index = index
        result.note(index, "题目", "T1")
        claimed.add(index)
        break
    subtitle_index: int | None = None
    if title_index is not None:
        follow = title_index + 1
        if (follow < boundary and follow not in claimed and not invisible(follow)):
            para = shape.paras[follow]
            text = visible(follow, para)
            # 用户 2026-09-20 裁决：中文副标题**不能是全英文的**（全英文行
            # 只能被 T3 收走）。
            if text and not ends_sentence(text) and not is_all_english(text):
                subtitle_index = follow
                result.note(follow, "中文副标题", "T2")
                claimed.add(follow)

    # ---------------------------------------------------------- ③ 补充识别器
    # S2：目录与三清单标签——夹缝 =（前向已认出的部件之后、后向停止位置之前）；
    # 每种标签只认第一次出现。
    forward_items = list(forward_found.values())
    forward_items += [index for index in (title_index, subtitle_index, english_title,
                                          result.english_subtitle) if index is not None]
    gap_start = max(forward_items, default=start - 1) + 1
    seen_label_parts = set(result.roles.values())
    for index, para in paragraphs:
        if index < gap_start or index >= backward_stop:
            continue
        if invisible(index):
            continue
        text = visible(index, para)
        if not text:
            continue
        for word in TOC_LABELS:
            if f"{word}标签" in seen_label_parts:
                continue
            if match_cjk_label(text, word):
                result.note(index, f"{word}标签", "S2")
                claimed.add(index)
                seen_label_parts.add(f"{word}标签")
                break

    # S3：带 TOC 域的段落登记（域指令段不参与识别；条目段在建表时已排除）
    for index, para in paragraphs:
        if index in domain_entries or index in placeholders:
            continue
        if has_toc_instr(para):
            result.skipped[index] = "带 TOC 域的段落（S3）：不参与识别（域指令段）"

    # S1：题注（跳过标题/标签/附录题名/域条目/占位语）
    chinese_caption_at: set[int] = set()
    for index, para in paragraphs:
        if invisible(index) or index in claimed or para.in_table:
            continue
        text = visible(index, para)
        if not text:
            continue
        chinese = CAPTION_CN_RE.match(text)
        english = None if chinese is not None else CAPTION_EN_RE.match(text)
        strong = bool(seq_labels(para)) or is_caption_style(para)
        soft_en = bool(CAPTION_EN_SOFT_RE.match(text))
        # “文字符合上述前缀”：编号在域里、可见文字上看不见数字的那种，
        # 表里明说“也允许使用题注功能插入” ⇒ 强信号 + 以图/表/Fig/Table 起头即算。
        prefix_ok = (chinese is not None or english is not None or soft_en
                     or (strong and CAPTION_CN_SOFT_RE.match(text) is not None))
        if not (strong and prefix_ok) and ends_sentence(text):
            continue
        if len(squeezed(text)) > CAPTION_MAX:
            continue
        kind: str | None = None
        title_part = ""
        if chinese is not None:
            kind = chinese.group(1)
            title_part = text[chinese.end():]
        elif english is not None:
            if not strong:
                previous = _previous_nonempty(shape.paras, index, start)
                if previous is None or previous not in chinese_caption_at:
                    continue
            kind = "图" if english.group(1).lower().startswith("fig") else "表"
            title_part = text[english.end():]
        elif strong:
            kind = seq_kind(seq_labels(para))
        if kind is None:
            continue
        is_english = english is not None or (chinese is None and soft_en)
        role = f"{kind}题"
        if is_english:
            role += "（英文对照）"
        elif is_continuation(text, title_part):
            role += "（续行）"
        result.note(index, role, "S1")
        result.captions[index] = kind
        if is_english:
            result.english_captions.add(index)
        else:
            chinese_caption_at.add(index)
        claimed.add(index)

    # S4：附录序号题名（附录标签之后、下一个 参考文献/结论/致谢 标签之前）
    appendix_at = next((index for index, role in result.roles.items()
                        if role == "附录"), None)
    if appendix_at is not None:
        stop_at = next((index for index in sorted(claimed)
                        if index > appendix_at
                        and result.roles.get(index) in ("参考文献", "结论", "致谢")),
                       len(shape.paras))
        for index, para in paragraphs:
            if index <= appendix_at or index >= stop_at:
                continue
            if invisible(index) or index in claimed:
                continue
            text = visible(index, para)
            if not text or ends_sentence(text):
                continue
            if not re.match(r"^附\s*录\s*[A-Za-z]", text):
                continue
            if len(squeezed(text)) > APPENDIX_MAX:
                continue
            # 只登记、不给“角色”：引擎的识别层里它不是一个部件身份，套样式
            # （S20）由加工侧自己做——写成角色反而会让期望侧成品丢掉它。
            result.appendix_entries.add(index)
            claimed.add(index)

    # ---------------------------------------------------------- 未判定登记
    for index, para in paragraphs:
        if index in result.roles or index in result.skipped or index in result.undecided:
            continue
        if para.num_id is not None and index not in numbers:
            result.undecided[index] = "用了 Word 自动编号，但反算不出可见编号 ⇒ 未判定"

    # ---------------------------------------------------------- G：归属划分
    owner: str | None = None
    for index, para in paragraphs:
        if index in domain_entries or index in placeholders:
            continue                      # 域条目与占位语：不算身份、也不打断归属
        if index in result.undecided or index in result.appendix_entries:
            continue
        role = result.roles.get(index)
        if role is not None:
            part = role[:-2] if role.endswith("标签") else role
            owner = part if part in OWNER_PARTS else None
            if index in result.inline_keywords:
                owner = None              # 同段写法：自己就是整个部件，不归属下方段落
            continue
        if owner is not None:
            if owner in ("摘要", "关键词", "英文摘要", "英文关键词"):
                result.note(index, f"{owner}正文", "G")
            else:
                result.note(index, "正文", "G")   # 锚点部件的正文按普通正文报
        elif para.text.strip():
            result.note(index, "正文", "G")
        # 空段：不给身份（表里要“非空”的地方都按非空判）
    return result


# ---------------------------------------------------------------- 遗留路径
# 期望对拍现在走 `aat_engine_api.expect_index()`（把身份写回索引）。下面这个
# “换算成 ManualStructure”的老路留着备查，不参与现役流程。


def to_manual_structure(shape: DocShape, expectation: Expectation, *, cover_end=None):
    """把期望身份换算成 `ManualStructure`（旧路径，备查）。"""
    from aat_engine_api import (CaptionMatch, FrontMatterLabel, HeadingMatch,
                                ManualStructure, TitleMatch)

    role_to_label = {f"{part}标签": literal for part, literal in F_LABELS.items()}
    role_to_level = {"一级标题": 1, "二级标题": 2, "三级标题": 3, "四级标题": 4}
    matches = [None] * len(shape.paras)
    front_matter: list = []
    captions: list = []
    title_index = None
    subtitle_index = None
    english_title_index = None
    english_subtitle_index = None
    for index, role in sorted(expectation.roles.items()):
        para = shape.paras[index]
        if role in role_to_level:
            level = role_to_level[role]
            matches[index] = HeadingMatch(
                paragraph_index=index, level=level,
                number=expectation.levels.get(index),
                marker=expectation.raws.get(index, ""),
                title=expectation.titles.get(index, ""),
                auto_numbered=index in expectation.auto_numbered)
        elif role in ANCHORS:
            matches[index] = HeadingMatch(paragraph_index=index, level=1, number=None,
                                          marker=role, title="", role=role)
        elif role == "题目":
            title_index = index
        elif role == "中文副标题":
            subtitle_index = index
        elif role == "英文题目":
            english_title_index = index
        elif role == "英文副标题":
            english_subtitle_index = index
        elif role in role_to_label:
            front_matter.append(FrontMatterLabel(paragraph_index=index,
                                                 label=role_to_label[role]))
        elif role.startswith("图题") or role.startswith("表题"):
            captions.append(CaptionMatch(paragraph_index=index, kind=role[0],
                                         text=para.text.strip(),
                                         is_english=index in expectation.english_captions,
                                         is_continuation="（续行）" in role))
    title = (TitleMatch(title_paragraph_index=title_index,
                        subtitle_paragraph_index=subtitle_index)
             if title_index is not None else None)
    front_matter.sort(key=lambda item: item.paragraph_index)
    captions.sort(key=lambda item: item.paragraph_index)
    fields = {spec.name for spec in dataclasses.fields(ManualStructure)}
    values = {
        "matches": matches, "demoted": [], "front_matter": front_matter,
        "title": title, "captions": captions,
        "english_title": english_title_index,
        "english_subtitle": english_subtitle_index, "quote_paragraphs": [],
    }
    return ManualStructure(**{k: v for k, v in values.items() if k in fields})


# ---------------------------------------------------------------- 报告


def render_expectation(shape: DocShape, expectation: Expectation, *,
                       cover_end=None, scheme: str = "chinese") -> str:
    lines = [f"夹具：{aat_scan_safe(shape.path)}",
             f"期望身份（依据：操作逻辑 3.6.2 识别表的字面判据；序号方案："
             f"{'阿拉伯数字' if scheme == 'dotted' else '中文'}）",
             "=" * 74]
    counts: dict[str, int] = {}
    for index, para in enumerate(shape.paras):
        role = expectation.roles.get(index)
        if role is None and index not in expectation.undecided:
            continue
        counts[role or "（未判定）"] = counts.get(role or "（未判定）", 0) + 1
        rule = expectation.rules.get(index, "")
        note = expectation.undecided.get(index, "")
        mark = f"{role}" if role else "⚠️ 未判定"
        lines.append(f"[{index:>3}] {mark:<14} {rule:<5} {para.text.strip()[:34]!r}")
        if note:
            lines.append(f"      └ {note}")
    lines.append("")
    lines.append("按身份汇总：" + "、".join(f"{k} {v}" for k, v in sorted(counts.items())))
    if expectation.skipped:
        indexes = sorted(expectation.skipped)
        lines.append("")
        lines.append(f"不参与识别 {len(indexes)} 段："
                     f"段 {indexes[0]}~{indexes[-1]}"
                     f"（例：{expectation.skipped[indexes[0]]}）")
    if expectation.duplicates:
        lines.append("")
        lines.append("未采信的重复标签：" + "、".join(
            f"段{index}（{part}）" for index, part in sorted(expectation.duplicates.items())))
    if expectation.long_suspects:
        lines.append("")
        lines.append("超长疑似标题（按表应报“疑似标题”）：段 "
                     + "、".join(str(index) for index in sorted(expectation.long_suspects)))
    if expectation.appendix_entries:
        lines.append("")
        lines.append("附录序号题名（S4）：段 "
                     + "、".join(str(index) for index in sorted(expectation.appendix_entries)))
    if expectation.undecided:
        lines.append("")
        lines.append(f"未判定 {len(expectation.undecided)} 段（表里没写清，见下）")
    lines.append("")
    lines.append("=" * 74)
    lines.append("本预估用到的假设（表里没写、按常理推断的）")
    lines.append("=" * 74)
    for item in ASSUMPTIONS:
        lines.append(f"  · {item}")
    return "\n".join(lines)


def aat_scan_safe(path) -> str:
    try:
        return str(pathlib.Path(path).resolve().relative_to(ROOT))
    except ValueError:
        return str(path)


def prepare(source: pathlib.Path, work_dir: pathlib.Path) -> pathlib.Path:
    """把源文件变成管线看到的那份（拆内容控件），供预估读取。

    ⚠️ 拆包口径必须与 `pipeline.run()` 一模一样（`keep_content` 按“有没有真
    目录/清单域”算），否则两侧段号会整体错位。这一步只做“让两份文档看得见
    同一批段落”，不掺任何识别判据。
    """
    from docx import Document
    from essay_tool.engine.cleanup import unwrap_section_break_sdt
    from essay_tool.engine.doc_index import body_toc_field_parts
    from essay_tool.engine.loader import load_as_working_copy

    shutil.rmtree(work_dir, ignore_errors=True)
    path = load_as_working_copy(source, work_dir)
    document = Document(str(path))
    unwrap_section_break_sdt(document, keep_content=bool(body_toc_field_parts(document)))
    document.save(str(path))
    return path


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="按新识别表算期望识别结果，并与自动识别差分")
    parser.add_argument("source", type=pathlib.Path)
    parser.add_argument("--preset", type=pathlib.Path, default=PRESET)
    parser.add_argument("--cover-end", type=int, default=None)
    parser.add_argument("--scheme", default="chinese", choices=("chinese", "dotted"))
    parser.add_argument("--no-diff", action="store_true", help="只出期望，不跑管线")
    args = parser.parse_args(argv)
    if not args.source.exists():
        raise SystemExit(f"找不到夹具：{args.source}")

    OUT.mkdir(parents=True, exist_ok=True)
    stem = args.source.stem
    prepared = prepare(args.source, OUT / "_work" / stem)
    shape = read_docx(prepared)
    visible = reverse_numbering(shape.paras, prepared)
    expectation = expect(shape, cover_end=args.cover_end,
                         generated_entries=toc_field_paragraphs(prepared),
                         visible_numbers=visible, scheme=args.scheme)
    if visible:
        print(f"反算出的自动编号：{len(visible)} 段 "
              f"（例：{sorted(visible.items())[:3]}）")

    text = render_expectation(shape, expectation, cover_end=args.cover_end,
                              scheme=args.scheme)
    (OUT / f"{stem}-期望.txt").write_text(text, encoding="utf-8")
    print(text)
    print()

    if args.no_diff:
        return 0

    from essay_tool.engine.alerts import AlertCollector
    auto_dir = OUT / "_auto" / stem
    want_dir = OUT / "_want" / stem
    shutil.rmtree(auto_dir, ignore_errors=True)
    shutil.rmtree(want_dir, ignore_errors=True)
    auto_product = pipeline.run(args.source, None, auto_dir,
                                preset_config=load_preset(args.preset),
                                alerts=AlertCollector(), cover_end=args.cover_end)
    want_product = pipeline.run(args.source, None, want_dir,
                                preset_config=load_preset(args.preset),
                                alerts=AlertCollector(), cover_end=args.cover_end,
                                index=expect_index(prepared, expectation))
    outcome = aat_diff.compare_shapes(read_docx(auto_product), read_docx(want_product))
    report = aat_diff.render(outcome, max_items=60)
    header = ("自动识别的成品（左） vs 期望识别的成品（右）\n"
              "差异 = 识别结果与识别表不一致的后果\n" + "=" * 78 + "\n")
    (OUT / f"{stem}-分歧.txt").write_text(header + report, encoding="utf-8")
    print(header + report)
    print()
    print(f"产出：Adversarial Automatic Test/_out/期望/{stem}-期望.txt / {stem}-分歧.txt")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
