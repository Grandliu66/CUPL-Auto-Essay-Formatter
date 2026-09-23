"""标题识别：从段落判断是不是标题、是第几级。

判断顺序（优先级从高到低）：
    0. **大纲级别（`w:outlineLvl`）完全不采信**（用户 2026-09-15："大纲格式
       似乎完全冗余，完全不采信即可"）。早期这里有一条"段落或样式设了大纲
       级别就当候选标题"的分支，实测它与下面第 1 步**逐行等价**（同一段
       "第一章 绪论"：带大纲级别与不带级别的识别结果、告警条数完全一样），
       已经删掉。注意：**输出侧仍然要写大纲级别**（四个标题样式 + 附录标题
       都补 outlineLvl），那是给 Word 目录域认标题用的，跟识别无关。
    1. 命中基础正则——采信。正则不止认"一、"这种编号，也认"引言""结论"
       "参考文献""致谢""附录"这几个固定名称（跟编号标题是同一级别，都是
       正则可以识别的"要素"，处理方式一样）。
       编号/名称后面的文字末尾带标点的，按正文处理，静默处理、不弹提醒——
       标题不该带标点，这个信号足够确定，不需要问用户。
       三、四级标题明显偏长的，也按正文处理，但会弹提醒——这两级是"内层"
       级别，越往内层，"漏了换行、把正文粘到编号后面"这种手滑越常见，但
       长本身不代表错（标题可以合法地长），所以不静默处理，而是收进
       DemotedTitle / demoted 返回值，交给上层弹窗提醒用户"不规范的标题
       层级，已自动识别为正文"，并告诉用户想让它被识别为标题该怎么改。
       一、二级标题不检查长度——章节级标题写长了不算异常信号。
    2. 都不命中——正文。

基础正则：
    一、          (L1，中文数字+顿号)
    （一）        (L2，中文数字，全/半角括号)
    1.  / 1． / 1、 (L3，阿拉伯数字+句号/顿号，不认小数点后还有数字的情况，如"1.4"）
    （1）/ (1)     (L4，阿拉伯数字，全/半角括号)
    引言/结论/参考文献/致谢/附录 (固定名称，视为一级，字与字之间允许有空格；
    **“绪论”故意不收**，理由见下面 `_FIXED_HEADINGS` 附近的注释)

跳级（比如一级标题下面直接是三级，没有二级）不算问题，不提示。
check_sequence() 检查三类可疑情况：
    - 同级掉号：一级标题从"二"跳到"四"，中间缺"三"。
    - 跨级未归零：新的上级标题出现后，下一级的编号没有从1开始
      （多见于序号列表，比如复制粘贴时级别错位）。
    - 孤立列举：某一级标题在它所在的上级范围内，前后都没有同级的兄弟标题
      （不管数字是否连续，只看前后是否存在同级）。

另外，Word 自己的"自动编号列表"（多级列表/项目符号）生成的编号不会出现在
python-docx 读到的文字里——比如页面上看着是"(1) 比如这里用了半角括号"，
python-docx 读到的 paragraph.text 其实只有"比如这里用了半角括号"，编号是
Word 渲染时才画出来的，根本不在文字里。本模块会用 numbering.py 反查
numbering.xml，按段落出现顺序把这类编号实际算出来，拼回文字前面再去跑
上面那套正则——这样自动编号的标题跟手打编号的标题走的是同一套判断逻辑。
算不出编号的（比如列表定义本身缺失、或者是项目符号而不是数字编号），会
通过 find_unresolved_auto_numbered_paragraphs() 单独报出来（**主管线目前
不调它**，只有诊断工具 tools/check_detect.py 在用），不在
detect_headings() 的判断结果里悄悄猜。
"""

from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass, field

from docx.document import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

# 图题/表题识别 2026-09-19 搬到 `captions.py`（识别阶段一次写进索引）；
# 这里 re-export，老调用点（诊断脚本、AAT）照旧从 detect 取。
from essay_tool.engine.captions import (  # noqa: F401
    _CAPTION_RE,
    _EN_CAPTION_RE,
    CaptionMatch,
    _seq_label,
    detect_captions,
)
from essay_tool.engine.numbering import NumberingScheme, load_numbering_scheme
# 工具自己插的占位语：识别阶段要认得出"这行是我自己插的槽位"，不能当成
# 用户内容（否则成品回炉会反复补占位语，见 placeholders.py 的模块说明）。
from essay_tool.engine.placeholders import (
    is_english_subtitle_placeholder,
    is_english_title_placeholder,
    is_placeholder,
    is_subtitle_placeholder,
    is_title_placeholder,
)

_CN_DIGITS = {"零": 0, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
# 廿/卅（20/30）**不进编号字样**（用户 2026-09-17：“廿、卅永远不认”，只由告警表
# A2 报疑似标题）。窗口正文 R14/R15 的可用字样里因此也没有它们；顺带挡掉
# `一廿、` 这类荒谬写法——它们本就不该是标题。
_CN_UNITS = {"十": 10, "百": 100, "千": 1000}
_CN_NUMERAL = "".join(_CN_DIGITS) + "".join(_CN_UNITS)

# 编号**首字**不能是百/千（用户 2026-09-15：“千百肯定不接受”）：规范写法是
# “一百二十一”，不是“百二十一”。只卡首位，多位中文数字本身照认
# （十一、三十四、一百二十一都认）。
_CN_NUMERAL_FORBIDDEN_INITIAL = {"百", "千"}

_L1_RE = re.compile(rf"^([{_CN_NUMERAL}]+)、(.*)$")
_L2_RE = re.compile(rf"^[（(]([{_CN_NUMERAL}]+)[）)]\s*(.*)$")
_L3_RE = re.compile(r"^(\d+)[.．、](?!\d)\s*(.*)$")
_L4_RE = re.compile(r"^[（(](\d+)[）)]\s*(.*)$")

_LEVEL_PATTERNS = {1: _L1_RE, 2: _L2_RE, 3: _L3_RE, 4: _L4_RE}

# 句子结束标点——不含书名号/引号等收尾符号，那些结尾的标题是合法的（比如"二、关于《民法典》的思考"）
#
# ⚠️ 用户 2026-09-17 裁决（审计建议补 `…`、`：`、`—`，**被否**）：
#   "题目允许以省略号、冒号、加号、各种右括号和书名号结尾。这是当然的。"
# ⇒ 省略号（…）、冒号（：）、加号（+）、右括号（）、]】）、书名号（》《）
#   **故意不在这个集合里**（＝以它们结尾仍然可能是题目/副标题/标题）。
#   别再拿"补全标点表"当 bug 修一遍。
_TRAILING_PUNCT = "。！？；，、.!?;,"

# 标题应当"简短明确"（规范原文），但这条**只拿来当三/四级标题的采信条件**：
# 整段文字（含编号）去掉空白后超过这个字数，就不当标题、当成正文，同时弹一条
# "疑似标题"提醒。理由是三、四级是"内层"级别，越往内层，"漏了换行、把正文
# 粘到编号后面"这种手滑越常见；一、二级是章节级标题，写长了不算异常，不检查。
#
# 数值与口径（用户 2026-09-15）：**50 字**，且"不算空白"——量的是整段文字（含
# 编号）去掉所有空白（含全角空格）之后的字符数；原来量的是"编号之后那段文字"
# 且阈值 40，与用户"条件里本来也没说超过 40 字不采信"的疑问不符，已改。
_TITLE_MAX_LEN = 50


def _squeezed_len(text: str) -> int:
    """去掉所有空白（含全角空格）之后的字符数。"""
    return len(re.sub(r"\s|\u3000", "", text))

# 英文题目/副标题这一行是英文，不能拿上面那个 50 字的中文阈值卡它
# （"This is the English Title of the stress testing file" 就 51 个字符了）。
# 但它也不能没边：长度只当"这是不是一整段正文"的粗筛，标点结尾才是主判断。
_EN_TITLE_MAX_LEN = 120
# 英文题名（题目/副标题）的采信条件（用户 2026-09-15 定）：
#   * **一个汉字都没有**——有汉字就不是英文题名；
#   * **不以句末标点结尾**——英文正文句子跟题名最硬的分别；
#   * 至少若干拉丁字母——纯数字/符号行不当真。
# 题目要 5 个字母，副标题只要 1 个（副标题可以很短，守太紧反而漏）。
_EN_TITLE_MIN_LETTERS = 5
_EN_SUBTITLE_MIN_LETTERS = 1

# 固定名称锚点，大纲级别上跟一级标题同级（要进目录），字与字之间允许半角/全角空格。
# 但它们各自的具体样式不一定跟普通一级标题相同——比如附录标题规范定的是小四号
# 黑体，跟引言/结论/参考文献/致谢的小三号宋体不是一回事——所以 HeadingMatch 上
# 单独带一个 role 字段区分是哪一个，不能只看 level 就假定套同一个样式。
#
# "绪论"**故意不在这里**：官方写作模板虽说"引言也可以称为序言、绪论或者导论"，
# 但规范正文的部件清单里只有"引言"，预设表里也没有"绪论"这一行。用户核实后明确
# 要求不收——宁可把"绪论"当正文，也不要由工具替用户认定它等于引言（否则一篇用
# "绪论"开头的论文会被判成缺引言、再自动补一个"引　言"出来，凭空多一个标题）。
# 这是用户拍过板的结论，不要自作主张加回来。
_ANCHOR_WORDS = ["引言", "结论", "参考文献", "致谢", "附录"]
_ANCHOR_WORDS_SET = set(_ANCHOR_WORDS)


def _anchor_pattern(word: str) -> re.Pattern:
    sep = r"[ 　]*"
    return re.compile("^" + sep.join(re.escape(c) for c in word) + sep + "$")


_ANCHOR_PATTERNS = [(w, _anchor_pattern(w)) for w in _ANCHOR_WORDS]

# 手打目录条目"不带页码"那一档的边界（用户 2026-09-17 裁决：「允许将引言作为边界」）。
# 只有标签后面能找到它，才把标题形状的行也算成条目；找不到就退回"页码尾巴"那一档。
_INTRO_BOUNDARY_WORD = "引言"

# 摘要/关键词这类前置部分标签——跟标题层级是两回事：不进目录、没有大纲级别，
# 有各自独立的格式要求（摘要标签、关键词标签都是黑体四号加粗顶格，但都不是
# "一级标题"小三号居中那套），所以完全不走 HeadingMatch，单独一套结果
# （见 FrontMatterLabel / detect_front_matter_labels）。规范要求用【】包裹，
# 但用户可能没打括号，两种写法都要认。
#
# 插图清单/附表清单/插图和附表清单（规范 [91]-[93]）也是这个家族：另起一页
# 放在目录后面、相当于"图和表的迷你目录"。它们的"内容"不是文字、是工具插的
# Word 域（跟目录域同机制），所以识别这三个标签只为两件事：源文件里已经有
# 清单时别重复补一份、以及给已有标签套标题格式。
_FRONT_MATTER_WORDS = [
    "摘要", "关键词", "目录",
    "插图清单", "附表清单", "插图和附表清单",
    "Abstract", "Keywords",
]


def _front_matter_pattern(word: str) -> re.Pattern:
    """前置标签的写法：`摘要` / `【摘要】` 都认。

    **中文字与字之间允许半角/全角空格**（用户 2026-09-17："认摘要的空格。"），
    与固定名标题 `_anchor_pattern`（引　言 / 参考文　献）同一口径——以前只有
    后者认空格、前者不认，写成「摘　要」时会被判成"没有摘要" ⇒ 在文稿里又补
    一份空的【摘要】。英文标签（Abstract/Keywords）不加这一条：那里插空格
    不是写法习惯，而是真的拼错了。
    """
    flags = re.IGNORECASE if word.isascii() else 0
    if word.isascii():
        return re.compile(r"^[【\[]?" + re.escape(word) + r"[】\]]?$", flags)
    sep = r"[ 　]*"
    body = sep.join(re.escape(c) for c in word)
    return re.compile(r"^[【\[]?" + body + r"[】\]]?$")


_FRONT_MATTER_PATTERNS = [(w, _front_matter_pattern(w)) for w in _FRONT_MATTER_WORDS]

# "这一页只该装目录/清单条目"的四个部件（`alerts.check_front_matter_body` 也用
# 这一份，别再抄一遍）。
TOC_PAGE_PARTS = {"目录", "插图清单", "附表清单", "插图和附表清单"}

# 手打目录条目的"页码尾巴"：结尾是 省略号/三个点/制表符 + 数字。
# 正常标题不会这么写，所以这是很窄的判据（样式是 `toc N` 的那种由
# `is_toc_paragraph` 拦，这里拦的是**用户手打、没有 toc 样式**的目录）。
_TOC_PAGE_TAIL_RE = re.compile(r"(?:…{2,}|\.{3,}|\t)\s*\d+\s*$")


def hand_typed_toc_entries(document: Document) -> set[int]:
    """手打目录条目的段号集合。两种写法都算，都只在「目录/清单」标签正下方生效：

    一、**带页码尾巴**（用户 2026-09-15 裁决，压力测试观察项 O10）：
        `一、平台责任的一般理论……3` / `一、……\t3`。
    二、**不带页码、但有「引言」做边界**（用户 2026-09-17 裁决：「允许将引言作为边界」）：
        紧跟在标签后面、长得像标题的那几行。这一档**必须先确认后面能找到「引言」**
        才生效——否则「目录标签后面直接写正文标题」会被误伤，那是真实写法。

    为什么这种行不该当标题：它单看文字跟真标题一模一样，认成标题就会套上标题样式、
    写进新生成的目录、还参与编号告警（跳号/未归零/孤立）。实测 `压力测试_Agent版.docx`
    段 22 就是这么漏网的（带制表符的那几条反而因为样式是 `toc N` 被跳过了）。

    判据收得很窄：
      * 只在那四个标签的正下方生效；
      * 从标签往下走，**链一断就停**（遇到既没有页码尾巴、又不是标题形状的段落）——
        目录页里多写的正文、以及它后面的真标题都不会被误伤；
      * 碰到下一个部件标签或引言/结论/参考文献/附录/致谢标题也停；
      * 空行只当排版空隙跳过。
    """
    paragraphs = document.paragraphs
    found: set[int] = set()
    for i, paragraph in enumerate(paragraphs):
        text = paragraph.text.strip()
        if not text:
            continue
        is_label = any(w in TOC_PAGE_PARTS and pattern.match(text)
                       for w, pattern in _FRONT_MATTER_PATTERNS)
        if not is_label:
            continue
        tail_hits: list[int] = []             # 有页码尾巴：一档判据，随时算数
        shape_hits: list[int] = []            # 只是像标题：要等确认有引言
        bounded = False
        for j in range(i + 1, len(paragraphs)):
            line = paragraphs[j].text.strip()
            if not line:
                continue                      # 空行只当排版空隙
            anchor = next((w for w, pattern in _ANCHOR_PATTERNS if pattern.match(line)), None)
            if anchor is not None:
                bounded = anchor == _INTRO_BOUNDARY_WORD   # 引言 ⇒ 这一档生效
                break
            if any(w in TOC_PAGE_PARTS and pattern.match(line)
                   for w, pattern in _FRONT_MATTER_PATTERNS):
                break                         # 又碰到一个前置标签
            if _TOC_PAGE_TAIL_RE.search(line):
                tail_hits.append(j)
                continue
            if _try_pattern_match(line) is not None:
                shape_hits.append(j)
                continue
            break                             # 链断了
        found.update(tail_hits)
        if bounded:
            found.update(shape_hits)
    return found


def chinese_numeral_to_int(text: str) -> int | None:
    if not text or any(c not in _CN_DIGITS and c not in _CN_UNITS for c in text):
        return None
    if text[0] in _CN_NUMERAL_FORBIDDEN_INITIAL:
        # “百二十一”“千一百”这类省略首位的写法不算编号（用户 2026-09-15）。
        # 返回 None 之后 `_try_numbered()` 会接着试下一个级别，最后落回
        # “不是标题”，再由 `_demote_if_odd_numbering()` 弹一条疑似标题提醒。
        return None
    section = 0
    digit = 0
    for ch in text:
        if ch in _CN_DIGITS:
            digit = _CN_DIGITS[ch]
        else:
            unit = _CN_UNITS[ch]
            section += (digit or 1) * unit
            digit = 0
    return section + digit


def _numeral_value(level: int, raw: str) -> int | None:
    if level in (1, 2):
        return chinese_numeral_to_int(raw)
    return int(raw)


@dataclass
class HeadingMatch:
    paragraph_index: int
    level: int
    # None 表示"确定是标题，但提不出编号"（固定名称锚点、大纲级别兜底都常见）；
    # 点分体系（1.1.1）是各段组成的元组。
    number: int | tuple[int, ...] | None
    marker: str
    title: str
    source: str = "regex"  # 恒为 "regex"（大纲级别兜底已于 2026-09-15 删除）
    role: str | None = None  # 固定名称锚点时是具体是哪一个（"附录"/"引言"/…），数字编号标题恒为 None
    # 编号是 Word 自动编号列表画的、**不在段落文字里**：这种段落套版时必须把
    # 编号"固化"成真正的文字，否则清直接格式那一步会把编号引用整个删掉，
    # 成品上编号直接消失（见 normalize.py 的 _materialize_number）。
    # 有默认值的字段一律排在最后，别再往前挪——挪到 level 前面会让 dataclass 直接报错。
    auto_numbered: bool = False
    # 编号要写成什么文字（规范写法，如 "(1)"→"（1） "、半角→全角；或者 Word
    # 自动编号回算出来、原样固化的 "1.1 "），在**识别时的层级**上算一次就
    # 定死：用户在结构页把这一段拖到别的层级时，编号文字不再跟着层级重算
    # （不许把 "（1）" 变成 "1."，见 2026-09-17/2026-09-19 口径）。
    # None = 这一段的编号不规范化，原文原样保留。
    canonical: str | None = None
    # 固定名称锚点（引言/结论）前面带着用户写的编号（`1 引言`）：整段文字
    # 原样保留，不改写成规范的"引　言"。
    keep_text: bool = False


@dataclass
class FrontMatterLabel:
    """摘要/关键词这类前置部分标签，不是标题层级的一部分。"""

    paragraph_index: int
    label: str  # "摘要"/"关键词"/"Abstract"/"Keywords"


@dataclass
class TitleMatch:
    """文档大标题（和可能存在的副标题）。"""

    title_paragraph_index: int
    subtitle_paragraph_index: int | None

    @property
    def subtitle_slot_index(self) -> int | None:
        """副标题**按规矩该落在哪一段**（＝题目紧挨着的下一段）。

        识别与**补件**两边必须取同一个值：副标题的位置判据就是"紧跟题目之后
        那一段"，而补件端从前用的是锚链（"下一个已存在部件之前"）——两者在
        正常稿子上重合，一旦题目被认到别处（封面栏位行 / 说明段抢位）就错开
        ⇒ 下一遍识别端认不出自己插的那条占位语 ⇒ **每遍回炉再插一条**
        （AAT 2026-09-17 穷举实测：写作模板 198→199→200、格式规范 323→324→325）。
        规则真源＝`slot_after()`，两处都调它（别再各算一遍）。
        """
        return slot_after(self.title_paragraph_index)


def slot_after(index: int | None) -> int | None:
    """`index` 紧挨着的下一段——就是"副标题/英文副标题"该落的那一段。

    中文副标题、英文副标题（`detect_english_subtitle()` 里的 `english_title + 1`）
    都是这一条：**上一个部件后面第一段**。所以补件端要插这两个部件时，
    取这里算的槽位当锚点（`apply_missing_parts(subtitle_slots=…)`），不要自己推导。
    """
    return None if index is None else index + 1


@dataclass
class DemotedTitle:
    """命中了标题线索、但**没被采信为标题**的段落。交给上层提醒用户。

    两种来源（看 `reason`）：
      * `too_long`（默认）：命中了编号正则，但三/四级标题太长，怀疑是漏了换行。
      * `odd_numbering`：用的是规范之外的编号写法（`①`、`Ⅰ.`、`第一章`、`1)`）。

    标点结尾的不会出现在这里——那种情况直接静默归正文，见 `_classify()` 的说明。

    ⚠️ 原来还有一个 `hint` 属性（写着"怎么改才能被识别为标题"），2026-09-15
    用户把告警文案统一成 `疑似标题：请检查“…”是否为标题，在段落（N）。` 之后
    没人再用它了，已删。`reason` / `numbering_style` 仍然带着——它们是数据，
    将来要重新区分两种来源时还得用。
    """

    paragraph_index: int
    level: int
    marker: str
    title: str
    reason: str = "too_long"
    numbering_style: str = ""  # reason == "odd_numbering" 时，注明是哪种写法


def _try_numbered(text: str) -> tuple[int, str, str, int] | None:
    for level in (1, 2, 3, 4):
        m = _LEVEL_PATTERNS[level].match(text)
        if not m:
            continue
        raw_number, title = m.group(1), m.group(2).strip()
        number = _numeral_value(level, raw_number)
        if number is None:
            continue
        return level, raw_number, title, number
    return None


def _try_anchor(text: str) -> tuple[int, str, str, None] | None:
    for word, pattern in _ANCHOR_PATTERNS:
        if pattern.match(text):
            return 1, word, "", None
    return None


def _try_pattern_match(text: str) -> tuple[int, str, str, int | None] | None:
    return _try_numbered(text) or _try_anchor(text)


# 规范之外的常见编号写法。**只用来提醒用户，不拿来套标题格式**——规范明文
# 规定了四级标题怎么写（`一、`/`（一） `/`1. `/`（1） `），认得太宽会把正文里
# 的列举（"1) ……" 、"①……"）也吸进来，误伤比漏报严重。用户 2026-09-12 的
# 决定就是"弹窗告警，大不了让用户手动在 GUI 里调标题层级"。
_ODD_NUMBERING_PATTERNS: tuple[tuple[str, str], ...] = (
    (r"^([①-⑳])", "带圈数字"),
    (r"^(Ⅰ|Ⅱ|Ⅲ|Ⅳ|Ⅴ|Ⅵ|Ⅶ|Ⅷ|Ⅸ|Ⅹ)+[.、．\s]", "罗马数字"),
    (r"^第[一二三四五六七八九十百零〇\d]+[章节讲条]", "“第X章/节”式编号"),
    (r"^\d+[)）]", "数字加右括号"),
    (r"^[㈠㈡㈢㈣㈤㈥㈦㈧㈨㈩]", "带括号序号"),
    # 中文数字以“百/千/廿/卅”开头（“百二十一、”“廿一、”）——用户
    # 2026-09-15：“千百肯定不接受”；2026-09-16：“不认廿卅，对其报疑似标题”。
    # 这种写法长得像一级标题，但不采信，只提醒（与其它非规范编号一个待遇）。
    (r"^[百千廿卅][零一二三四五六七八九十百千廿卅]*[、.．]", "省略首位的“百/千/廿/卅”式编号"),
)


def _try_odd_numbering(text: str) -> str | None:
    """匹配"非规范编号写法"，返回那种写法的名字（拿给用户看）；不像则返回 None。"""
    for pattern, name in _ODD_NUMBERING_PATTERNS:
        if re.match(pattern, text):
            return name
    return None


def _demote_if_odd_numbering(demoted: list[DemotedTitle], index: int, text: str) -> None:
    """看起来像标题、用的却是规范外的编号写法 → 进"提醒用户"队列。

    两条静默门限跟 `_classify()` 保持一致：以标点结尾的（更像正文里的列举）、
    以及过长的（同样按“去掉空白后的整段”计），都不报——宁可漏报，也别把
    正文一条条弹出来烦人。
    """
    text = text.strip()
    if (not text or text.endswith(tuple(_TRAILING_PUNCT))
            or _squeezed_len(text) > _TITLE_MAX_LEN):
        return
    style = _try_odd_numbering(text)
    if style is None:
        return
    demoted.append(
        DemotedTitle(index, 0, "", text, reason="odd_numbering", numbering_style=style)
    )


def _classify(level: int, title: str, paragraph_text: str = "") -> tuple[bool, str | None]:
    """返回 (是否采信为标题, 需要提醒用户的原因或 None)。

    标点结尾：标题不该带标点，这个信号足够确定——直接判定不是标题，静默归
    正文，不弹提醒（不确定的事才需要问用户，这个不属于不确定）。

    偏长：**只查三、四级**，量的是整段文字（`paragraph_text`，含编号）去掉
    空白后的长度；超过 `_TITLE_MAX_LEN` 就不采信、归正文，并返回 "too_long"
    让上层弹一条提醒。一、二级是章节级标题，写长了不算什么异常信号，不检查。
    """
    if title.endswith(tuple(_TRAILING_PUNCT)):
        return False, None
    if level in (3, 4) and _squeezed_len(paragraph_text or title) > _TITLE_MAX_LEN:
        return False, "too_long"
    return True, None


def _numPr_of(paragraph: Paragraph) -> tuple[str, int] | None:
    pPr = paragraph._p.find(qn("w:pPr"))
    numPr = pPr.find(qn("w:numPr")) if pPr is not None else None
    if numPr is None:
        return None
    num_id_el = numPr.find(qn("w:numId"))
    num_id = num_id_el.get(qn("w:val")) if num_id_el is not None else None
    if num_id in (None, "0"):
        return None
    ilvl_el = numPr.find(qn("w:ilvl"))
    ilvl = int(ilvl_el.get(qn("w:val"))) if ilvl_el is not None else 0
    return num_id, ilvl


def _resolve_text(paragraph: Paragraph, numbering: NumberingScheme) -> tuple[str, bool]:
    """段落实际"看起来"是什么文字——如果用了 Word 自动编号，把反算出来的
    编号拼到前面；每段只调用一次，保证计数器按文档顺序正确推进。

    第二个返回值说明这个编号是不是**只存在于 Word 的渲染里**（段落文字里
    根本没有）——是的话，套版时得先把它固化成人能看见的文字，否则清直接
    格式会把编号引用 numPr 一起删掉，成品上编号凭空消失（见 normalize.py
    的 _materialize_number）。
    """
    text = paragraph.text.strip()
    num_ref = _numPr_of(paragraph)
    if num_ref is None:
        return text, False
    marker = numbering.render(*num_ref)
    if not marker:
        return text, False
    return f"{marker}{text}", True


@dataclass
class ManualStructure:
    """人工核对/修正后的识别结果，四个探测器的输出打包在一起。

    来自"校对和调整识别结果"窗口（`adjust_preview.build_manual_structure`）：
    用户在那边拖动调整过身份地图之后，`pipeline.run()` 拿这份结果整体替换掉
    它自己重新跑一遍 `detect_headings`/`detect_front_matter_labels`/
    `detect_document_title`/`detect_captions` 的结果，而不是每次导出都从头
    自动识别、把用户刚做的修正扔掉。
    """

    matches: list[HeadingMatch | None]
    front_matter: list[FrontMatterLabel]
    title: TitleMatch | None
    captions: list[CaptionMatch]
    # 标成"长引文"的段落下标——这个没有对应的自动探测器，完全来自用户在
    # 调整窗口里手动勾的 ☐引用，见 apply_structure.apply_long_quotes。
    quote_paragraphs: list[int]
    # 英文题目在第几段——英文题目在调整窗口里是独立部件（kind="entitle"，
    # 见 adjust_preview.build_rows），不是从 detect_english_title() 重新
    # 识别一遍，是从调整窗口当前的排列直接读的。
    english_title: int | None = None
    # 英文副标题在第几段（用户 2026-09-15 要的"完整的对称部件"）。跟英文题目
    # 一样，从调整窗口当前的排列直接读（kind="subtitle" + block="英文副标题"），
    # 不在管线里重新识别一遍。
    english_subtitle: int | None = None
    # 用户指定的封面范围：**第 0 段到第 cover_end 段（含）都是封面**。
    # 用户 2026-09-15："我们不作任何程度的封面自动识别，纯靠用户手敲"——
    # 所以这里只有用户手填的值，没有、也不该有自动判定出来的值。
    # 落在这一区间的段落：不参与识别（识别结果被滤掉）、不套样式、不清直接
    # 格式、不改图片；成品里**原样保留**（要删用户自己在 Word 里删）。
    cover_end: int | None = None
    # 导出时把封面那几段**真删掉**（用户 2026-09-15 在"设置封面范围"里加的
    # 复选框）。False（默认）＝保留原样。删除位置很讲究，见 `pipeline.run()`
    # 里那一段注释：必须在所有按下标干活的步骤之后、又必须在自动封面拼进来之前。
    delete_cover: bool = False
    # "看着像标题、但没被采信"的段落（`DemotedTitle` 列表，见
    # `detect_headings`）。2026-09-14 接线：以前手动结构这条路径把它整份丢掉
    # （`pipeline.run()` 里写死 `demoted = []`），于是在调整窗口里核对过结构
    # 之后，这类提醒就再也不出现了——而那批段落是**一点标题格式都拿不到**
    # 的，静默丢掉最坑人。
    demoted: list[DemotedTitle] = field(default_factory=list)
    # 用户在结构页里勾了"重新为文档生成"的目录/清单部件（`{"目录"}`、
    # `{"插图清单"}`……）。勾了就把源文件里那份域删掉、由工具重新生成；
    # 没勾（默认）＝已有真域一律不动（用户 2026-09-15）。
    regenerate_fields: set[str] = field(default_factory=set)
    # 被用户从自动编号标题**拖回“正文”**的段落：`{段号: 要写成的手打编号文字}`。
    # 这些段落在成品里已经不是标题了（不套样式、不进目录），但那个编号是自动
    # 编号画的、不在段落文字里——不管它的话成品里那段编号就凭空消失（用户
    # 2026-09-17 报的严重 bug：“拖到正文将会导致自动编号永久丢失”）。所以还是
    # 把编号写成手打文字；文字＝识别时那个层级的规范写法（跟它当标题时显示、
    # 写出来的一模一样，见 `normalize.normalize_headings()`）。
    auto_bodies: dict[int, str] = field(default_factory=dict)
    # 结构页上那个"自动重新编号参考文献"复选框（用户 2026-09-15："允许用户自定"）：
    #   剥掉手打的 `[1]`/`［1］`、挂上 `[%1]` 编号定义（增删条目自动重排）。默认**开**。
    # （题注侧原来也有一个"自动重新编号题注"复选框，2026-09-16 用户定"我们只管样式、
    #   编号与我们无关"之后**整个删掉**了——手打编号和用户的域都原样保留。）
    # 没传 `manual_structure`（＝没打开过结构页）时这个开关不生效，走老行为（照旧重编）。
    renumber_references: bool = True


def detect_headings(document: Document) -> tuple[list[HeadingMatch | None], list[DemotedTitle]]:
    paragraphs = document.paragraphs
    numbering = load_numbering_scheme(document)
    results: list[HeadingMatch | None] = [None] * len(paragraphs)
    demoted: list[DemotedTitle] = []
    # 用户手打的目录条目（见 `hand_typed_toc_entries`）：也不当标题，而且**不报
    # "疑似标题"**——它们本来就不该有标题身份，报了只是噪声（它旁边那条
    # "多余的正文"告警才是在说正确的事）。
    hand_typed_toc = hand_typed_toc_entries(document)

    for i, paragraph in enumerate(paragraphs):
        # 目录/清单域生成的条目（样式 `toc 1`/`toc 2`…）不是标题。它们里面装的
        # 是**标题的抄本**（"一、一级标题\t2"），单看文字跟真标题一模一样，
        # 不看样式就会把一整目录的条目都当标题：给它们套上标题样式、写进新目录、
        # 还会报一堆"孤立标题"。写作模板.docx 实拍（2026-09-15）：旧目录 14 条
        # 全部被当成标题、报了 15 条告警。
        if is_toc_paragraph(paragraph) or i in hand_typed_toc:
            continue
        # 图表目录/清单的条目段落（Word 内置样式 `table of figures`）同理。
        # 用户 2026-09-16 实测（`TOC.docx`）：那一行文字是 `表 1 1`（＝条目
        # `表 1` + 页码 `1`），没被认出来就是一条真一级标题。
        if is_list_entry_paragraph(paragraph):
            continue
        text, auto_numbered = _resolve_text(paragraph, numbering)

        # 大纲级别（`w:outlineLvl`）**完全不采信**（用户 2026-09-15：“大纲格式
        # 似乎完全冗余，完全不采信即可”）。这里原来有一支“有大纲级别→当成
        # 候选标题”的分支，实测它与下面这条主路**逐行等价**（同一段“第一章
        # 绪论”：带级别与不带级别的识别结果、告警条数完全一样），所以直接删掉。
        # 注意：**输出侧仍然要写大纲级别**（四个标题样式 + 附录标题都要补
        # outlineLvl），那是为了让 Word 的目录域能收到标题，跟识别无关。
        if not text:
            continue

        matched = _try_pattern_match(text)
        if matched is None:
            _demote_if_odd_numbering(demoted, i, text)
            continue

        level, raw_number, title, number = matched
        role = raw_number if raw_number in _ANCHOR_WORDS_SET else None
        # 整段文字（`text`，含编号）交给 `_classify()` 量长度：用户 2026-09-15
        # 定的口径是“50 字，不算空白，量整段”。
        is_heading, reason = _classify(level, title, text)
        if is_heading:
            results[i] = HeadingMatch(
                i, level, number, raw_number, title, role=role, auto_numbered=auto_numbered
            )
        elif reason is not None:
            demoted.append(DemotedTitle(i, level, raw_number, title, reason))

    return results, demoted


def check_sequence(matches: list[HeadingMatch | None]) -> list[tuple[int, str, str]]:
    """检查标题编号的三种异常：未归零、跳号、重号、孤立。跳级本身不算问题。

    返回 `(段落下标, 错误名称, 错误概要)` 列表——段落下标就是结构页"段号"列
    里显示的那个数字（2026-09-14 用户要求告警必须能对到屏幕上的段号），
    **名称和概要里都不写位置**，位置统一由 `alerts.Alert.line` 拼。

    错误名称/概要是用户 2026-09-15 逐条定的（详见
    `docs/告警文案编号表.md`）：`未归零`／`跳号`／`重号`／`孤立`。原来的
    缺号（2→5）与回退（3→2）分两种说法，用户要求**合并成一句"跳号"**；
    重号也不再自己一句，跟跳号共用"未连续使用的标题序号"（名称换成
    `重号`）。
    """
    flat = [m for m in matches if m is not None]
    warnings: list[tuple[int, str, str]] = []

    last_number: dict[int, int | None] = {}
    expect_reset: dict[int, bool] = {}

    for idx, m in enumerate(flat):
        level = m.level

        for deeper in [lv for lv in last_number if lv > level]:
            del last_number[deeper]
        for deeper in [lv for lv in expect_reset if lv > level]:
            del expect_reset[deeper]

        # 跨级未归零：上一次遇到更浅层级之后，这是该层级第一次出现，编号该是1
        if expect_reset.get(level) and m.number is not None and m.number != 1:
            warnings.append((m.paragraph_index, "未归零", "未从头开始的标题序号"))
        expect_reset[level] = False
        expect_reset[level + 1] = True

        # 同级编号不连续：缺号、重号、回退三种，措辞见上面的说明
        prev = last_number.get(level)
        if prev is not None and m.number is not None and m.number != prev + 1:
            name = "重号" if m.number == prev else "跳号"
            warnings.append((m.paragraph_index, name, "未连续使用的标题序号"))
        last_number[level] = m.number

        # 孤立列举：前后（跨过更深层级）都没有同级的兄弟标题
        if not _has_sibling(flat, idx, level, -1) and not _has_sibling(flat, idx, level, 1):
            warnings.append((m.paragraph_index, "孤立", "单项使用的标题序号"))

    return warnings


def _has_sibling(flat: list[HeadingMatch], idx: int, level: int, direction: int) -> bool:
    i = idx + direction
    while 0 <= i < len(flat):
        if flat[i].level < level:
            return False  # 碰到更浅的层级，说明已经出了当前的上级范围
        if flat[i].level == level:
            return True
        i += direction
    return False


def find_unresolved_auto_numbered_paragraphs(document: Document) -> list[int]:
    """找出用了 Word 自动编号列表、但反算不出编号文字的段落（列表定义缺失、
    或者是项目符号/不支持的编号格式）——这种真算不出来，报出来交给人看。
    正常能反算出编号的，已经并入 detect_headings() 正常判断，不在这里重复报。
    """
    numbering = load_numbering_scheme(document)
    found = []
    for i, paragraph in enumerate(document.paragraphs):
        num_ref = _numPr_of(paragraph)
        if num_ref is None:
            continue
        if numbering.render(*num_ref) is None:
            found.append(i)
    return found


# 附录序号题名：`附录A …`（用户 2026-09-15："附录A-Z识别要做"）。
# 只认拉丁字母序号（A/B/C…），后面可以跟空白/顿号/点号/冒号再接题名。
_APPENDIX_ENTRY_RE = re.compile(r"^附\s*录\s*([A-Za-z])[ \t\u3000、．.：:]*")
_APPENDIX_ENTRY_MAX_LEN = 40


def _appendix_entries_in_range(document: Document, start: int, end: int) -> list[tuple[int, str]]:
    """共用的扫描核心：`(start, end)` 区间里形如 `附录A …` 的段落。"""
    paragraphs = document.paragraphs
    found: list[tuple[int, str]] = []
    for i in range(start + 1, min(end, len(paragraphs))):
        paragraph = paragraphs[i]
        if is_toc_paragraph(paragraph):
            continue
        text = paragraph.text.strip()
        if not text or len(text) > _APPENDIX_ENTRY_MAX_LEN:
            continue
        if text.endswith(tuple(_TRAILING_PUNCT)):
            continue
        match = _APPENDIX_ENTRY_RE.match(text)
        if match:
            found.append((i, match.group(1).upper()))
    return found


def detect_appendix_entries(document: Document, matches) -> list[tuple[int, str]]:
    """附录区间里形如 `附录A …` 的段落，返回 `[(段号, 序号字母), …]`。

    规范 [226] 要求"另起一行标注附录序号和题名"，但这类段落**不是标题**：
    不进目录、不参与编号告警（否则"附录A"会被当成孤立的一级标题报出来），
    只该套 S20（附录序号题名）样式——所以单独一套结果，不走 HeadingMatch。

    `matches`：`list[HeadingMatch|None]`（旧 `detect_headings()` 的产出）——
    **仍在用**，`essay_tool.gui.adjust_preview.build_rows()` 还没切到新的
    识别架构，这个签名不能动；`pipeline.py`（新架构）改走
    `detect_appendix_entries_from_index()`，两者共享 `_appendix_entries_in_range()`
    这个扫描核心，不是两套判据。

    范围限制得很死（宁可不认）：只在**"附录"标题之后**、下一个锚点部件
    （参考文献/致谢…）或文末之前找；段落以"附录+一个拉丁字母"开头、
    不以句末标点结尾、长度不超过 `_APPENDIX_ENTRY_MAX_LEN`；目录条目
    （`toc N` 样式）不算。
    """
    paragraphs = document.paragraphs
    starts = [m.paragraph_index for m in matches
              if m is not None and m.role == "附录"]
    if not starts:
        return []
    start = starts[0]
    end = len(paragraphs)
    for m in matches:
        if (m is not None and m.paragraph_index > start
                and m.role in ("参考文献", "结论", "致谢")):
            end = m.paragraph_index
            break
    return _appendix_entries_in_range(document, start, end)


def detect_appendix_entries_from_index(document: Document, index) -> list[tuple[int, str]]:
    """`detect_appendix_entries()` 的等价物，喂 `recognize.recognize()` 产出的
    识别索引（`pipeline.py` 新架构专用）——判据跟上面那个完全一样，只是身份
    的来源从 `HeadingMatch.role` 换成 `Entry.part`。"""
    from essay_tool.engine.doc_index import paragraph_ordinals  # 局部导入：避免循环引用

    ordinals = paragraph_ordinals(index)
    starts = [i for i, e in ordinals.items() if e.part == "附录"]
    if not starts:
        return []
    start = starts[0]
    end = len(document.paragraphs)
    for i, e in ordinals.items():
        if i > start and e.part in ("参考文献", "结论", "致谢"):
            end = min(end, i)
    return _appendix_entries_in_range(document, start, end)


def detect_front_matter_labels(document: Document) -> list[FrontMatterLabel]:
    """摘要/关键词（及其英文对应）——允许有没有【】/[]包裹两种写法。"""
    found = []
    for i, paragraph in enumerate(document.paragraphs):
        text = paragraph.text.strip()
        if not text:
            continue
        for word, pattern in _FRONT_MATTER_PATTERNS:
            if pattern.match(text):
                found.append(FrontMatterLabel(i, word))
                break
    return found


def detect_english_title(
    document: Document,
    front_matter: list[FrontMatterLabel],
    cover_limit: int | None = None,
    known: Iterable[int] = (),
) -> int | None:
    """英文题目——**独立识别**，不依赖【ABSTRACT】标签。

    用户 2026-09-15："英文题目/摘要/关键词既然是三个不同的部分，当然应该
    独立检出。" 所以这个探测器只回答一个问题：源文件里有没有一段像英文
    题目的正文。

    **只看一个区间**（用户 2026-09-16：“英文标题怎么可能出现在中文摘要和
    关键词以内？这是严重恶性 bug”）：

    * 下界 = 中文摘要/关键词里**靠后的那一个标签**（有【关键词】就取它，
      没有就取【摘要】；两个都没有就取封面之后）；
    * 上界 = 英文摘要/英文关键词标签（【ABSTRACT】/【KEYWORDS】）、目录与三张
      清单的标签、以及第一个正文标题，三者取最靠前的那个；
    * 区间里**第一个像英文题名的段落**就是它；区间里先出现几行别的文字
      （关键词自己的内容）不影响。

    位置口径的来由（2026-09-16 对真实稿子的实测）：规范里正文的部件顺序是
    中文摘要 → 中文关键词 → 英文题目 → 英文副标题 → 英文摘要 → 英文关键词
    → 目录，所以英文题目就在“中文关键词之后、【ABSTRACT】（或目录、或正文）
    之前”这一小段里。`图表脚注测试.docx` 与 `压力测试_Agent版.docx` 两份夹具
    的英文题目都正好落在这里。以前按“已识别部件切区间、逐个区间找”的写法，
    把放在中文摘要与中文关键词之间的句子也当成了英文题目（实测复现）。

    `cover_limit`：封面末段（含），封面区域不参与（用户手敲的封面里可能
    有一行英文校名，不该被当成题目）。`known`：调用方已经识别出来的其它
    部件段号（中文题目、中文副标题…），一并当边界。

    它是不是一个英文题名，看 `_looks_like_english_title()`：**一个汉字都
    没有 + 至少 5 个拉丁字母 + 不以句末标点结尾**（+ ≤120 字符的粗筛）。
    **不看是否全大写**——规范要求全大写，但用户可能完全不管这条，不能拿
    大小写当识别依据。

    “英文题目存在”与“英文摘要那一页存在”是两回事（后者只看【ABSTRACT】
    标签），别把两者揺在一起——一个文档完全可以只有英文题目。
    """
    paragraphs = document.paragraphs
    end = len(paragraphs)
    # `cover_limit` 是**下界**（封面末段，含），不是上界：探测范围始终到文末，
    # 只是封面那几段不参与（写错成上界的话，传了封面范围反而只能看封面那几段）。
    floor = -1 if cover_limit is None else min(cover_limit, end - 1)

    # 工具自己插的占位语不是内容，是"某个部件的槽位"：成品回炉时英文副
    # 标题的占位语就贴在【ABSTRACT】上方，旧逻辑不校验就直接把它当成英文
    # 题目 ⇒ "副标题缺失" ⇒ 又插一条，**每回炉一次多一行**（实测 83→84 段）。
    # 现在：碰到"英文题目"那个占位语就认它是题目自己的槽位；碰到别的占位语
    # 就跳过（它不是内容，也不该当正文句子把区间判死）。
    numbering = load_numbering_scheme(document)
    label_at = {f.paragraph_index: f.label for f in front_matter
                if floor < f.paragraph_index < end}
    known_at = {i for i in known if i is not None and floor < i < end}
    hand_typed = hand_typed_toc_entries(document)   # 手打目录条目不当边界（也不是正文）
    heading_at: set[int] = set()
    for i in range(floor + 1, end):
        if i in label_at or i in known_at or i in hand_typed:
            continue
        if is_toc_paragraph(paragraphs[i]):
            continue
        text, _auto = _resolve_text(paragraphs[i], numbering)
        if text and _try_pattern_match(text) is not None:
            heading_at.add(i)

    # 下界：中文摘要/关键词里靠后的那一个标签——而且看**第一份前置**里的那一个。
    # 取全局最大下标的写法在"正文被整体复制"的稿子上会出事：第二份的
    # 【关键词】把窗口整个推到第二份去，而"补件"的锚点落在第一份 ⇒ 英文副标题
    # 每遍都被判成缺失、每遍插一条（AAT 2026-09-17："body 内容整体复制"
    # 每遍 +1 段，实测 `esub=None`、`enter` = 第二份的英文题目）。
    # 只取第一个【摘要】与第一个【关键词】中靠后的那个，等同于原来的写法
    # （每个标签只出现一次时完全一致）：【摘要】…【关键词】之间的英文句子
    # 依旧被排除，用户 2026-09-16 的那条裁决不变。
    cn_anchor = floor
    first_seen: dict[str, int] = {}
    for i, label in sorted(label_at.items()):
        if label in ("摘要", "关键词") and label not in first_seen:
            first_seen[label] = i
    for i in first_seen.values():
        cn_anchor = max(cn_anchor, i)
    # 上界：英文标签、目录/清单标签、第一个正文标题，取最靠前的。
    stops = [i for i, label in label_at.items()
             if label in ("Abstract", "Keywords") or label in TOC_PAGE_PARTS]
    stops += list(heading_at)
    stop = min([i for i in stops if i > cn_anchor], default=end)

    for i in range(cn_anchor + 1, stop):
        text = paragraphs[i].text.strip()
        if not text:
            continue
        if is_english_title_placeholder(text):
            return i
        if is_placeholder(text):
            continue
        if _looks_like_english_title(text):
            return i
    return None


def detect_english_subtitle(
    document: Document,
    english_title: int | None,
    recognized: set[int] | frozenset[int] = frozenset(),
) -> int | None:
    """英文副标题——**英文题目之后紧邻的那一段**（用户 2026-09-15 要的对称部件）。

    判据（用户 2026-09-15："保留不以标点结尾，然后改动"）：
      * 英文题目本身就缺失 → 没有主标题，副标题自然不存在；
      * 紧跟英文题目之后**那一段**（不跳空段）——空着一段就是作者有意隔开；
      * 非空、**不以标点结尾**（带句末标点的是一句话，不是题名）；
      * 本身不是已识别的标题/标签（`recognized`：标题 + 前置标签）；
      * **一个汉字都没有**、至少 1 个拉丁字母（见 `_looks_like_english_title`
        的 `min_letters`）。长度不设中文那套阈值——就是紧跟主标题的另一行
        英文，位置信息已经够硬了。

    规范（[196]）假定英文题目只有一行，没规定过英文副标题——所以这是"用户的
    文件比规范更丰富"时工具该兼容的东西，不是规范要求项。不做的话，两行英文
    题名里只有紧贴【ABSTRACT】那一行会被当英文题目，主题目行反而没有身份。
    """
    if english_title is None:
        return None
    index = english_title + 1
    paragraphs = document.paragraphs
    if index >= len(paragraphs) or index in recognized:
        return None
    text = paragraphs[index].text.strip()
    if not text:
        # 空段不跳（作者有意隔开 ⇒ 不是副标题）；但工具插的占位语可能落在空段
        # 后面（插入点是"下一个已知部件"）⇒ 认它，不然每遍重插一条。
        # AAT 2026-09-17："body 内容整体复制"那份稿子每遍 +1 段就是这个。
        j = index
        while j < len(paragraphs) and not paragraphs[j].text.strip():
            j += 1
        if (j < len(paragraphs) and j not in recognized
                and is_english_subtitle_placeholder(paragraphs[j].text.strip())):
            return j
        return None
    if is_english_subtitle_placeholder(text):
        # 工具自己给"英文副标题"留的坑：它就在副标题该在的位置上——认它，
        # 不然回炉导出时会再插一条（用户 2026-09-15 实测的那个 bug）。
        return index
    if not _looks_like_english_title(text, min_letters=_EN_SUBTITLE_MIN_LETTERS):
        return None
    return index


def _looks_like_english_title(text: str, min_letters: int = _EN_TITLE_MIN_LETTERS) -> bool:
    """一段文字像不像英文题名（题目与副标题共用，只差字母数阈值）。

    三条（用户 2026-09-15 定的条件）：
      * **一个汉字都没有**（有汉字就不是英文题名，那是中文部分的内容）；
      * **不以句末标点结尾**（英文正文句子跟题名最硬的分别）；
      * 至少 `min_letters` 个拉丁字母（纯数字/符号行不当真）。
    长度只当粗筛（`_EN_TITLE_MAX_LEN`）：太长的一定是一整段正文，不是题名。
    原来说过"不看是否全大写"——规范要求全大写，但用户可以不管这条，不能拿
    大小写当识别依据。
    """
    if not text or len(text) > _EN_TITLE_MAX_LEN:
        return False
    if text.endswith(tuple(_TRAILING_PUNCT)):
        return False
    if any("\u4e00" <= ch <= "\u9fff" for ch in text):
        return False
    return sum(1 for ch in text if ch.isascii() and ch.isalpha()) >= min_letters


# ------------------------------------------------------- 目录域 / 清单域
# 目录、三张清单（插图清单/附表清单/插图和附表清单）要不要生成，靠的是
# "真正的 Word 域"，而不是用户手打的那几个字。用户 2026-09-15："至少对目录，
# 只能认真正的 toc 域，不能认用户手动输入的标签和正文；目录列本来就不应该
# 有正文。" 标签可以拿来用（用户只写了【目录】两个字、没内容 → 我们就在它
# 后面接上自己生成的域），但标签下面若还写着正文，那就是文档写错了，
# 由 `alerts.check_front_matter_body` 报警。
#
# 目录域和清单域是同一家族（都是 `TOC`）：目录是 `TOC \o "1-3"`（按大纲
# 级别收标题），清单是 `TOC \c "图"` / `TOC \c "表"`（按题注样式收条目）。
# 这些域指令住在 `w:instrText` / `w:fldSimple/@w:instr` 里，**不在 runs 里**——
# `paragraph.text` 看不到它们，得直接扒 XML。
_TOC_INSTR_RE = re.compile(r"^\s*TOC(\s|$)", re.IGNORECASE)


def paragraph_field_codes(paragraph: Paragraph) -> list[str]:
    """段落里出现的域指令原文（`w:instrText` 的文本 + `w:fldSimple/@w:instr`）。"""
    codes: list[str] = []
    for el in paragraph._p.iter():
        if el.tag == qn("w:instrText") and el.text:
            codes.append(el.text)
        elif el.tag == qn("w:fldSimple"):
            instr = el.get(qn("w:instr"))
            if instr:
                codes.append(instr)
    return codes


def has_toc_field(paragraph: Paragraph) -> bool:
    """这一段里是不是目录/清单域（TOC 家族）。"""
    return any(_TOC_INSTR_RE.match(code) for code in paragraph_field_codes(paragraph))


def toc_field_part(paragraph) -> str:
    """这个域属于哪个部件：目录还是哪张清单。

    目录域是 `TOC \\o "1-3"`；清单域带 `\\c "图"` / `\\c "表"`，收哪种就是
    哪张清单（合并形态是**两个域**，各收一种）。
    """
    kinds = list_field_kinds(paragraph)
    if "图" in kinds and "表" in kinds:
        return "插图和附表清单"
    if "图" in kinds:
        return "插图清单"
    if "表" in kinds:
        return "附表清单"
    return "目录"


# `TOC \c "图"`——Word「插入表目录」写的清单域用 `\c` 参数说明收哪类题注。
_LIST_FIELD_KIND_RE = re.compile(r"\\c\s+\"?([^\"\s\\]+)\"?")
# `TOC \h \z \t "图题,1"`——**我们自己写**的清单域用 `\t` 按样式名收
# （见 `sections._LIST_FIELD_INSTRS`）。两种写法都要认：只认 `\c` 的话，
# 我们自己刚插进去的清单域在探测端等于不存在（实测 `list_field_kinds()` 返回
# 空集）⇒ 一串连锁故障：`toc_field_part()` 把清单域判成"目录"、结构页三行都
# 显示"word自动目录"、用户勾"重新为文档生成"之后 `drop_regenerated_fields()`
# 把两个清单域一起删掉（2026-09-15 实测：3 个域只剩 1 个，清单标题还在、
# 条目没了，且不给任何提示）。
_LIST_FIELD_STYLE_RE = re.compile(r'\\t\s+"([^"]*)"')
# 样式名 → 题注种类。不认识的样式名（用户自己定义的）不算，宁可不认。
_STYLE_NAME_KINDS = {"图题": "图", "表题": "表"}


def list_kinds_in_instr(code: str) -> set[str]:
    r"""一条域指令里"收了哪几种题注"——`\c "图"` 与 `\t "图题,1"` 两种写法都认。

    公开出来是因为**不只探测需要它**：`cleanup.drop_regenerated_fields()` 里
    那个"域包在内容控件里"的分支也要判"这个域是哪个部件"，以前那里自己写了
    一份只认 `\c` 的正则（2026-09-15 审计抓到的同类残留：只认 `\c` 会把
    我们自己写的清单域判成"目录"，勾"重新生成目录"时把它删掉）。
    """
    kinds = {m.group(1).strip() for m in _LIST_FIELD_KIND_RE.finditer(code)}
    for match in _LIST_FIELD_STYLE_RE.finditer(code):
        # `"图题,1"` / `"图题,1,表题,1"`：按逗号切开，只看样式名那一半（后一半
        # 是目录级别，对我们没意义）。
        for item in match.group(1).split(","):
            kind = _STYLE_NAME_KINDS.get(item.strip())
            if kind is not None:
                kinds.add(kind)
    return kinds


# 旧名字（模块内一直这么叫），保留别名免得别处（含历史脚本）引用不到。
_list_kinds_in_instr = list_kinds_in_instr


def list_field_kinds(paragraph: Paragraph) -> set[str]:
    """这一段里的清单域收了哪几种题注（`TOC \\c "图"` / `TOC \\t "图题,1"`）。"""
    kinds: set[str] = set()
    for code in paragraph_field_codes(paragraph):
        if not _TOC_INSTR_RE.match(code):
            continue
        kinds |= _list_kinds_in_instr(code)
    return kinds


def find_list_field_kinds(document: Document) -> set[str]:
    """整篇文档里清单域已经收了哪几种题注（"图"/"表"）。"""
    kinds: set[str] = set()
    for paragraph in document.paragraphs:
        kinds |= list_field_kinds(paragraph)
    return kinds


def _body_instr_texts(document: Document) -> list[str]:
    """整篇 body 里所有域指令（**包括藏在内容控件里的**）。

    `paragraph.text`／`document.paragraphs` 都看不见域，也看不见内容控件
    （`w:sdt`）内部的段落；判断"文档里到底有没有目录/清单域"必须直接扫 XML，
    否则一份把目录包在内容控件里的稿子会被当成"没有目录"、又被插一份。
    """
    body = document._element.body
    found = [el.text or "" for el in body.iter(qn("w:instrText"))]
    for el in body.iter(qn("w:fldSimple")):
        instr = el.get(qn("w:instr"))
        if instr:
            found.append(instr)
    return found


def body_toc_field_parts(document: Document) -> set[str]:
    """整篇文档里已经存在哪些目录/清单域（含内容控件里的）。

    返回值是部件名集合，如 `{"目录", "插图清单"}`；空集＝一个域都没有。
    """
    parts: set[str] = set()
    for code in _body_instr_texts(document):
        if not _TOC_INSTR_RE.match(code):
            continue
        kinds = _list_kinds_in_instr(code)
        if "图" in kinds and "表" in kinds:
            parts.add("插图和附表清单")
        elif "图" in kinds:
            parts.add("插图清单")
        elif "表" in kinds:
            parts.add("附表清单")
        else:
            parts.add("目录")
    return parts


def body_list_field_kinds(document: Document) -> set[str]:
    """整篇文档里清单域收了哪几种题注（含内容控件里的）。"""
    kinds: set[str] = set()
    for code in _body_instr_texts(document):
        if not _TOC_INSTR_RE.match(code):
            continue
        kinds |= _list_kinds_in_instr(code)
    return kinds


def _fld_char_kinds(paragraph: Paragraph) -> list[str]:
    """这一段里 `w:fldChar` 的类型序列（begin/separate/end），按文档顺序。"""
    out: list[str] = []
    for el in paragraph._p.iter(qn("w:fldChar")):
        out.append(el.get(qn("w:fldCharType")) or "")
    return out


def toc_field_span_indices(paragraphs: list[Paragraph],
                           parts: set[str] | None = None) -> set[int]:
    """与"TOC 家族域"的区间相交的段落下标（**含域渲染出来的条目段**）。

    `parts` 给了就只算这些部件的域（`{"目录"}` 只看目录域）；不给＝任何
    TOC 家族域（目录 + 三张清单）都算。

    **为什么需要它**：域是**跨段落**的——`fldChar begin` 在头一段、`end` 在
    最后一段，中间每一段都是域结果。而且 Word 的"自动目录"内容控件里，那些
    条目段**可能一个 `w:pStyle` 都不挂**（格式全靠控件 + 直接格式），
    `is_toc_paragraph()`（按 `toc N` 样式名认）一条都认不出来。用户 2026-09-16
    实拍 `民总.docx`：勾了"重新为文档生成目录"之后，成品里旧目录的 18 条
    条目文字原样留着，前面又插进去了新目录。同一批段落还会被
    `alerts.check_front_matter_body` 误报成"目录下多余的正文"。

    所以判据统一成"这一段与那个域的区间相交"：从头一段（含 begin 的那一段）
    一直算到 depth 归零的那一段。条目里嵌套的 HYPERLINK/PAGEREF 成对出现，
    不影响深度计数。

    已知边界（宁缺毋滥）：域指令与 `begin` 不在同一段（罕见写法）时判不出来，
    调用方仍可退回"按样式认条目"。

    ⚠️ **不闭合的域一律不算区间**（AAT 2026-09-17 实测的"回炉不收敛"）：
    稿子里最后一个 `end` 被弄丢之后，老写法会让深度永远不归零 ⇒ 余下**半篇**
    都被当成"域内、不参与识别" ⇒ 参考文献/致谢识别不到 ⇒ 补件逻辑每遍再补
    一遍（四遍 120→124→128）。改成"只有真的 begin…end 配上了才把它的区间
    收进来"，断了的那一段就当作普通内容看——宁可少删一份旧目录，也不能
    静默吞掉半篇。
    """
    marked: set[int] = set()
    # 每一层域：(开始段号, 这一层是不是要收的 TOC 家族域)。**只有闭合时**
    # 才把区间写进 `marked`——没弹出来的就丢在栈里不管。
    stack: list[tuple[int, bool]] = []
    for index, paragraph in enumerate(paragraphs):
        # 这一段自己是不是"要收的域"（指令就在这一段里）
        owned = has_toc_field(paragraph) and (
            parts is None or toc_field_part(paragraph) in parts)
        for kind in _fld_char_kinds(paragraph):
            if kind == "begin":
                stack.append((index, False))
            elif kind == "end":
                if stack:
                    start, is_toc = stack.pop()
                    if is_toc:
                        marked.update(range(start, index + 1))
        # 域还开着 ⇒ 把它那一层标上。
        # ⚠️ 域可以**自成一段**（begin 与 end 同段，比如空的目录域）：那时上面
        # 已经弹空了，再写 `stack[-1]` 会 IndexError（2026-09-16 实拍）。
        if owned and stack:
            start, _was_toc = stack[-1]
            stack[-1] = (start, True)
    return marked



def is_toc_paragraph(paragraph: Paragraph) -> bool:
    """目录条目（样式叫 `toc 1`/`toc 2`…）。

    公开名（不带下划线）是因为 `alerts.check_front_matter_body` 也要用它：
    判断"目录页里这一段是域生成的条目，还是用户手打的正文"。

    这一条很关键：标题里完全可能出现"表 1 各国立法比较"，那目录里就会有一条
    `表 1 各国立法比较……3`——它不以标点结尾、也不长，单看文字跟真表题一模一样。
    只有样式能区分它们。
    """
    return str(paragraph.style.name).strip().lower().startswith("toc ")


def is_list_entry_paragraph(paragraph: Paragraph) -> bool:
    """**图表目录/清单**的条目段落（Word 内置样式 `table of figures`）。

    跟 `is_toc_paragraph` 是一家人（都是域渲染出来的条目、不是用户写的正文），
    但**故意分成两个函数**：删旧域那一步（`cleanup.drop_regenerated_fields`）
    只该按"目录条目"删，不能顺手把清单条目一起删掉（只勾了重新生成目录时，
    清单的条目得留着）。

    用户 2026-09-16 实测（桌面 `TOC.docx`）：一个 `TOC \\h \\c "表"` 清单域的
    条目段，样式是 `table of figures`、文字是 `表 1 1`（条目 `表 1` + 页码 `1`）
    ——不认样式的话，它会被当标题、被当题注，还会被认成论文题目。
    """
    return str(paragraph.style.name).strip().lower() == "table of figures"


def has_seq_field(paragraph: Paragraph) -> bool:
    """段落里有没有 SEQ 域——有就说明编号已经是"真题注"的写法了。"""
    return _seq_label(paragraph) is not None


def typed_caption_number(text: str) -> tuple[str, str, int, int] | None:
    """题注文字里"手打"的那截编号——返回 (类别, 编号文字, 起, 止)，没有就 None。

    两种写法都认：中文行 `图 1 …`、英文行 `Fig 1 …`；编号藏在域里的不算。
    起止是**编号文字在 text 里的字符偏移**，调用方先把 text.strip() 再传，
    自己加回前导空白的偏移。

    注意：工具**不再改编号**（2026-09-16 删掉了整个题注化），这个函数现在只
    给诊断脚本用（打印"这条题注手里是哪一段编号"），生产代码里已经不调了。
    """
    m = _CAPTION_RE.match(text)
    if m is not None:
        return m.group(1), m.group(2), m.start(2), m.end(2)
    m = _EN_CAPTION_RE.match(text)
    if m is not None:
        kind = "图" if m.group(1).lower().startswith("fig") else "表"
        return kind, m.group(2), m.start(2), m.end(2)
    return None



def detect_document_title(
    document: Document,
    matches: list[HeadingMatch | None],
    front_matter: list[FrontMatterLabel],
    start: int = 0,
    skip: set[int] | None = None,
) -> TitleMatch | None:
    """大标题：必须出现在所有正则识别出的标题（不管是数字编号标题，还是
    引言/摘要这类固定名称标题/标签）之前——从头开始找第一个"不以标点结尾"
    的正文段落，一旦先碰到任何已识别的标题/标签，就说明标题区域已经开始了，
    前面没找到就是没有，不能跳过去继续往后面找。
    副标题：**紧跟大标题之后的那一段**，非空、不以标点结尾、本身不是已识别
    的标题/标签——**不跳过空段、也不要求以破折号开头**（用户 2026-09-15 的
    两条裁决，理由见下面那段注释）。

    `start`：从哪里开始找。默认 0（从文档头）；用户指定了封面范围时传
    `cover_end + 1`——封面上的"课程名称：____"这类栏位行不可以被当成论文
    题目，而光靠滤掉识别结果拦不住它（封面那几段本来就什么都识别不出来）。

    `skip`：**不可能当题目/副标题**的段落下标（调用方传"已判成题注的那几段"）。
    另外这三类也在内部直接跳过，不看调用方：目录/清单**域所在的段落**
    （它的文字是 Word 渲染的条目）、`toc N` 与 `table of figures` 样式的
    **条目段落**。用户 2026-09-16 实测（桌面 `TOC.docx`）：一个只有清单域的
    短文档，那个域段（文字 `表 1 1`）被认成了论文题目——因为它在文档最前面、
    不以标点结尾。跳过之后该文档正确地"没有题目"。
    """
    paragraphs = document.paragraphs
    front_matter_indices = {f.paragraph_index for f in front_matter}
    skip_indices = skip or set()

    def is_recognized(index: int) -> bool:
        return matches[index] is not None or index in front_matter_indices

    def is_plain_body(index: int) -> bool:
        return not is_recognized(index)

    def claimable(index: int) -> bool:
        """这一段能不能当题目：不是题注、不是目录/清单的域或条目。"""
        if index in skip_indices:
            return False
        paragraph = paragraphs[index]
        if has_toc_field(paragraph):
            return False
        return not (is_toc_paragraph(paragraph) or is_list_entry_paragraph(paragraph))

    title_index = None
    for i, paragraph in enumerate(paragraphs):
        if i < start:
            continue
        if is_recognized(i):
            # 撞到**任何**已识别的标题/标签（含「目录/清单」标签）就停：表 3.6
            # 的 R1 第二句写的是「如果提前识别到其他部件，将认定为标题不存在」，
            # 目录也是部件，没有例外。
            #
            # ⚠️ 这里原来有一条特例：目录/清单标签在最前面时跳过它、继续往下找
            # 题目（当时的理由是"目录在最前、题目在其后"的回炉稿会永远判成
            # 没有题目 ⇒ 每遍回炉补一条）。**用户 2026-09-17 裁决按表来**：
            # 「目录在前面就不认题目，按照表格来。如果谁确实把题目放目录后面，
            # 那说明他脑袋有坑。」回炉稳定不再靠这条特例——占位语自己认得出
            # （见 placeholders.py），补件端另有幂等兜底（`_find_slot_placeholder`）。
            break
        if not claimable(i):
            continue
        text = paragraph.text.strip()
        if not text:
            continue
        if is_title_placeholder(text):
            # 工具自己给"题目"留的坑：它就在题目该在的位置上——认它。
            # （它同样以「］」结尾，下一条"不以标点结尾"会把它挡掉 ⇒ 每遍重插）
            title_index = i
            break
        if text.endswith(tuple(_TRAILING_PUNCT)):
            continue
        title_index = i
        break

    if title_index is None:
        return None

    # 副标题（用户 2026-09-15 定的四条判据）：**紧跟题目之后的那一段**，
    #   * **不跳过空段**——题目下面空着一段，说明作者是有意把题名和后面那段
    #     分开的，那一段就不是副标题（原来的写法会跳过空段去找，等于无视这个
    #     信号）；
    #   * 非空；
    #   * **不以标点结尾**——带句末标点的是一句话，不是题名（这一条原来压根
    #     没有，是这次新加的）；
    #   * 不是已识别的标题/标签——否则"题目 / 【摘要】"会被读成"题目 + 副标题"。
    #
    # **不再要求以破折号开头**（原来要求 `—`/`-`）：规范只是允许用副标题，
    # 从没规定必须用破折号引出，把它当硬条件会把不带破折号的副标题一律漏掉。
    # 代价是题名下面的其它前置行（作者名、指导教师……）也可能被判成副标题——
    # 这是把标准放宽到"有规范的排版文件之外"的必然后果，交给用户在结构页里
    # 一个拖动改掉，比整类漏判好。
    subtitle_index = None
    if title_index + 1 < len(paragraphs):
        text = paragraphs[title_index + 1].text.strip()
        if text and claimable(title_index + 1) and is_plain_body(title_index + 1):
            if is_subtitle_placeholder(text):
                # 工具自己给"副标题"留的坑：它就在副标题该在的位置上——认它。
                # 不然回炉导出时会再插一条：占位语以「］」结尾，过不了下面那条
                # "不以标点结尾"（AAT 2026-09-17 实测：每遍 +1 段）。
                subtitle_index = title_index + 1
            elif not text.endswith(tuple(_TRAILING_PUNCT)):
                subtitle_index = title_index + 1
        elif not text:
            # 题目下面空着 ⇒ 按 2026-09-15 的裁决，这不算副标题（作者有意隔开）。
            # **唯一一个例外**：工具自己插的"副标题占位语"可能落在空段后面——
            # 补件时的插入点是"下一个已知部件"（这里是【摘要】），中间的空段
            # 被跳过了 ⇒ 占位语摆在那儿却认不出来 ⇒ 每遍回炉再插一条。
            # AAT 2026-09-17 实测（"首段 run 包进未接受的修订"、"body 内容整体
            # 复制"）：三条占位语叠在一起，每遍 +1 段。只有占位语穿过空段，
            # 真正的内容段落依旧不穿（裁决不变）。
            j = title_index + 1
            while j < len(paragraphs) and not paragraphs[j].text.strip():
                j += 1
            if (j < len(paragraphs)
                    and is_subtitle_placeholder(paragraphs[j].text.strip())
                    and claimable(j) and is_plain_body(j)):
                subtitle_index = j

    return TitleMatch(title_index, subtitle_index)
