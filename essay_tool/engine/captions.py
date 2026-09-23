"""图题/表题的识别（**只回答"这一段是不是图题/表题、是哪一种"**）。

2026-09-19 从 `detect.py` 搬出来，好让 `recognize.py` 在识别阶段**一次**把结果
写进索引（`Entry.meta["caption"]`），之后各处（结构页、告警、导出）都从索引读，
不再各自重新识别（此前 `build_rows`／`collect_source_alerts`／导出各扫一遍）。
`detect.py` 仍然 re-export 这里的公开名，老调用点不用改。

⚠️ 本模块**不许 import `detect`/`recognize`**（层次相反）。
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from essay_tool.engine.doc_index import (
    has_toc_field,
    is_list_entry_paragraph,
    is_toc_paragraph,
)

_TRAILING_PUNCT = "。！？；，、.!?;,"

# ------------------------------------------------------- 图题 / 表题
#
# 用户里有两种写法，两种都要兜住：
#   1. **正规做法**：Word「引用→插入题注」，段落用内置的「题注」样式，正文里
#      带一个 SEQ 域。这两个信号极强，几乎不会误判。
#   2. **手打**：图/表下面（或上面）自己敲一行"图 1 xxx"，没有任何标记。
#      只能靠文字特征猜——所以额外要求"不以标点结尾、且不过长"，因为
#      "图 1 显示了这个结果，……"这种明显是正文。
#
# **英文图题表题**（规范 [128][134]："如需英文图名，应中英文对照，英文图序
# 与图名另起一行放中文下方"）：字是 `Fig 1` 和 `Table 1`。它的弱信号比中文
# 危险得多——`Table 1 shows the results of the experiment` 是英文正文里极
# 常见的写法，长度和标点两条都不一定拦得住。所以英文弱信号额外要求**上一段
# 就是中文图题**（规范本来就规定它跟在中文下方）；强信号（SEQ 域/题注样式）
# 不受这条限制。
_CAPTION_RE = re.compile(r"^([图表])\s*(\d+(?:[.\-]\d+)*)")
# `Fig 1`、`Fig. 1`、`Figure 1`、`Table 1`……规范要求图序和数字之间空一格，
# 所以 `Fig1` 这种连在一起的写法不收（它更可能是变量名/型号）。
_EN_CAPTION_RE = re.compile(r"^(fig(?:ure)?|table)\.?\s+(\d+(?:[.\-]\d+)*)", re.IGNORECASE)
# "Fig/Table + 空格"，**不要求后面跟数字**——只给**强信号**段落判语种用：
# 强信号段落可能**已经没有可见编号**：用户自己用 Word 题注时数字在 SEQ 域里，
# 文字上只剩 `Fig `+说明，带数字的正则匹配不上（2026-09-16：我们不再替用户
# 写域，但用户原本就有的域要照认）。
_EN_CAPTION_SOFT_RE = re.compile(r"^(fig(?:ure)?|table)\.?\s+", re.IGNORECASE)
_CAPTION_MAX_LEN = 60


@dataclass
class CaptionMatch:
    """图题或表题。kind 是 "图"/"表"，决定套哪个样式。

    2026-09-16 大简化（用户口径）：**我们只管样式，编号与我们无关**。
    所以这里只剩两件事：**是什么种类**、**挂不挂样式**（＝进不进清单）：

    * 不区分“主图/分图”——用户原话：“只要能认出来是题注，区分主分是没有意义
      的，不管是主分，认进来就完事。”（`图 1（a）xxx`、`图 1a xxx` 一律当
      普通题注：挂样式、进清单，每张一条）
    * **不写域、不改编号、不重排、不告警**——手打和域混着用都行，域从今往后
      只是**识别信号**（有 `SEQ` 域 ⇒ 强信号）。

    `is_english`：文字是 `Fig 1 …` / `Table 1 …` 写法的"英文对照"行（规范
    [128][134] 要求英文图序另起一行放中文下方）。**这类行只写直接格式、不挂
    图题/表题样式**——挂了的话会被"插图清单/附表清单"域按样式收进去，一条图
    在中英对照的文档里会变成两条目（这在用户场景里是常态，不是边缘情况）。
    """

    paragraph_index: int
    kind: str
    text: str
    is_english: bool = False
    # 续图/续表行（`表 1（续）`）：**同一张图/表**（用户 2026-09-15：“续图压根
    # 是同一张图”）。它算题注，但**不挂图题/表题样式**——清单域是按样式名收
    # 条目的，一挂就会多出一条“表 1（续）”。外观靠直接格式跟图题/表题保持一致
    # （跟英文对照行一个做法）。判据不收紧（用户 2026-09-16：“续允许出现在
    # 任何地方，都算续”）：任意位置出现（续）/（续表）都算。
    is_continuation: bool = False

    @property
    def wears_caption_style(self) -> bool:
        r"""要不要挂**真样式**（图题/表题）。

        挂样式＝会被清单域（`TOC \h \z \t "图题,1"`）收进去。所以：
        中文图题/表题（含分图）挂；英文对照行、续行**不挂**（只写直接格式）。
        以上就是这个类的全部行为差异——不涉及编号。
        """
        return not (self.is_english or self.is_continuation)


def _seq_label(paragraph: Paragraph) -> str | None:
    """段落里的 SEQ 域用的什么标签（Word 题注的域长这样：`SEQ 图 \\* ARABIC`）。"""
    instrs = []
    for fld in paragraph._p.iter(qn("w:fldSimple")):
        instrs.append(fld.get(qn("w:instr")) or "")
    for instr in paragraph._p.iter(qn("w:instrText")):
        instrs.append(instr.text or "")
    for instr in instrs:
        m = re.search(r"SEQ\s+(\S+)", instr)
        if m:
            return m.group(1)
    return None


def _is_caption_style(paragraph: Paragraph) -> bool:
    style = paragraph.style
    return style is not None and str(style.name).strip().lower() in ("caption", "题注")


def _looks_like_body(text: str) -> bool:
    """明显是正文，不该被当成图题/表题。

    只两条，但足够：以句子标点结尾、或者长得不像题名。思路跟 `_classify()`
    里那套标题判断是同一个（"标点结尾静默判正文"+"过长当异常信号"）。
    """
    return text.endswith(tuple(_TRAILING_PUNCT)) or len(text) > _CAPTION_MAX_LEN


# 续行（`表 1（续）`）的判据（用户 2026-09-16：“续允许出现在任何地方，都算续”）。
#
# 续行：**算题注、但不挂样式**——用户原话：“续图和续表只识别出表题，一律不
# 计列表”。不挂图题/表题样式 ⇒ 清单域（按样式名收）不会多出“表 1（续）”这一条；
# 外观用直接格式跟图题保持一致。判据不锁位置（“表 1（续）”“表 1 表名（续）”都认）。
#
# **分图不再是一个概念**（用户 2026-09-16：“只要能认出来是题注，区分主分是没有
# 意义的，不管是主分，认进来就完事”）——`图 1（a）xxx`、`图 1a xxx` 一律当
# 普通题注：挂样式、进插图清单，每张一条。
_CONTINUATION_RE = re.compile(r"[（(]\s*续[表图]?\s*[）)]")
# 2026-09-20 补（用户裁决"回去改代码认续表"）：不带括号的写法也认——
#   * 开头就是"续表 1 …""续图 1 …"（这种不以"图/表"开头，以前根本不算题注）；
#   * 数字后面紧跟"续表/续图"（"表 1 续表"）。
_CONT_PREFIX_RE = re.compile(r"^续\s*([图表])\s*\d+(?:[.\-]\d+)*")
_CONT_AFTER_NUMBER_RE = re.compile(r"^[图表]\s*\d+(?:[.\-]\d+)*\s*续[图表]")


def _is_continuation_line(text: str) -> bool:
    """续图/续表行——同一个图/表的续页题注（算题注，但不挂样式）。

    三种写法都算：
      * "（续）""(续)""（续表）""（续图）"：**任意位置**出现就命中（括号全/半角、
        括号内带空格都认）——用户 2026-09-16 定的口径就是"出现在任何地方都算续"
        （规范给的写法是"表1（续）"，表题夹在中间的"表 1 表名（续）"也常见）；
      * 开头是"续表 1""续图 1"；
      * 数字后紧跟"续表""续图"（"表 1 续表"）。

    已知边界（宁缺毋滥）：把"（续）"当正文括号用的（"图 1 流程（续）说明"）
    会被当成续行 ⇒ **不进插图清单**（代价只是清单少一条，不涉及编号了）。
    """
    return (_CONTINUATION_RE.search(text) is not None
            or _CONT_PREFIX_RE.match(text) is not None
            or _CONT_AFTER_NUMBER_RE.match(text) is not None)


def _follows_chinese_caption(paragraphs: list[Paragraph], index: int,
                             chinese_indices: set[int]) -> bool:
    """上一段（跳过空行）是不是刚认出来的中文图题。

    专门用来拦英文正文：`Table 1 shows the results of …` 这种写法在英文论文里
    到处都是，光靠"不以标点结尾、不过长"拦不住短句。而规范本来就把英文图题
    规定成"另起一行放中文下方"，所以"上面就是中文图题"是个又稳又便宜的条件。
    """
    for previous in range(index - 1, -1, -1):
        if not paragraphs[previous].text.strip():
            continue  # 中间夹了空行也认
        return previous in chinese_indices
    return False


def _kind_from_seq_label(label: str) -> str | None:
    """从 SEQ 域的标签反推这是图题还是表题。

    **精确匹配**（用户 2026-09-15：“只管图、表，不管图表和公式，这两类排除
    出去”）：标签必须正好是 `图`/`表`（英文 `Fig`/`Figure`/`Table`）。
    以前是“含‘图’就算图” ⇒ `SEQ 图表` 的题注被算成图题、混进插图清单；
    `SEQ 公式` 则完全没人管。现在两类都归“不认识”——不套样式、不进清单、
    不动它原来的文字/样式。
    """
    label = label.strip().lower()
    if label in ("图", "fig", "figure"):
        return "图"
    if label in ("表", "table"):
        return "表"
    return None


def detect_captions(
    document,
    heading_indices: set[int] | None = None,
) -> list[CaptionMatch]:
    """老入口（按文档）：诊断脚本/AAT 仍在用；识别本体见 `detect_captions_in`。"""
    return detect_captions_in(document.paragraphs, heading_indices)


def detect_captions_in(
    paragraphs: list[Paragraph],
    heading_indices: set[int] | None = None,
) -> list[CaptionMatch]:
    """识别图题/表题。

    强信号（SEQ 域 / 题注样式）和弱信号（文字以"图/表 + 数字"开头）都认。

    两条门限的顺序（用户 2026-09-17）：**先看是不是题注（强信号）、再看符不
    符合前缀**，两条都符合就**豁免末尾标点**；其余情况照旧要过"明显是正文"
    的排除（题注样式也可能被人拿去干别的，手打段落更容易是正文）。长度上限
    （`_CAPTION_MAX_LEN`）对谁都适用。宁可漏判，也不要把正文改成图题。

    英文图题（`Fig 1 …`、`Table 1 …`）的弱信号还要额外满足"上一段是中文图题"，
    理由见文件上面那段注释。

    `heading_indices`：已经判成标题的段落下标，它们不可能同时是图题。
    """
    skip = heading_indices or set()
    found: list[CaptionMatch] = []
    chinese_indices: set[int] = set()

    for i, paragraph in enumerate(paragraphs):
        if i in skip or is_toc_paragraph(paragraph) or is_list_entry_paragraph(paragraph):
            continue
        # 目录/清单**域所在的段落**同样跳过：它的文字是 Word 渲染出来的条目
        # （`表 1 1`），单看文字跟真题注一模一样（用户 2026-09-16：`TOC.docx`
        # 里那行被套成了表题）。注意只能拦 TOC 家族（`has_toc_field`），
        # **不能拦所有域**——题注自己就常带 SEQ 域。
        if has_toc_field(paragraph):
            continue
        text = paragraph.text.strip()
        if not text:
            continue

        continuation = _is_continuation_line(text)
        strong = _seq_label(paragraph) is not None or _is_caption_style(paragraph)
        # `续表 1 …`/`续图 1 …` 不以"图/表"开头，靠 `_CONT_PREFIX_RE` 认（group(1) 同样是种类）
        chinese = _CAPTION_RE.match(text) or _CONT_PREFIX_RE.match(text)
        english = None if chinese is not None else _EN_CAPTION_RE.match(text)
        # 用户 2026-09-17 裁决："先看是不是题注，然后看符不符合前缀；如果这两条
        # 都符合，允许豁免末尾标点规则。"
        #   * 是不是题注 ＝ 强信号（SEQ 域 / 题注样式）；
        #   * 符不符合前缀 ＝ 文字开头是 `图 1 `/`表 1 `/`Fig 1 `/`Table 1 `
        #     （强信号段自己把编号放在域里的那种"软前缀"也算）。
        # 两条都符合 ⇒ **不再**因为"以句末标点结尾"被否掉：`图 3 …（含案例
        # 说明）。` 这种题注作者自己加了句号，身份是域/样式给定的，不该被标点
        # 推翻（否则清单里少一条、旁边那张图还会被多报一条"未识别到题注"）。
        # 长度那一半（> `_CAPTION_MAX_LEN`）**不豁免**：用户只说了末尾标点。
        strong_prefix = strong and (
            chinese is not None or english is not None
            or _EN_CAPTION_SOFT_RE.match(text) is not None
        )
        if not strong_prefix and _looks_like_body(text):
            continue
        if len(text) > _CAPTION_MAX_LEN:
            continue

        kind: str | None
        # 续行（`表 1（续）`）：文字里只有“表/图 + 数字”才能定类（“（续表）”
        # 开头的那种没有类别，退回看 SEQ 标签）。
        if chinese is not None:
            kind = chinese.group(1)
        elif strong:
            # 题注样式/SEQ 域在，但文字开头不像"图 1"——那就看 SEQ 用的什么标签
            kind = _kind_from_seq_label(_seq_label(paragraph) or "")
        elif english is not None and _follows_chinese_caption(paragraphs, i, chinese_indices):
            kind = "图" if english.group(1).lower().startswith("fig") else "表"
        else:
            continue

        if kind is None:
            continue
        # 英文行的判定：常规看文字（带数字）；用户自己在英文行里放了域的那种，
        # 文字上只剩 `Fig / Table + 空格`，强信号段落走"软"判定（见 _EN_CAPTION_SOFT_RE）。
        is_english = chinese is None and (
            english is not None or (strong and _EN_CAPTION_SOFT_RE.match(text) is not None)
        )
        # 英文对照行与续行**不挂样式**（⇒ 不进清单），见 CaptionMatch 的说明；
        # 其余（中文图题/表题，含分图、含 `图 1a` 这种）一律挂样式、进清单。
        # **不区分主/分**，也不碰编号（用户 2026-09-16 定的口径）。
        found.append(CaptionMatch(i, kind, text, is_english=is_english,
                                  is_continuation=continuation))
        if not is_english:
            # "中文图题"的集合（英文弱信号拿它查"上一段是不是中文图题"）。
            # 这里的"中文行"是广义的——按"不是英文行"收集，不能只收文字正则
            # 匹配上的（用户自己在中文行里放了域时，文字上没数字。）。
            chinese_indices.add(i)
    return found



# ------------------------------------------------ 索引 ⇄ 题注（识别一次，到处读）
#
# 识别结果放在 `Entry.meta["caption"] = {"kind": "图"/"表", "role": "main"/"cont"/"en"}`。
# 结构页勾选"图/表/续"时由 `doc_index.build_ops()` 写回同一个键；告警与导出
# 只读它，不再自己重新识别。

def caption_role(caption: CaptionMatch) -> str:
    """题注角色：`en`＝英文对照行、`cont`＝续行、`main`＝其余（挂样式、进清单）。"""
    if caption.is_english:
        return "en"
    if caption.is_continuation:
        return "cont"
    return "main"


def caption_meta(caption: CaptionMatch) -> dict:
    return {"kind": caption.kind, "role": caption_role(caption)}


def captions_from_index(index) -> list[CaptionMatch]:
    """从（可能被结构页改过的）索引里读出题注列表，`paragraph_index` 是**段落序号**。"""
    from essay_tool.engine.doc_index import paragraph_ordinals

    found: list[CaptionMatch] = []
    for ordinal, entry in sorted(paragraph_ordinals(index).items()):
        meta = entry.meta.get("caption")
        if not meta:
            continue
        role = meta.get("role") or "main"
        found.append(CaptionMatch(
            ordinal, meta["kind"], entry.text.strip(),
            is_english=role == "en", is_continuation=role == "cont"))
    return found
