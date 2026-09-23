"""识别层重写（2026-09-18 起）：标签式双向扫描，直接把结果填进索引法的 Entry。

## 背景（别再走回头路，理由都在这）

旧的 `detect.py` 是十几个各自独立扫描 `document.paragraphs` 的探测器
（`detect_headings`/`detect_front_matter_labels`/`detect_document_title`/
`detect_english_title`/……），互相把结果当参数传来传去，同一套依赖链在
`pipeline.recognize_source()`、`adjust_preview.build_rows()`、
`structure.build_structure_inventory()` 三处被独立复述——这是这次重写要
解决的根本问题。**本模块不 import `detect.py` 的任何函数**，是一次从
2026-09-18 讨论定下的字面逻辑重新搭建，只挪了两类"纯算法、体量小"的
东西：`numbering.py` 的 Word 自动编号反算（不是探测器架构的一部分），
以及正则本身验证过的细节（"1.4 不算三级标题"的负向断言、百/千不算合法
首位……）。

## 架构：前向 + 后向两个方向扫描器，中间夹出目录

- **前向识别器**（`_scan_forward`，2026-09-19 重写）：中文摘要/关键词、
  英文 Abstract/Keywords 这四个"纯字面标签"，加上英文题目/副标题，
  **五件事全部在同一段范围 `[start, 前置区终点)` 里各自独立扫描，互不
  连坐**（老版本"英文题目从摘要/关键词认到的位置往后找、英文摘要/关键词
  又从英文题目认到的位置往后找"这种链式起点，被 AAT 实测出"关键词一挪到
  文档末尾，前面一整片（英文题目/英文副标题/Abstract/Keywords）全跟着
  漏认"的连坐 bug，这次连根拔掉）：
  - **前置区终点**（`_front_matter_boundary`）：从 `start` 往后扫，第一次
    撞见"引言家族/结论/参考文献/附录/致谢/目录/三张清单/像正文标题的
    编号行"这六类东西中任意一个的位置，就是终点——这条边界只跟"正文/
    后置区从哪开始"有关，不受下面这五件事自己找没找到、找在哪影响。
  - **四个字面标签独立扫描**：摘要/关键词/Abstract/Keywords 各自在
    `[start, 终点)` 里逐段核对"是不是恰好等于这个标签"，谁先扫都一样，
    互不依赖对方的结果。认**第一次出现**的位置为准（决定"标签+内容
    两段式"怎么划）；同一个标签在范围内又出现一次，识别到但不采信第
    二次的位置——位置记进 canonical 那条 `Entry.meta
    ["duplicate_label_positions"]`，交给 `alerts.
    check_duplicate_front_labels()` 报一句"重复的部件"（2026-09-19
    用户拍板：这种重复不该再是静默的）。
  - **英文题目/副标题：靠内容形状认，不靠位置**（2026-09-19 用户明确
    要求——"永远不要靠位置来锚定，把英文标题连着写在中文标题后是很正常
    的"）。同样在 `[start, 终点)` 里找**第一段**满足 `_looks_like_english()`
    的段落当英文题目，**紧接着的下一段**（必须相邻）满足（门槛更低）就是
    英文副标题。排除掉上面四个标签自己认到的那几段，以及目录/清单域、
    题注样式段落。已知缺口（2026-09-19 用户拍板：不管）——Abstract 的
    **内容**（不只是标签本身）如果被挪到比英文题目还靠前、内容里又恰好
    有一行不带句末标点的短句，仍可能被误认成英文题目，因为这一步排除
    得掉 Abstract **标签**那一段、排不掉它管辖的**内容**（内容归属要等
    "标签+内容两段式"扫描才算得出来）。用户原话："用户不给摘要写标点是
    他自己的事"——按下不表，不做两遍扫描。
- **后向识别器**（`_scan_backward`）：从文档尾往前走，同时认两件事——
  "引言"家族标签（引言/绪论/序言/前言/导论/导言/引论/绪言——2026-09-18 起绪论
  类词汇正式解禁，跟"引言"同等对待，这是对 2026-09-15 老裁决的明确翻转；
  点分体系下前面允许带编号，见 `HeadingGrammar.split_anchor_number`）和
  后置标签（结论/参考文献/附录/致谢），以及正文标题（一~四级）。
  **标签会回收它"管辖范围"内被误认的正文标题**：参考文献的条目常常长得
  像"1. 张三：……"这种三级标题，扫描时会先被暂存成候选，一旦扫到
  "参考文献"这个标签本身，暂存的这批全部打回正文——这是旧架构完全没有
  的保护，也是这次重写的主要收益之一。
  停止条件（满足其一就收工，见 `_scan_backward` 内部注释对应关系）：
    ① 认到"引言"家族标签；
    ② 认到前向识别器已经认过的部件、目录标签、目录域；
    ③ 扫到文档头；
    ④ 连续两段命中"已认出的正文标题列表"里某一条的汉字部分——这是
      抓"手打目录条目"的最后一道防线（真标题在文档靠后的正文区已经先被
      认出、进了列表；扫到前面的目录页时,重复的条目文字会撞上列表）。
      命中就撤销这两段的标题身份，扫描停在这一步，不再倒着往前找。
- **目录 + 三张清单是"夹逼"出来的**（`_scan_gap`）：题目和目录都没有
  专门的探测器认它们所在的具体是"目录"还是别的——前向识别器认到英文
  关键词（或放弃）为止，后向识别器认到"引言"（或目录标签本身）为止，
  两者中间夹出的那一小段范围里，用最简单的字面标签匹配去找"目录"/
  "插图清单"/"附表清单"/"插图和附表清单"，不需要方向性扫描那一套。
- **题目/副标题不是标签，是"游离正文"**（2026-09-18 用户口径）：在
  **所有**已经认出来的标签/标题之前，如果还有非空的正文段落，那就是
  游离正文——第一行是题目、第二行是副标题，不再套"是否以标点结尾"
  这类逐行判据。

## 标题编号方案：中文 / 点分（2026-09-19 起两套都有，不认混合写法）

`HeadingGrammar` 是标题编号方案的可插拔协议，两个扫描器只通过它认标题：

- `ChineseHeadingGrammar`（默认）：一、／（一）／1.／（1），四级都规范化；
- `ArabicDotHeadingGrammar`：`1` / `1.1` / `1.1.1` / `1.1.1.1` 对应一到四级，
  另外 `（1）` 也算四级。**只有 `（1）` 做规范化**（半角括号→全角、编号后
  一个空格）；点分写法一律原样保留——不转成中文编号，也不把 `1。1`
  这类不规范写法"纠正"成 `1.1`（不规范的就是不认，走"疑似标题"提醒）。

用哪一套由用户在界面上明确选（`grammar_for()`，见主窗口导入区的单选），
识别层不去猜。

"任何情况下不修改用户的文字内容，规范化除外"（用户 2026-09-19）：`canonical`
（规范写法）是识别那一刻按**识别时的层级**算好、定死的，之后用户把标题
拖到别的级别只改样式、编号文字不动。`canonical is None` ＝ 这一段不规范化。

点分体系里，带编号的"引言家族／结论"（`1 引言`、`1. 引言`、`1.引言`）仍然
认作引言／结论锚点（`Entry.meta["keep_text"]` 标记：编号是用户的文字，
套版时整段原样保留，不改写成"引　言"）。

## 识别层和索引法怎么接上（2026-09-18 用户明确要求）

`recognize()` 直接消费 `doc_index.build_index()` 产出的 `Entry` 列表，
把识别结果写进每条 `Entry.part`（部件名——沿用预设表/样式表已经在用的
那套字符串，比如"摘要""一级标题""引言"）和 `Entry.meta`（层级/编号/
自动编号固化文字/标签是"标签位"还是"内容位"……）。这样"调整识别结果"
窗口和未来的加工层（索引法执行器）读的是**同一份对象**，不用再各转
一遍格式。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Protocol

from docx.oxml.ns import qn

from essay_tool.engine.doc_index import (
    Entry,
    build_index,
    has_toc_field,
    is_list_entry_paragraph,
    is_toc_paragraph,
    paragraph_ordinals,
    toc_field_span_indices,
)
from essay_tool.engine.captions import caption_meta, detect_captions_in
from essay_tool.engine.numbering import NumberingScheme, load_numbering_scheme
from essay_tool.engine.placeholders import (
    is_english_subtitle_placeholder,
    is_english_title_placeholder,
    is_subtitle_placeholder,
)

# ============================================================ 标签匹配

# 中文标签：允许左右各一个方括号（不要求成对），字与字之间允许半角/全角空格
# （2026-09-17 老决定："认摘要的空格"，这里对所有中文字面标签一视同仁）。
def _cjk_label_pattern(word: str) -> re.Pattern[str]:
    sep = r"[ 　]*"
    body = sep.join(re.escape(c) for c in word)
    return re.compile(rf"^[【\[]?{body}[】\]]?$")


def _ascii_label_pattern(word: str) -> re.Pattern[str]:
    return re.compile(rf"^[【\[]?{re.escape(word)}[】\]]?$", re.IGNORECASE)


def _english_keywords_pattern() -> re.Pattern[str]:
    """英文关键词标签：忽略大小写、忽略 key 与 word 之间的空格、单复数不分
    （`Keywords` / `Key Words` / `KEY WORDS` / `key word` 一律认；2026-09-19
    用户裁决）。"""
    return re.compile(r"^[【\[]?key[ 　]*words?[】\]]?$", re.IGNORECASE)


# 前向识别器地盘：中文摘要/关键词标签。
_FRONT_LABEL_WORDS = ("摘要", "关键词")
_FRONT_LABELS = {w: _cjk_label_pattern(w) for w in _FRONT_LABEL_WORDS}

# 前向识别器地盘：英文摘要/关键词标签（大小写不敏感；关键词另认 key/word
# 之间的空格与单复数，见 `_english_keywords_pattern`）。
_ENGLISH_LABEL_WORDS = ("Abstract", "Keywords")
_ENGLISH_LABELS = {
    "Abstract": _ascii_label_pattern("Abstract"),
    "Keywords": _english_keywords_pattern(),
}

# 夹逼区地盘：目录 + 三张清单，谁都不专门认，靠前后夹出范围后一次性扫。
_GAP_LABEL_WORDS = ("目录", "插图清单", "附表清单", "插图和附表清单")
_GAP_LABELS = {w: _cjk_label_pattern(w) for w in _GAP_LABEL_WORDS}

# 后向识别器地盘之一："引言"家族——2026-09-18 起绪论/序言/前言/导论跟引言
# 同等对待（翻转 2026-09-15 的老裁决），命中任意一个都当停止条件①，统一
# 落成"引言"这个部件名（下游只认这一个部件，不分叉；但**文字**不改——用户
# 写的是"绪论"，成品里仍是"绪论"，只按预设补空格，见 `normalize`）。2026-09-19
# 补了导言/引论/绪言（用户："绪论、导言、前言、引论这些字样都算""绪言那一大
# 家子都认"）。
INTRO_WORDS = ("引言", "绪论", "序言", "前言", "导论", "导言", "引论", "绪言")
_INTRO_LABELS = {w: _cjk_label_pattern(w) for w in INTRO_WORDS}

# 后向识别器地盘之二：后置标签，命中会回收它管辖范围内暂存的正文标题。
# 每个后置标签全篇只认一次（2026-09-19 改：附录不再允许重复——"附录A"
# "附录B"是附录这一整块内部的分项，不是第二个"附录"标签）。同一词后续再
# 命中不新开一段，但仍触发"回收 pending 正文标题"这个副作用，安全起见。
_BACK_LABEL_WORDS = ("结论", "参考文献", "附录", "致谢")
_BACK_LABELS = {w: _cjk_label_pattern(w) for w in _BACK_LABEL_WORDS}
# 后置标签里允许带编号前缀的（点分体系：`5 结论`）——跟引言家族一样是"彻头
# 彻尾的一级标题、只是样式特殊"；参考文献/附录/致谢规范里本来就不编号。
_NUMBERABLE_BACK_WORDS = frozenset({"结论"})


# ============================================================ Word 自动编号

def _num_ref(paragraph) -> tuple[str, int] | None:
    """这一段挂的自动编号 (numId, ilvl)，没有就 None。"""
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


def _effective_text(paragraph, numbering: NumberingScheme) -> tuple[str, bool]:
    """段落"看起来"是什么文字：Word 自动编号时把渲染出的编号拼到文字前面。

    第二个返回值＝这个编号是不是只存在于 Word 渲染里（没有就是段落文字
    自己带的），后续固化编号文字时要用。
    """
    text = paragraph.text.strip()
    ref = _num_ref(paragraph)
    if ref is None:
        return text, False
    marker = numbering.render(*ref)
    if not marker:
        return text, False
    return f"{marker}{text}", True


# ============================================================ 标题编号方案（可插拔）

_CN_DIGITS = {"零": 0, "一": 1, "二": 2, "三": 3, "四": 4,
             "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
_CN_UNITS = {"十": 10, "百": 100, "千": 1000}
_CN_NUMERAL = "".join(_CN_DIGITS) + "".join(_CN_UNITS)
# "百二十一""千一百"这类省略首位的写法不算编号（沿用 2026-09-15 判据）；
# 廿/卅永远不进编号字样，不在 _CN_NUMERAL 里，天然认不出（2026-09-17 判据）。
_CN_FORBIDDEN_INITIAL = {"百", "千"}

_L1_RE = re.compile(rf"^([{_CN_NUMERAL}]+)、(.*)$")
_L2_RE = re.compile(rf"^[（(]([{_CN_NUMERAL}]+)[）)]\s*(.*)$")
# "不紧接数字"防"1.4"被三级标题正则截成"标题 1、内容 4"——沿用验证过的写法。
_L3_RE = re.compile(r"^(\d+)[.．、](?!\d)\s*(.*)$")
_L4_RE = re.compile(r"^[（(](\d+)[）)]\s*(.*)$")

_TRAILING_PUNCT = "。！？；，、.!?;,"
_TITLE_MAX_LEN = 60      # 用户 2026-09-20：50 → 60；超过的归正文，同时报"疑似标题"

LEVEL_PART_NAMES = {1: "一级标题", 2: "二级标题", 3: "三级标题", 4: "四级标题"}


def _squeeze_len(text: str) -> int:
    return len(re.sub(r"\s|　", "", text))


def _chinese_numeral_to_int(text: str) -> int | None:
    if not text or any(c not in _CN_DIGITS and c not in _CN_UNITS for c in text):
        return None
    if text[0] in _CN_FORBIDDEN_INITIAL:
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


@dataclass
class HeadingCandidate:
    """一次标题匹配的结果——跟具体用哪套编号方案（中文/阿拉伯）无关。"""

    level: int
    number: int | tuple[int, ...] | None   # 中文体系是整数；点分体系是各段组成的元组
    marker: str      # 编号文字（中文体系："三"/"1"；点分体系：段落里原样写的那一截，如 "1.1."）
    title: str
    # 规范写法的完整编号前缀（固化自动编号、规范化手打编号时要写的那一截）。
    # None ＝ 这一种编号不规范化，用户的文字原样保留（点分体系的 1.1.1 之类）。
    canonical: str | None


class HeadingGrammar(Protocol):
    """标题编号方案的可插拔接口（当前两套实现：中文、点分）。两个扫描器只
    通过它认标题，换方案在 `recognize()` 调用处传另一个实例即可。"""

    name: str

    def match(self, text: str) -> HeadingCandidate | None: ...

    def overlong(self, text: str) -> bool:
        """只因为超过 `_TITLE_MAX_LEN` 才没被认成标题（其余条件都满足）。"""
        ...

    def split_anchor_number(self, text: str) -> tuple[int | None, str]:
        """引言家族/结论这类"固定名称锚点"前面允许的编号前缀：返回
        `(编号, 去掉前缀后的文字)`；这一套方案里锚点不带编号（或者压根没有
        前缀）就返回 `(None, text)`。"""
        ...


def _match_paren_level4(text: str, check_length: bool = True) -> HeadingCandidate | None:
    """`（1）标题`：两套方案共用的四级写法，**规范化**（全角括号、编号后
    一个半角空格；全角数字/`01` 这类写法按解析出的整数重排）。

    `check_length=False`：不看长度上限（只给 `overlong()` 判"是不是只因为太长
    才没被认成标题"用）。"""
    m = _L4_RE.match(text)
    if not m:
        return None
    raw, title = m.group(1), m.group(2).strip()
    number = int(raw)
    if title.endswith(tuple(_TRAILING_PUNCT)):
        return None  # 标点结尾，静默归正文
    if check_length and _squeeze_len(text) > _TITLE_MAX_LEN:
        return None  # 过长，归正文（并由 `overlong()` 记下来告警）
    return HeadingCandidate(level=4, number=number, marker=raw, title=title,
                            canonical=f"（{number}） ")


class ChineseHeadingGrammar:
    """中文数字标题：一、/（一）/1./（1）。固定名称锚点（引言/结论……）
    不算在内——那些是"标签"，走 `_INTRO_LABELS`/`_BACK_LABELS`，不走这里。
    """

    name = "chinese"

    def split_anchor_number(self, text: str) -> tuple[int | None, str]:
        return None, text

    def match(self, text: str) -> HeadingCandidate | None:
        return self._match(text, True)

    def overlong(self, text: str) -> bool:
        """这段是不是**只因为太长**才没被认成标题（三、四级；一、二级没有长度限制）。"""
        candidate = self._match(text, False)
        return (candidate is not None and candidate.level in (3, 4)
                and _squeeze_len(text) > _TITLE_MAX_LEN)

    def _match(self, text: str, check_length: bool) -> HeadingCandidate | None:
        for level, pattern in ((1, _L1_RE), (2, _L2_RE), (3, _L3_RE)):
            m = pattern.match(text)
            if not m:
                continue
            raw, title = m.group(1), m.group(2).strip()
            number = _chinese_numeral_to_int(raw) if level in (1, 2) else int(raw)
            if number is None:
                continue
            if title.endswith(tuple(_TRAILING_PUNCT)):
                return None  # 标点结尾，静默归正文
            if check_length and level == 3 and _squeeze_len(text) > _TITLE_MAX_LEN:
                return None  # 过长，归正文（并由 `overlong()` 记下来告警）
            # 三级用**解析出的整数**重新渲染（不是原样照抄 raw）：这样
            # 全角数字、"01"这类写法在固化编号文字时会被规整成标准阿拉伯
            # 数字。
            canonical = {
                1: f"{raw}、",
                2: f"（{raw}） ",
                3: f"{number}. ",
            }[level]
            return HeadingCandidate(level=level, number=number, marker=raw, title=title,
                                    canonical=canonical)
        return _match_paren_level4(text, check_length)


# 点分标题：`1` / `1.1` / `1.1.1` / `1.1.1.1`（一到四级）。**只认半角数字、半角
# 点**——全角点、`1。1` 这类不规范写法不认（用户 2026-09-19："不允许把 1。1
# 规范化为 1.1"，不规范化就只能不认，走"疑似标题"提醒）。编号后面：空格，或
# 直接跟汉字/字母（`1.引言`）；直接跟数字/符号不行，防"3.5%""12.5米"这类
# 行首的小数被当成标题。一级只认 1~2 位数：防"2020 年……"这类行首的年份。
_DOT_HEADING_RE = re.compile(
    r"^(?P<num>[0-9]{1,2}(?:\.[0-9]{1,2}){0,3})\.?"
    r"(?:\s+|(?=[A-Za-z一-鿿]))(?P<title>.+)$")
# 锚点（引言家族/结论）前的编号：`1 引言` / `1. 引言` / `1.引言`。
_DOT_ANCHOR_NUMBER_RE = re.compile(r"^(?P<num>[0-9]{1,2})(?:\.\s*|\s+)(?P<rest>.+)$")


class ArabicDotHeadingGrammar:
    """点分数字标题：`1`/`1.1`/`1.1.1`/`1.1.1.1` 是一到四级，`（1）` 也是四级。

    **规范化只对 `（1）` 做**（跟中文体系同一套写法）；点分写法的
    `canonical` 是 None——用户的文字原样保留，既不转成中文编号，也不把
    `1。1` 之类"纠正"成 `1.1`。判据主要是**不以标点结尾**（沿用中文体系那
    一条），再加标题总长不超过 `_TITLE_MAX_LEN`。
    """

    name = "dotted"

    def split_anchor_number(self, text: str) -> tuple[int | None, str]:
        m = _DOT_ANCHOR_NUMBER_RE.match(text)
        if not m:
            return None, text
        return int(m.group("num")), m.group("rest").strip()

    def match(self, text: str) -> HeadingCandidate | None:
        return self._match(text, True)

    def overlong(self, text: str) -> bool:
        """这段是不是**只因为太长**才没被认成标题（点分体系一至四级都有长度限制）。"""
        return self._match(text, False) is not None and _squeeze_len(text) > _TITLE_MAX_LEN

    def _match(self, text: str, check_length: bool) -> HeadingCandidate | None:
        m = _DOT_HEADING_RE.match(text)
        if not m:
            return _match_paren_level4(text, check_length)
        title = m.group("title").strip()
        if not title or title.endswith(tuple(_TRAILING_PUNCT)):
            return None  # 标点结尾，静默归正文
        if check_length and _squeeze_len(text) > _TITLE_MAX_LEN:
            return None  # 过长，归正文（并由 `overlong()` 记下来告警）
        parts = tuple(int(p) for p in m.group("num").split("."))
        return HeadingCandidate(level=len(parts), number=parts,
                                marker=text[:m.start("title")].strip(), title=title,
                                canonical=None)


HEADING_SCHEMES: dict[str, type] = {
    "chinese": ChineseHeadingGrammar,
    "dotted": ArabicDotHeadingGrammar,
}
DEFAULT_HEADING_SCHEME = "chinese"


def grammar_for(scheme: str | None) -> HeadingGrammar:
    """按界面上选的方案名拿文法实例；没选（None）走默认的中文方案。"""
    return HEADING_SCHEMES[scheme or DEFAULT_HEADING_SCHEME]()


# ============================================================ 英文题名判据

_EN_TITLE_MAX_LEN = 120
_EN_TITLE_MIN_LETTERS = 5
_EN_SUBTITLE_MIN_LETTERS = 1


def _looks_like_english(text: str, min_letters: int) -> bool:
    """一段文字像不像英文题名：不含汉字、不以句末标点结尾、至少 min_letters
    个拉丁字母、长度不超过粗筛上限。不看是否全大写——规范要求全大写，但
    用户可能不管这条，不能拿大小写当识别依据。"""
    if not text or len(text) > _EN_TITLE_MAX_LEN:
        return False
    if text.endswith(tuple(_TRAILING_PUNCT)):
        return False
    if any("一" <= ch <= "鿿" for ch in text):
        return False
    return sum(1 for ch in text if ch.isascii() and ch.isalpha()) >= min_letters


# ============================================================ 目录域 / 样式识别（结构性事实，非判断）
#
# `has_toc_field`/`is_toc_paragraph`/`is_list_entry_paragraph` 直接复用
# `doc_index.py` 那一份（2026-09-18 起唯一实现），不再各自维护一份——
# `_is_toc_style()` 把后两者合成一个判据，跟旧 detect.py 里的等价写法一致。


def _is_toc_style(paragraph) -> bool:
    return is_toc_paragraph(paragraph) or is_list_entry_paragraph(paragraph)


_CAPTION_STYLE_NAMES = {"caption", "题注", "图题", "表题"}


def _is_caption_style(paragraph) -> bool:
    """题注类样式——不是本轮要重做的识别范围，但游离正文（题目/副标题）
    候选段落里绝不能把题注段落算进去（`TOC.docx` 实测：清单域旁边那条
    "表 1" Caption 样式段落，不排除的话会被当成副标题）。"""
    return str(paragraph.style.name).strip().lower() in _CAPTION_STYLE_NAMES


def _cjk_only(text: str) -> str:
    """只保留汉字部分（含中文数字，如"一""二"）——停止条件④比对用。"""
    return "".join(ch for ch in text if "一" <= ch <= "鿿")


# ============================================================ 前向识别器

def _front_matter_boundary(paragraphs, texts, grammar: HeadingGrammar, *,
                           start: int, end: int) -> int:
    """前置区终点：从 `start` 往后扫，第一次撞见"引言家族/结论/参考文献/
    附录/致谢/目录/三张清单/像正文标题的编号行"这六类东西中任意一个的
    位置，就是终点（不含该位置本身）。

    这条边界故意跟"摘要/关键词/Abstract/Keywords/英文题目副标题"这五件事
    自己找没找到、找在哪完全无关——2026-09-19 之前的旧版本让后面几步的
    搜索起点跟着前面几步"已确认找到的位置"往后挪，AAT 实测出"关键词被挪
    到文档末尾，前面一整片全部漏认"的连坐 bug，根源就是这种链式起点。
    换成这条只由更靠后、更确定的锚点（正文/后置区起点）决定的固定终点，
    上面五件事各自在 `[start, 终点)` 里独立扫描，谁先找到、找不找得到都
    不会互相拖累。
    """
    for i in range(start, end):
        text = texts[i][0].strip()
        if not text:
            continue
        if (any(p.match(text) for p in _INTRO_LABELS.values())
                or any(p.match(text) for p in _BACK_LABELS.values())
                or any(p.match(text) for p in _GAP_LABELS.values())
                or grammar.match(text) is not None):
            return i
    return end


def _scan_literal_label(texts, pattern: re.Pattern[str], *,
                        start: int, boundary: int) -> tuple[int | None, list[int]]:
    """在 `[start, boundary)` 里找字面标签：认第一次出现的位置为准，后续
    重复出现的识别到但不采信（原样返回，交给调用方决定怎么处理/告警）。"""
    first: int | None = None
    extra: list[int] = []
    for i in range(start, boundary):
        text = texts[i][0].strip()
        if not text:
            continue
        if pattern.match(text):
            if first is None:
                first = i
            else:
                extra.append(i)
    return first, extra


def _scan_forward(paragraphs, texts, grammar: HeadingGrammar, *, start: int, end: int):
    """摘要/关键词/Abstract/Keywords/英文题目/英文副标题：五件事各自独立地
    在 `[start, 前置区终点)` 里扫，互不连坐（见模块开头的架构说明和
    `_front_matter_boundary` 的说明）。

    返回 `(found, forward_end, duplicates)`：`found` 是每个部件认到的位置；
    `forward_end` 是"已经确认找到的部件里最靠后那个"的下一段（给夹逼区
    扫描当起点用，绝不因为扫描范围延伸到 `boundary` 而往后漂移）；
    `duplicates` 是四个字面标签里"认到但不采信"的重复位置，键跟 `found`
    共用同一套部件名。
    """
    boundary = _front_matter_boundary(paragraphs, texts, grammar, start=start, end=end)

    found: dict[str, int] = {}
    duplicates: dict[str, list[int]] = {}

    for word in _FRONT_LABEL_WORDS:
        first, extra = _scan_literal_label(texts, _FRONT_LABELS[word], start=start, boundary=boundary)
        if first is not None:
            found[word] = first
        if extra:
            duplicates[word] = extra

    for word, part in (("Abstract", "英文摘要"), ("Keywords", "英文关键词")):
        first, extra = _scan_literal_label(texts, _ENGLISH_LABELS[word], start=start, boundary=boundary)
        if first is not None:
            found[part] = first
        if extra:
            duplicates[part] = extra

    # 英文题目/副标题：靠内容形状认，不靠位置——排除掉上面四个标签自己
    # 认到的那几段（不然"Abstract"/"Keywords"这两个词本身就是纯英文、
    # 不带句末标点，会被误认成英文题目），以及目录/清单域、题注样式段落。
    # 先认"工具自己插的占位语"（2026-09-18 实测：成品回炉时占位语本身是
    # 纯中文提示，过不了"像英文"那条判据，认不出来就会在它旁边再插一条，
    # 每回炉一次多两段——跟旧 detect.py 踩过的坑是同一类，用同一套
    # placeholders.py 的字面比对堵住）。
    exclude = set(found.values())
    english_title = None
    for i in range(start, boundary):
        if i in exclude:
            continue
        text = texts[i][0].strip()
        if not text:
            continue
        if has_toc_field(paragraphs[i]) or _is_toc_style(paragraphs[i]) or _is_caption_style(paragraphs[i]):
            continue
        if is_english_title_placeholder(text) or _looks_like_english(text, _EN_TITLE_MIN_LETTERS):
            english_title = i
            break
    if english_title is not None:
        found["英文题目"] = english_title
        j = english_title + 1
        if j < boundary and j not in exclude:
            sub_text = texts[j][0].strip()
            if (sub_text
                    and not has_toc_field(paragraphs[j])
                    and not _is_toc_style(paragraphs[j])
                    and not _is_caption_style(paragraphs[j])
                    and (is_english_subtitle_placeholder(sub_text)
                         or _looks_like_english(sub_text, _EN_SUBTITLE_MIN_LETTERS))):
                found["英文副标题"] = j

    forward_end = max([idx + 1 for idx in found.values()], default=start)
    return found, forward_end, duplicates


# ============================================================ 后向识别器

@dataclass
class BackwardResult:
    intro: int | None
    intro_word: str | None
    anchors: dict[str, list[int]]
    headings: dict[int, HeadingCandidate]
    boundary: int   # 语义见 _scan_backward 顶部说明
    # 带编号前缀的锚点（点分体系的 `1 引言` / `5 结论`）：`{段落下标: (编号,
    # 编号前缀原文)}`。这些段落的文字整段是用户的，套版时原样保留（见
    # `Entry.meta["keep_text"]`）。
    numbered: dict[int, tuple[int, str]]
    # 只因为超过长度上限才没被认成标题的段落（归正文，但要告警"疑似标题"）。
    # 跟 `headings` 一样：被后置标签回收的那一片里的不算（参考文献条目常常又长
    # 又带编号，不能都报）。
    overlong: list[int] = field(default_factory=list)


def _scan_backward(paragraphs, texts, grammar: HeadingGrammar, *,
                   claimed: set[int], end: int) -> BackwardResult:
    """从 `end - 1` 起向前：认后置标签与正文标题，标签回收它管辖内暂存的
    正文标题；四选一停止条件命中即收工。

    `boundary` 的含义统一成：下标 `>= boundary` 的段落已经被本扫描器"处理
    过"（认成了什么，或者认过了但判定不算），下标 `< boundary` 的一概没碰。
    三种停止方式因此各自的 `boundary` 取值不同，见对应分支里的注释。
    """
    intro: int | None = None
    intro_word: str | None = None
    anchors: dict[str, list[int]] = {}
    numbered: dict[int, tuple[int, str]] = {}
    pending: dict[int, HeadingCandidate] = {}
    overlong: list[int] = []
    cjk_seen: set[str] = set()
    streak_index: int | None = None
    streak_len = 0
    boundary = 0

    i = end - 1
    while i >= 0:
        text = texts[i][0].strip()
        if not text:
            i -= 1
            continue
        # 锚点（引言家族/结论）前允许有编号前缀——只有点分体系会给出，中文体系
        # 这里恒为 `(None, text)`。
        anchor_number, anchor_text = grammar.split_anchor_number(text)
        anchor_prefix = text[:len(text) - len(anchor_text)].strip()

        # 条件②：目录域/目录清单类字面标签/前向识别器已认过的部件——
        # 本段交给夹逼区扫描，后向识别器不处理它，到此为止。
        if (i in claimed or has_toc_field(paragraphs[i]) or _is_toc_style(paragraphs[i])
                or any(p.match(text) for p in _GAP_LABELS.values())):
            boundary = i + 1
            break

        # 条件①：引言家族（含绪论/序言/前言/导论/导言/引论/绪言）——本段被认成
        # "引言"，停止。
        hit_intro = next((w for w, p in _INTRO_LABELS.items() if p.match(anchor_text)), None)
        if hit_intro is not None:
            intro, intro_word, boundary = i, hit_intro, i
            if anchor_number is not None:
                numbered[i] = (anchor_number, anchor_prefix)
            break

        # 后置标签：回收暂存的正文标题（它们大概率是这个部件内容里的条目，
        # 不是真章节标题），本身继续往前走，不停。
        hit_back = next(
            (w for w, p in _BACK_LABELS.items()
             if p.match(anchor_text if w in _NUMBERABLE_BACK_WORDS else text)), None)
        if hit_back is not None:
            if hit_back not in anchors:
                anchors[hit_back] = [i]
                if anchor_number is not None and hit_back in _NUMBERABLE_BACK_WORDS:
                    numbered[i] = (anchor_number, anchor_prefix)
            pending.clear()
            overlong.clear()
            cjk_seen.clear()
            streak_index, streak_len = None, 0
            i -= 1
            continue

        candidate = grammar.match(text)
        if candidate is not None:
            cjk = _cjk_only(text)
            is_duplicate = bool(cjk) and cjk in cjk_seen
            if is_duplicate and streak_index == i + 1:
                streak_len += 1
            elif is_duplicate:
                streak_len = 1
            else:
                streak_len = 0
            streak_index = i
            if is_duplicate and streak_len >= 2:
                # 条件④：连续两段命中已认出的正文标题列表——多半是手打目录
                # 条目撞上了真标题的文字。撤销这两段，扫描到此为止（`i`
                # 本身已经被检查过、判定不算，`i` 之前一律不再看）。
                pending.pop(i, None)
                pending.pop(i + 1, None)
                boundary = i
                break
            pending[i] = candidate
            if cjk:
                cjk_seen.add(cjk)
            i -= 1
            continue

        # 普通正文：不是标签也不是候选标题，打断"连续命中"的连号，继续往前。
        # 只是因为太长才没认成标题的，记下来（归正文，同时告警"疑似标题"）。
        if grammar.overlong(text):
            overlong.append(i)
        streak_index, streak_len = None, 0
        i -= 1
    else:
        boundary = 0   # 条件③：扫到了文档头，一路没有其他条件命中

    return BackwardResult(intro=intro, intro_word=intro_word, anchors=anchors,
                          headings=pending, boundary=boundary, numbered=numbered,
                          overlong=sorted(overlong))


# ============================================================ 夹逼区：目录 + 三张清单

def _scan_gap(paragraphs, texts, *, start: int, end: int) -> dict[str, int]:
    found: dict[str, int] = {}
    for i in range(start, end):
        text = texts[i][0].strip()
        if not text:
            continue
        for word, pattern in _GAP_LABELS.items():
            if word not in found and pattern.match(text):
                found[word] = i
    return found


# ============================================================ 题目 / 副标题（游离正文）

def _find_title_subtitle(paragraphs, texts, *, start: int, boundary: int):
    """`boundary` 之前（从 `start` 起）的非空段落＝游离正文：第一行题目，
    紧接着题目下一段才算副标题。目录/清单域生成的条目、题注类样式的段落
    不算候选——它们单看文字可能完全不像标签/标题，但显然不是"游离正文"
    （`TOC.docx` 实测：清单条目段与它旁边的 Caption 样式段都会撞进来）。

    ⚠️ 副标题必须**紧邻**题目（2026-09-19 用户纠正）：题目和第二行候选之间
    隔着空段（或目录/题注类段落）时，那不是副标题——空段之后的文字是
    "游离正文"，不该被认成副标题。`lines` 已经把空段过滤掉了，所以这里
    改成直接比较**原始段号**是不是紧挨着（`lines[1] == lines[0] + 1`），
    不能只看"是不是候选列表里的第二项"。
    """
    lines = [i for i in range(start, boundary)
            if texts[i][0].strip()
            and not has_toc_field(paragraphs[i])
            and not _is_toc_style(paragraphs[i])
            and not _is_caption_style(paragraphs[i])]
    # 用户 2026-09-20 纠正：正标题/副标题**仍然要求不以句末标点结尾**（此前误
    # 记成"不看标点"）。正标题＝范围内第一个不以句末标点结尾、且不是全英文的
    # 段落；副标题＝紧邻正标题下一段（原始段号相邻）、同样不以句末标点结尾。
    def ends_with_punct(i: int) -> bool:
        return texts[i][0].strip().endswith(tuple(_TRAILING_PUNCT))

    def all_english(i: int) -> bool:
        text = texts[i][0]
        return (not any("一" <= ch <= "鿿" for ch in text)
                and any(ch.isascii() and ch.isalpha() for ch in text))

    title = next((i for i in lines if not ends_with_punct(i) and not all_english(i)), None)
    subtitle = None
    if title is not None:
        following = next((i for i in lines if i > title), None)
        if following is not None and not ends_with_punct(following):
            # 紧邻才算副标题（见上）。**唯一的例外**：工具自己插的"副标题占位语"——它是
            # 排在"题目之后、下一个部件之前"的，而题目后面的空行会先被搬过去，所以占位语
            # 常常落在空行**之后**。认不出它，回炉导出时会当成"没有副标题"再插一条
            # （2026-09-20 复现：题目后有空行 → 每导一遍多一条副标题占位语）。
            if following == title + 1 or is_subtitle_placeholder(texts[following][0]):
                subtitle = following
        if subtitle is None:
            # 占位语被插在题目后面的别的自由文字（作者、日期之类）**之后**也一样：只要有
            # 一段跟工具占位语一字不差，就是工具自己占的副标题位，不再补第二条。
            subtitle = next((i for i in lines
                             if i > title and is_subtitle_placeholder(texts[i][0])), None)
    return title, subtitle


# ============================================================ 顶层入口

# 有"标签+内容"两段式结构的部件——标签本身、以及"紧贴标签之后、到下一个
# 已认出部件之前"的非空段落都要标出来。附录的"从第二段起当正文"这条细分
# （附录A 序号题名）本轮不做，整段先当"附录"内容处理，留给后续再细分。
_CONTENT_OWNERS = {"摘要", "关键词", "英文摘要", "英文关键词",
                  "引言", "结论", "参考文献", "附录", "致谢"}
# 只有标签、没有"内容"这回事的部件（2026-09-19 补）：目录/三张清单的正文
# 是 Word 域自己生成的，不能套用上面那套"标签后面一路扫到下一个已认部件"
# 的内容归属——`民总.docx` 实测过：已有的目录域往往只有第一段带真正的域
# 代码，其余条目段落是纯 Normal 样式，逐段查 `has_toc_field`/`_is_toc_style`
# 根本认不出来，一扫就把它们当"目录的内容"，后面重建时被整段改写成占位语
# （详见 `apply_structure._LABEL_CONTENT_STYLE_IDS["目录"]` 的第二个元素是
# `None`——这个部件设计上就没有"内容样式"这一档）。但标签本身仍然要标
# `slot="label"`，不然 `apply_from_index` 的 `slot == "label"` 判据永远
# 拿不到 True，标签套不上 S14（三号黑体居中、段前分页），现象就是"目录
# 标签没居中、没分页"——这条名单只管"标一下这是标签"，不触发内容扫描。
_LABEL_ONLY_PARTS = {"目录", "插图清单", "附表清单", "插图和附表清单"}


def _mark_numbered_anchor(entry: Entry, numbered: tuple[int, str] | None,
                          auto_numbered: bool) -> None:
    """带编号前缀的锚点（`1 引言`/`5 结论`）：编号是用户的文字，套版时整段
    原样保留（`keep_text`）；同时留下层级/编号，让序号检查把它当一级标题
    那一串里的一员（不然 `1 引言` 之后的 `2 …` 会被报"未归零"）。

    编号如果是 Word 自动编号画的（不在段落文字里），跟点分标题一样要固化成
    文字（`canonical` ＝ 渲染出来的原样编号 + 一个空格），否则清直接格式
    时编号引用被删、成品里编号消失。"""
    if numbered is None:
        return
    number, prefix = numbered
    entry.meta.update(keep_text=True, level=1, number=(number,))
    if auto_numbered:
        entry.meta.update(auto_numbered=True, canonical=prefix + " ")


def recognize(document, *, grammar: HeadingGrammar | None = None,
             cover_end: int | None = None) -> list[Entry]:
    """识别一份文档，把结果写进 `doc_index.build_index()` 产出的 `Entry`
    列表（`.part`/`.meta`），直接返回这份索引——调整识别结果窗口和未来
    的加工层都从这里取数据，不需要另一套格式。

    调用方需自己先 `unwrap_section_break_sdt(document, keep_content=True)`
    ——跟 `doc_index.build_index` 的约定一致（见 `open_and_recognize`）。

    `grammar`：标题编号方案（`grammar_for()` 按界面上的选择给出）；不传
    就是默认的中文方案。
    `cover_end`：用户手敲的封面范围（第 0 到第 N 段，含），不参与识别。
    """
    grammar = grammar or grammar_for(None)
    index = build_index(document)
    ordinals = paragraph_ordinals(index)
    paragraphs = document.paragraphs
    numbering = load_numbering_scheme(document)
    # 目录/清单域区间里、域指令那一段**之后**的缓存条目段（民总的目录就是这样：
    # 只有第一行带域代码，其余是纯 Normal 文字）：它们是 Word 渲染出来的条目，
    # 不是用户写的内容——识别时当空段（不可能成为标题/题注/题目/内容），见
    # 操作逻辑 3.6"域区间内的段落视为条目"。域指令那一段本身不动（后向扫描靠它停）。
    toc_rest = {i for i in toc_field_span_indices(paragraphs)
                if not has_toc_field(paragraphs[i])}
    texts = [("", False) if i in toc_rest else _effective_text(p, numbering)
             for i, p in enumerate(paragraphs)]
    n = len(paragraphs)
    start = 0 if cover_end is None else min(cover_end + 1, n)

    forward_found, forward_end, forward_duplicates = _scan_forward(
        paragraphs, texts, grammar, start=start, end=n)
    forward_claimed = set(forward_found.values())

    backward = _scan_backward(paragraphs, texts, grammar, claimed=forward_claimed, end=n)

    gap_found: dict[str, int] = {}
    if backward.boundary > forward_end:
        gap_found = _scan_gap(paragraphs, texts, start=forward_end, end=backward.boundary)

    # "所有标签之前"——这里的"标签"也把已经认出的正文标题算进去：正文标题
    # 显然不是"游离正文"，题目/副标题不能越过第一个正文标题去抢它的位置
    # （2026-09-18 用 TitleIdentify.docx 实测发现：这份文件没有任何字面标签，
    # 不把标题位置算进边界的话，第一段真标题会被误判成论文题目）。
    label_positions = (list(forward_found.values()) + list(gap_found.values())
                       + list(backward.headings.keys()))
    if backward.intro is not None:
        label_positions.append(backward.intro)
    for positions in backward.anchors.values():
        label_positions.extend(positions)
    title_boundary = max(min(label_positions, default=n), start)
    title_idx, subtitle_idx = _find_title_subtitle(paragraphs, texts, start=start,
                                                   boundary=title_boundary)

    # ---- 写回 Entry ----
    if title_idx is not None:
        ordinals[title_idx].part = "题目"
    if subtitle_idx is not None:
        ordinals[subtitle_idx].part = "副标题"

    for part, idx in forward_found.items():
        ordinals[idx].part = part
        extra = forward_duplicates.get(part)
        if extra:
            ordinals[idx].meta["duplicate_label_positions"] = tuple(extra)

    for part, idx in gap_found.items():
        ordinals[idx].part = part
        ordinals[idx].meta["field"] = {
            "目录": "toc", "插图清单": "图", "附表清单": "表",
            "插图和附表清单": "图表",
        }[part]

    if backward.intro is not None:
        ordinals[backward.intro].part = "引言"
        ordinals[backward.intro].meta["matched_word"] = backward.intro_word
        _mark_numbered_anchor(ordinals[backward.intro], backward.numbered.get(backward.intro),
                              texts[backward.intro][1])

    for name, positions in backward.anchors.items():
        for idx in positions:
            ordinals[idx].part = name
            _mark_numbered_anchor(ordinals[idx], backward.numbered.get(idx), texts[idx][1])

    for idx in backward.overlong:
        ordinals[idx].meta["overlong_title"] = True     # 归正文；告警见 alerts.check_demoted_titles

    for idx, candidate in backward.headings.items():
        entry = ordinals[idx]
        auto_numbered = texts[idx][1]
        canonical = candidate.canonical
        if canonical is None and auto_numbered:
            # 不规范化的编号（点分体系）：Word 自动编号仍然必须固化成文字，不然
            # 清直接格式那一步会把编号引用删掉、成品里编号消失——固化的就是
            # Word 渲染出来的那串原样文字（回算成文字格式，跟手打的一视同仁）。
            canonical = candidate.marker + " "
        entry.part = LEVEL_PART_NAMES[candidate.level]
        entry.meta.update(
            level=candidate.level, number=candidate.number, marker=candidate.marker,
            title=candidate.title, canonical=canonical,
            auto_numbered=auto_numbered,
        )

    # ---- 标签+内容两段式：给标签标 slot="label"，给内容段标 slot="content" ----
    # 2026-09-19 改：不再局限于段落——表格、图片独占的段落、空段都可能落在
    # 标签管辖的区间里；以前只在"段落序号"这条窄编号里扫，表格根本看不见
    # （表格不在 `document.paragraphs` 里），没有文字的段落又被故意跳过，
    # 这些东西的归属只能靠导出时"还没处理、按原文位置兜底"（`doc_index.
    # build_ops()` 的 `flush_before`）撑着——这正是"输出层不许看原文档顺序，
    # 完全按索引来"这条新要求要拔掉的口子（2026-09-19 用户原话）。改成在
    # body 全部顶层元素（`index` 本身就是按 `top_index` 排好的顺序，
    # `index[k].top_index == k`）里扫，用 `top_index` 定区间，不再用段落
    # 序号；`has_toc_field`/`_is_toc_style` 只对 `kind=="paragraph"` 的条目
    # 有意义，查的时候用 `top_to_ordinal` 换回段落对象。
    # ⚠️ `top_index` **不保证连续**：`doc_index.build_index()` 只认 w:p/w:tbl/
    # w:sdt/w:sectPr 这四种顶层子元素，书签/批注范围这类其它字符串标签的
    # 顶层子元素既不追加 Entry、也不占 `index` 的列表位置，会在 `top_index`
    # 序列里留下空洞——不能拿 `index[top]` 直接当"top_index 为 top 的那条"
    # 用（2026-09-19 实测：中国政法大学硕士、博士学位论文写作规范（更新）.docx
    # 里就有这种空洞，直接下标访问会越界）。一律用 `top_index → Entry` 的
    # 字典查、按 `top_index` 属性值（不是列表位置）过滤区间。
    by_top = {e.top_index: e for e in index}
    top_to_ordinal = {e.top_index: ordinal for ordinal, e in ordinals.items()}
    doc_end_top = index[-1].top_index + 1 if index else 0
    start_top = ordinals[start].top_index if start < n else doc_end_top
    claimed_top_sorted = sorted(
        e.top_index for e in index if e.part and e.top_index >= start_top)
    for pos, top in enumerate(claimed_top_sorted):
        entry = by_top[top]
        if entry.part in _LABEL_ONLY_PARTS:
            entry.meta.setdefault("slot", "label")
            continue
        if entry.part not in _CONTENT_OWNERS:
            continue
        entry.meta.setdefault("slot", "label")
        nxt_top = claimed_top_sorted[pos + 1] if pos + 1 < len(claimed_top_sorted) else doc_end_top
        for content_entry in index:
            if content_entry.top_index <= top:
                continue
            if content_entry.top_index >= nxt_top:
                break           # `index` 按 top_index 递增排列，可以提前收工
            if content_entry.part:
                continue
            if content_entry.kind == "paragraph":
                ordinal = top_to_ordinal.get(content_entry.top_index)
                paragraph = paragraphs[ordinal] if ordinal is not None else None
                if paragraph is not None and (has_toc_field(paragraph) or _is_toc_style(paragraph)
                                              or ordinal in toc_rest):
                    continue    # 目录/清单域本体或它生成的条目不算"内容"
            content_entry.part = entry.part
            content_entry.meta["slot"] = "content"

    # ---- 图题/表题：识别一次，写进 `meta["caption"]`（不改 part）----
    # 告警、导出、结构页都从这里读，不再各自重新识别（2026-09-19）。已判成
    # 数字编号标题的段落不可能同时是题注。
    # ⚠️ 封面范围（`start` 之前那一段）必须一并排除——`detect_captions_in`
    # 以前收到的是**全篇未裁剪**的 `paragraphs`，封面范围内凑巧像"图 1"/
    # "表 1"这种写法的文字会被当真题注、套上题注样式（2026-09-22 审计揪出的
    # 真 bug：封面范围本该"工具不识别也不改"，见 `apply_from_index` 对
    # `skip_indices` 的处理）。跟标题/目录条目一样，直接并进"跳过"集合。
    heading_ordinals = {i for i, e in ordinals.items() if e.part in LEVEL_PART_NAMES.values()}
    cover_ordinals = set(range(start)) if cover_end is not None else set()
    for caption in detect_captions_in(paragraphs, heading_ordinals | toc_rest | cover_ordinals):
        ordinals[caption.paragraph_index].meta["caption"] = caption_meta(caption)

    return index


def open_and_recognize(path, *, grammar: HeadingGrammar | None = None,
                       cover_end: int | None = None):
    """读一份 docx 并识别（先拆内容控件，跟 `doc_index.open_and_index` 同一口径）。"""
    from docx import Document

    from essay_tool.engine.cleanup import unwrap_section_break_sdt

    document = Document(str(path))
    unwrap_section_break_sdt(document, keep_content=True)
    index = recognize(document, grammar=grammar, cover_end=cover_end)
    return document, index
