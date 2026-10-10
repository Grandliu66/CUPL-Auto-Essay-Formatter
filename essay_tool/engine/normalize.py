"""把识别结果里"写法不符合规范"的地方改写成规范写法。

识别是**宽进**的：`1、`、半角括号、`(一）` 这种混搭、缺【】的标签，都能认
出来（见 `recognize.py` 的模块 docstring）。但"认出来"不等于"能原样套版"——
规范对标题怎么编号、标签该怎么写有明文规定，套版时必须把文字改成规范的
写法，否则学生会拿到一份"识别对了、但正文里还是错写法"的成品。

在管线里的位置：**识别之后**（识别要靠原文找线索），清直接格式/套样式之前。
识别结果只存段落下标、不引用 XML 元素，所以改文字不会让下标失效。

改写的三条原则：
  * 只动段落开头那一段"编号/标签"文字，**标题内容本身一个字不碰**；
  * 写法本来就规范的段落**完全不碰**（连 run 都不重排），避免无谓改动；
  * 改写靠拼装 run 的文字完成，段落里其余内容（脚注引用、域代码这类不属于
    文字的 XML 元素）原样保留。

"规范的写法"从哪来：
  * **标题编号**：识别层（`recognize` 里当前选用的文法）已经把规范写法
    算好放在 `Entry.meta["canonical"]` 里，这里直接用，不重算一遍：
    中文体系是 `一、`、`（一） `、`1. `、`（1） `（全角圆括号、半角句点、
    编号后一个半角空格、一级标题编号后不空格）；点分体系只有 `（1） `
    有规范写法，`1.1` 这类点分编号的 `canonical` 是 None——**用户的文字
    原样保留，不规范化**。
    规范写法在识别那一刻按当时的层级定死，用户之后拖动层级不重编号。
  * **标签文字**（摘要/关键词/目录/英文摘要/引言/结论/参考文献/附录/致谢）：
    从预设"结构部件"的"插入文字"那一格来——用户改预设就能改
    写法，不用动代码。格子是空的（题目/副标题本来就没标准答案）或者读不到
    的，就保持原样不动。**"引言"家族的同义词（绪论/序言/前言/导论……）
    保留用户写的词，只借"引言"那一格的空格和括号**（`_intro_label_text`）——
    `recognize.py` 已经把它们的 `Entry.part` 归一成"引言"，但文字不归一。
"""

from __future__ import annotations

import re

from docx.document import Document as DocumentObject
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from essay_tool.engine.cleanup import set_run_text
from essay_tool.engine.doc_index import Entry, paragraph_ordinals
from essay_tool.engine.recognize import INTRO_WORDS, LEVEL_PART_NAMES, split_inline_keyword

# 识别出来的"英文摘要"/"英文关键词"标签该写成什么样（规范写法）。旧预设里
# 只有"英文摘要"一格、装两行（上行摘要标签、下行关键词标签），读不到
# "英文关键词"那一格时回落到那个写法的第二行。
_ENGLISH_LEGACY_FALLBACK = {"英文关键词": ("英文摘要", 1)}

# 有"标签/标题"身份、需要按 insert_text 整段改写文字的部件（前置标签 +
# 锚点标签，跟内容段是两回事——内容段一个字不碰）。
_BARE_LABEL_PARTS = frozenset({
    "摘要", "关键词", "英文摘要", "英文关键词",
    "目录", "插图清单", "附表清单", "插图和附表清单",
    "引言", "结论", "参考文献", "附录", "致谢",
})
_ENGLISH_LABEL_PARTS = frozenset({"英文摘要", "英文关键词"})


def _text_runs(paragraph: Paragraph) -> list[Run]:
    """段落里**参与 `paragraph.text` 的所有 run**，按文档顺序。

    ⚠️ 不能直接用 `paragraph.runs`：它只给段落的直接子 run，而 `paragraph.text`
    把 `w:hyperlink` 里的 run 也算上。本模块按 `paragraph.text` 的下标算前缀长度、
    再靠 run 写回——整段被超链接包住时 `paragraph.runs` 是空的，写回落空，兜底的
    `add_run` 又把规范编号**追加到段尾**，每回炉一遍多补一个（AAT 2026-09-19
    「全部段落包进超链接不收敛」，尾部空格逐遍递增）。读和写必须用同一批 run。
    """
    out: list[Run] = []
    for child in paragraph._p:
        if child.tag == qn("w:r"):
            out.append(Run(child, paragraph))
        elif child.tag == qn("w:hyperlink"):
            out.extend(Run(r, paragraph) for r in child.findall(qn("w:r")))
    return out


def _splice_prefix(paragraph: Paragraph, prefix_len: int, replacement: str) -> bool:
    """把段落开头的 prefix_len 个字符替换成 replacement。

    替换文字写在第一个受影响的 run 里——那个 run 本来就装着编号，编号的字体
    格式跟着它走是合理的；后面被吃掉的 run 清空文字但**不删除 run 本身**，
    免得误伤挂在 run 上的其他元素。返回是否真的改动过。
    """
    if prefix_len <= 0:
        return False

    text = paragraph.text
    if text[:prefix_len] == replacement:
        return False  # 本来就是规范写法，一个 run 都不动

    remaining = prefix_len
    placed = False
    for run in _text_runs(paragraph):
        run_text = run.text
        if remaining <= 0:
            break
        if len(run_text) <= remaining:
            remaining -= len(run_text)
            if not placed:
                set_run_text(run, replacement)
                placed = True
            else:
                set_run_text(run, "")
        else:
            tail = run_text[remaining:]
            set_run_text(run, (replacement + tail) if not placed else tail)
            placed = True
            remaining = 0

    if not placed and replacement:
        # 段落里一个 run 都没有（理论上不该出现），兜底补一个
        paragraph.add_run(replacement)
    return True


def _prefix_len(paragraph_text: str, title: str) -> int:
    """段落开头"编号/标签"部分占几个字符——也就是标题内容之前的那一段。

    段首的空白也算进去（规范写法里不该有前导空白，一并吃掉）；标题内容本身
    从 `title` 反推：识别时 title 是取编号后那截再 strip 的，所以规范文本里
    标题一定落在最后。

    ⚠️ **编号后面一个字都没写时（`1.` / `1.  `），整段都是前缀**（2026-09-17
    修的 bug，AAT 报上来的）：那时 `title` 是空串，`len(stripped) - 0` 只数到
    "1."两个字符，**尾部原先的空白留在原地**，而规范写法 `1. ` 又自带一个
    空格 ⇒ 每回炉一次多一个空格：`1.` → `1. ` → `1.  ` → `1.   `…**永远不
    收敛**。有标题文字的行不受影响（"1. 标题" 里空格正好落在替换范围内）。
    """
    if not title:
        return len(paragraph_text)
    stripped = paragraph_text.strip()
    lead = len(paragraph_text) - len(paragraph_text.lstrip())
    return lead + len(stripped) - len(title)


def _drop_num_pr(paragraph: Paragraph) -> None:
    """摘掉"这一段属于某个自动编号列表"的引用（`w:numPr`）。

    不摘的话 Word 会照着自己的列表再画一道号，成品上就变成两个编号（我们刚
    写进去的手打文字 + Word 现画的）。
    """
    pPr = paragraph._p.find(qn("w:pPr"))
    if pPr is None:
        return
    numPr = pPr.find(qn("w:numPr"))
    if numPr is not None:
        pPr.remove(numPr)


def _materialize_number(paragraph: Paragraph, prefix: str) -> None:
    """把 Word 自动编号取消掉、写成真正的文字（＝手打编号）。

    自动编号不在段落文字里，是 Word 按 numbering.xml 现算现画的；成品里要
    取消它，于是把编号写成真文字、再把 `w:numPr` 摘掉。`prefix` 是**规范
    写法**的编号文字（`（1） `、`一、`，来自 `Entry.meta["canonical"]`，
    识别时按当时的层级算一次就定死，不会随用户后续调整层级重算）。

    编号文字塞进**第一个 run**的文字里、而不是新建一个 run——这样编号能跟着
    标题原本的字体走。
    """
    if not prefix:
        return
    # 幂等：成品/回炉稿里这个编号已经是文字了（`（1） 标题`），再写一遍就
    # 成了两个号。`w:numPr` 那时也已经不在了，两重保险。
    if paragraph.text.lstrip().startswith(prefix.strip()):
        return
    runs = _text_runs(paragraph)
    if runs:
        first = runs[0]
        set_run_text(first, prefix + first.text)
    else:
        paragraph.add_run(prefix)
    _drop_num_pr(paragraph)


def _label_text(insert_text: dict[str, str], name: str) -> str:
    return (insert_text.get(name) or "").strip()


def _english_label_text(insert_text: dict[str, str], part: str) -> str:
    """英文摘要/关键词标签的规范写法：各自部件的"插入文字"格；读不到时
    回落到旧写法（"英文摘要"那一格的第二行是关键词标签）。"""

    def lines(name: str) -> list[str]:
        raw = insert_text.get(name) or ""
        return [line.strip() for line in raw.splitlines() if line.strip()]

    got = lines(part)
    if got:
        return got[0]
    legacy = _ENGLISH_LABEL_PARTS and _ENGLISH_LEGACY_FALLBACK.get(part)
    if legacy is not None:
        old = lines(legacy[0])
        if len(old) > legacy[1]:
            return old[legacy[1]]
    return ""


_CJK_CHAR = re.compile(r"[一-鿿]")
_LABEL_DECORATION = re.compile(r"[\s【】\[\]]")


def _intro_label_text(template: str, current: str) -> str:
    """引言家族标签的写法：**保留用户写的那个词，只借预设那一格的空格和括号**。

    预设"引言"格写的是"引　言"，用户写的是"绪论"就补成"绪　论"（2026-09-19
    用户裁决："认出来了就不强制规范化文字内容为'引言'"，铁律"任何情况下
    不修改用户的文字内容，规范化除外"）。做法：把预设格里的汉字**逐个换成**
    用户那个词的汉字，空格、括号原样留着。

    两种不适用的情形都退回"不改"：`current` 去掉空格括号后不是家族词（用户
    在结构页把别的行拖成了引言锚点，此时没有"用户的词"可保留，走预设写法）；
    预设格的汉字个数跟用户的词对不上（预设格写得不是一个词，没法逐字换）。
    """
    core = _LABEL_DECORATION.sub("", current)
    if core not in INTRO_WORDS:
        return template
    if len(_CJK_CHAR.findall(template)) != len(core):
        return current
    chars = iter(core)
    return _CJK_CHAR.sub(lambda _m: next(chars), template)


def _label_replacement(part: str, insert_text: dict[str, str], current: str = "") -> str:
    """`current`：段落当前的文字（引言家族要保留用户的词，其余部件不看它）。"""
    if part in _ENGLISH_LABEL_PARTS:
        return _english_label_text(insert_text, part)
    template = _label_text(insert_text, part)
    if part == "引言" and template:
        return _intro_label_text(template, current)
    return template


def inline_keyword_text(part: str, insert_text: dict[str, str], raw: str) -> str | None:
    """关键词同段写法（`关键词：民法典；合同`）在成品里的文字：**只把开头的标签**
    换成预设那一格的写法，关键词一个字不碰。`raw` 不是同段写法返回 None。

    标签后面的分隔：新标签以括号收尾（`【关键词】`）→ 去掉用户的冒号，空白原样保留；
    新标签不带括号（学术十星 `关键词`）→ 冒号、空白原样保留，什么分隔都没有就补
    一个空格（免得粘成"关键词民法典"）。预设那一格被清空＝保持原样。
    """
    parts = split_inline_keyword(raw, part)
    if parts is None:
        return None
    label = _label_replacement(part, insert_text, raw[parts.label_start:parts.label_end])
    if not label:
        return raw
    colon = raw[parts.label_end:parts.colon_end]
    space = raw[parts.colon_end:parts.rest_start]
    if label.rstrip().endswith(("】", "]")):
        sep = space
    else:
        sep = colon + space or " "
    return label + sep + raw[parts.rest_start:]


def normalize_document(
    document: DocumentObject,
    index: list[Entry],
    insert_text: dict[str, str],
    auto_bodies: dict[int, str] | None = None,
) -> None:
    """按识别索引把标题编号、标签文字改写成规范写法。取代旧的
    `normalize_headings`/`normalize_front_matter_labels`——两者的判据都已经
    在识别阶段算完、写进了 `Entry.part`/`Entry.meta`，这里只是照着写。

    `auto_bodies`：被用户从自动编号标题拖回"正文"的段落（`{段号: 编号文字}`）
    ——这些段落已经不是标题了（不套样式、不进目录），但 Word 的自动编号
    仍要取消、写成手打文字，否则那段编号会凭空消失。本轮识别层还没接
    "调整识别结果"窗口，自动路径下这个参数恒为空，先留着接口。
    """
    paragraphs = document.paragraphs
    ordinals = paragraph_ordinals(index)
    heading_parts = frozenset(LEVEL_PART_NAMES.values())

    for i, entry in ordinals.items():
        part = entry.part
        if not part:
            continue
        paragraph = paragraphs[i]

        if part in heading_parts:
            if entry.meta.get("auto_numbered"):
                _materialize_number(paragraph, entry.meta.get("canonical") or "")
                continue
            prefix = entry.meta.get("canonical")
            if prefix is None:
                continue
            _splice_prefix(paragraph, _prefix_len(paragraph.text, entry.meta.get("title", "")),
                           prefix)
            continue

        if entry.meta.get("inline") and part in ("关键词", "英文关键词"):
            # 同段写法：只改开头的标签，关键词原样（见 `inline_keyword_text`）。
            text = paragraph.text
            new = inline_keyword_text(part, insert_text, text)
            if new is not None and new != text:
                parts = split_inline_keyword(text, part)
                tail = len(text) - parts.rest_start
                _splice_prefix(paragraph, len(text) - tail, new[:len(new) - tail])
            continue

        if part in _BARE_LABEL_PARTS and entry.meta.get("slot", "label") == "label":
            if entry.meta.get("keep_text"):
                # 带编号的引言/结论（`1 引言`）：整段是用户的文字，原样保留；
                # 只有编号是 Word 自动编号画的，才把它固化成文字。
                if entry.meta.get("auto_numbered"):
                    _materialize_number(paragraph, entry.meta.get("canonical") or "")
                continue
            replacement = _label_replacement(part, insert_text, paragraph.text)
            if not replacement:
                continue  # 用户在 Excel 里把这格清空了，保持原样
            _splice_prefix(paragraph, len(paragraph.text), replacement)

    for idx, prefix in (auto_bodies or {}).items():
        if 0 <= idx < len(paragraphs):
            _materialize_number(paragraphs[idx], prefix)


# ------------------------------------------------------------------ 只算不写
# 下面两个函数是给**结构页显示**用的：一个字符都不改文档，只是把"导出时会
# 写成什么样"算出来——素材和上面的改写函数完全共用，结构页显示的写法与
# 成品里的写法不会各说各话。

def heading_style_text(match, insert_text: dict[str, str], raw: str = "") -> str:
    """这一段标题在成品里会写成什么样（结构页显示用）。

    `match` 是结构页那一行上的 `detect.HeadingMatch`（鸭子类型：`.role`/
    `.canonical`/`.keep_text`/`.title`/`.auto_numbered`）。

    ⚠️ **编号写法在识别那一刻按当时的层级定死**（`canonical`），之后用户把
    这一行拖到别的级别只改样式，编号文字一个字不动（用户 2026-09-19："拖动
    绝对不能编号""任何情况下不修改用户的文字内容，规范化除外"——翻转了
    2026-09-15 那条"拖到别的级别就按新级别重编"的老口径）。所以这里没有
    "按当前级别重算"这一步，显示和导出（`normalize_document` 读的同一个
    `canonical`）不会各说各话。

    三种结果：
      * 固定名称锚点（引言/结论）：带编号的（`keep_text`）原样显示；不带的
        显示成"插入文字"那一格的规范写法；
      * 有规范写法（`canonical`，中文体系的四级、点分体系的 `（1）`）：
        规范编号 + 标题内容；
      * 没有规范写法（点分体系的 `1.1`）：原文。
    """
    if match.role:
        if match.keep_text:
            return raw or match.title
        return _label_replacement(match.role, insert_text, raw or match.title) or raw or match.title
    if match.auto_numbered:
        return match.canonical or raw or match.title
    if match.canonical:
        return match.canonical + match.title
    return raw or match.title


def label_style_text(part: str, insert_text: dict[str, str], raw: str = "") -> str:
    """前置/后置标签在成品里写成什么样（摘要→【摘要】、目录→【目录】……）。

    写法只认预设「插入文字」那一格，跟 `normalize_document()` 同源；空格是
    空的（用户自己清掉了）就返回 `raw`。
    """
    return _label_replacement(part, insert_text, raw) or raw
