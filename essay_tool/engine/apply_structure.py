"""把已经确定的标签方案（识别结果，以后可能经过 GUI 调整）套到文档上。

这一层不判断标签对不对——那是识别（detect.py/structure.py）和以后 GUI
调整阶段的事。这里只接收一份已确定的方案，管"给定一个标签，应该长
什么样"。所以函数都不自己调 detect_headings() 这些识别函数，而是
接收调用方（现在是 pipeline.py，直接传识别结果；以后是 GUI 调整完的
结果）传进来的方案，方便以后在识别和应用之间插入调整环节而不用动这里。

摘要/关键词（含英文 Abstract/Keywords）是"标签独占一段，内容从下一段
开始写，直到下一个已识别的标题/标签为止"——用测试文件验证过，标签和
内容不会挤在同一段里，摘要和关键词是同一种模式，不用分两套逻辑。目录
的"内容"是以后要生成的 TOC 域，不是现在要套格式的段落，这里不处理。
"""

from __future__ import annotations

import copy
import re

from docx.document import Document as DocumentObject
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.hyperlink import Hyperlink
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from essay_tool.engine.cleanup import iter_tables, set_run_text
from essay_tool.engine.detect import CaptionMatch
from essay_tool.engine.doc_index import (
    Entry,
    LIST_PART_KINDS,
    has_toc_field,
    is_toc_paragraph,
    paragraph_ordinals,
    toc_field_span_indices,
)
from essay_tool.engine.numbering import (ensure_reference_numbering,
                                        paragraph_num_id, reference_level_text,
                                        copy_level_indent)
from essay_tool.engine.placeholders import (
    ANCHOR_CONTENT_PROMPTS,
    APPENDIX_BODY_PROMPT,
    CONTENT_PLACEHOLDER_PROMPTS,
    ENGLISH_PLACEHOLDER_PROMPTS,
    PLACEHOLDER_PROMPTS,
    REFERENCE_CONTENT_PROMPT_PLAIN,
)
from essay_tool.engine.sections import (insert_toc_field_after, make_list_field,
                                        make_toc_field)
from essay_tool.engine.styles import StyleRegistry, apply_spec, write_spec_onto_paragraph

_LEVEL_STYLE_IDS = {1: "S02", 2: "S03", 3: "S04", 4: "S05"}

_ANCHOR_STYLE_IDS = {
    "引言": "S02",
    "结论": "S02",
    "参考文献": "S02",
    "致谢": "S02",
    "附录": "S19",
}

# 附录套的是 S19（"附录标题"）：它由 styles 建出一个**独立样式**「附录标题」，
# 小四黑体加粗那套字体落在**样式层**而不是段落上——这一点很要紧：写在段落上的
# 话，Word 生成目录条目时会把标题的直接字符格式一起复制过去，目录里那条就变成
# 黑体了（用户实测发现的 bug）。进目录靠的是 `styles._HEADING_OUTLINE_LEVELS`
# 里给这个样式补的大纲级别（目录域的 `\u` 开关按大纲级别收集条目）。

# "段前分页"（另起一页）现在是样式自己的属性（Excel"样式"表的"段前分页"
# 列，styles.py 的 apply_style_spec 里应用），不在这一层单独判断——哪些
# 部件要强制换页，是排版偏好，跟字体字号一样该让用户在 Excel 里调，不该
# 按角色在代码里写死。


def _convert_anchor_to_inline(anchor) -> None:
    """把一个 `wp:anchor`（浮动图）就地转成 `wp:inline`（嵌入型）。

    只搬 inline 需要的子元素（extent/effectExtent/docPr/cNvGraphicFramePr/
    graphic），丢掉定位与环绕（positionH/positionV/wrap*）——图从此随文字流
    走，不再"飘"。距离属性（distT/B/L/R）原样带过去。
    """
    parent = anchor.getparent()
    inline = parent.makeelement(qn("wp:inline"), {})
    for attr in ("distT", "distB", "distL", "distR"):
        value = anchor.get(attr)
        if value is not None:
            inline.set(attr, value)
    for tag in ("wp:extent", "wp:effectExtent", "wp:docPr", "wp:cNvGraphicFramePr", "a:graphic"):
        child = anchor.find(qn(tag))
        if child is not None:
            inline.append(copy.deepcopy(child))
    parent.replace(anchor, inline)


def _zero_paragraph_indent(paragraph) -> None:
    """把段落缩进**强制清零**（写全 8 个属性，盖过样式里的缩进）。

    与封面模块的 `_clear_char_unit_indent`（只补缺）不同：这里就是要覆盖
    正文样式带来的两字符首行缩进——图片要居中，不能带缩进。
    """
    pPr = paragraph._p.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        spacing = pPr.find(qn("w:spacing"))
        jc = pPr.find(qn("w:jc"))
        if spacing is not None:
            spacing.addnext(ind)
        elif jc is not None:
            jc.addprevious(ind)
        else:
            pPr.append(ind)
    for attr in (
        "w:left", "w:leftChars", "w:firstLine", "w:firstLineChars",
        "w:hanging", "w:hangingChars",
    ):
        ind.set(qn(attr), "0")


def normalize_image_paragraphs(document: DocumentObject,
                              skip_indices: frozenset[int] | set[int] = frozenset()) -> None:
    """图片段落规范化（2026-09-13 用户实拍"图片不知道飞到哪去了"）：

    1) **嵌入型**：非"衬于文字下方"的浮动图（`wp:anchor`）转成 `wp:inline`
       ——图片回到文字流里随文排布，不再飘（夹具 `图表脚注测试.docx` 里就
       有一张 wrapNone + 相对栏绝对偏移的浮图）；
    2) **单倍行距**：图片段落会继承正文样式，而正文样式的行距是"**固定值
       20 磅**"——固定行距会把嵌入图**裁掉**（test2.docx 实拍：145 磅高的
       图在成品里整个"消失"）。这里直接写 `line=240 auto`（单倍）盖掉它，
       并顺带清零缩进（正文样式的两字符首行缩进会把图推歪）；
    3) **居中**：插图排版要求。

    VML 旧式图（`w:pict`）只认不做：转换风险高，且它多半是文本框/装饰，
    不动它、也不给所在段落改格式，遇到问题再人工处理。

    ⚠️ **只处理"这一段除了图没别的文字"的段落**（用户 2026-09-16 裁决，方案 C）：
    图和文字挤在同一段的（`民总.docx` 那张浮动图就挂在整段正文上），
    **整段一律不碰**——转成嵌入型会把图塞进那段正文中间，居中又会让整段正文
    歪掉（用户实拍："正文内的标题没有真正居中"之外，那一段正常正文被整段居中
    了）。而且浮动图本来就没有"文字流里的位置"，位置只能由作者自己负责；
    把图挪到宿主段落里去只会更不可预测。

    `skip_indices`：用户指定的封面区间——**封面里的图不碰**（那是用户自己
    排的版，改成嵌入型/居中/单倍行距只会把封面弄乱）。
    """
    for index, paragraph in enumerate(document.paragraphs):
        if index in skip_indices:
            continue
        anchors = list(paragraph._p.iter(qn("wp:anchor")))
        inlines = list(paragraph._p.iter(qn("wp:inline")))
        if not anchors and not inlines:
            continue
        if paragraph.text.strip():
            # 图和文字同段：不动图、也不动这一段的格式（见 docstring 里方案 C）。
            # `paragraph.text` 只看段落自己的 run，所以文本框里的字不算。
            continue

        for anchor in anchors:
            if anchor.get("behindDoc") == "1":
                continue  # 衬于文字下方 = 有意为之（水印式），不动
            _convert_anchor_to_inline(anchor)

        paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.line_spacing = 1.0
        _zero_paragraph_indent(paragraph)


# 标签部件 -> (标签样式, 内容样式)。键用识别层（recognize.py）的规范部件名——
# 不再是字面标签"Abstract"/"Keywords"（那是旧 detect.py 的产物）。目录/三张
# 清单没有内容样式——它们的"内容"是 Word 域，不是现在要套格式的段落。
_LABEL_CONTENT_STYLE_IDS: dict[str, tuple[str, str | None]] = {
    "摘要": ("S07", "S08"),
    "关键词": ("S09", "S10"),
    "英文摘要": ("S12", "S13"),
    "英文关键词": ("S30", "S31"),
    "目录": ("S14", None),
    # 三张清单的标题"参照目录"排（用户 2026-09-12 的说法）：黑体三号居中、
    # 自带"段前分页"（满足规范 [93]"另起一页"）。所以直接复用 S14 的格式—
    # 用户在 Excel 里改"目录标题"，三张清单标题跟着一起变；哪天要独立调整，
    # 再拆出单独的样式行。
    "插图清单": ("S14", None),
    "附表清单": ("S14", None),
    "插图和附表清单": ("S14", None),
}


def label_style_id(part: str) -> str | None:
    """这个部件的**标签**套预设样式表里的哪一行（如 "S02"）；不是标签类部件返回 None。
    界面（格式页「标签写法」）用它算"标签前另起一页"跟随样式时的默认值。"""
    if part in _ANCHOR_STYLE_IDS:
        return _ANCHOR_STYLE_IDS[part]
    if part in _LABEL_CONTENT_STYLE_IDS:
        return _LABEL_CONTENT_STYLE_IDS[part][0]
    return None


def _apply_label_break(paragraph, part: str, label_page_break: dict[str, bool] | None) -> None:
    """标签前"另起一页"的按部件覆盖（预设 `label_page_break`，用户 2026-09-20）：
    写了这个部件就把段落的"段前分页"直接设成对应值（显式的 False 会盖过样式里的 True）；
    没写就什么都不做——跟随样式表里那一格。"""
    if label_page_break and part in label_page_break:
        paragraph.paragraph_format.page_break_before = bool(label_page_break[part])


def apply_from_index(
    document: DocumentObject,
    index: list[Entry],
    styles: StyleRegistry,
    *,
    skip_indices: set[int] = frozenset(),
    renumber_references: bool = True,
    reference_style: str | None = None,
    label_page_break: dict[str, bool] | None = None,
) -> None:
    """**2026-09-18 新架构**：识别层（`recognize.py`）已经把每一段的身份写进
    `Entry.part`/`Entry.meta`，这里只剩一件事——按身份把预设格式套上去。

    取代的是旧架构里 `apply_headings`/`apply_title`/`apply_english_title`/
    `apply_english_subtitle`/`apply_front_matter_labels`/
    `apply_anchor_content_styles` 六个函数——那些函数各自重新算一遍"内容范围
    到哪里为止"，而 `recognize.recognize()` 已经在识别阶段把标签/内容两段式
    的归属算好、写进了 `Entry.meta["slot"]`（"label"/"content"），这里直接
    按 slot 分派，不需要再找边界。

    `skip_indices`：图题/表题段落下标——不能吃"内容样式"（多是直接格式，
    会盖掉图题/表题自己的居中、行距，2026-09-12 在夹具上实测过）。

    参考文献条目的自动编号逻辑（`_make_auto_numbered_reference`）原样保留，
    只是改成按 `Entry` 遍历，不再依赖 `HeadingMatch` 列表。
    """
    ordinals = paragraph_ordinals(index)
    paragraphs = document.paragraphs
    reference_numbering_id: int | None = None
    # 参考文献条目**自己**当前已经在用的 numId（改动任何一条之前先扫一遍）：
    # `ensure_reference_numbering` 只在这个集合里找"格式已经一致，可以直接
    # 复用"的定义，不再拿正文里其他毫不相干的编号列表凑数（见该函数说明）。
    reference_existing_num_ids = {
        nid for i, e in ordinals.items()
        if e.part == "参考文献" and e.meta.get("slot") == "content"
        and (nid := paragraph_num_id(paragraphs[i])) is not None
    }
    # 附录"第一段是序号题名（S20）、第二段起当正文（S01）"——靠"是不是附录
    # 标签之后第一条内容"来判断，不需要 recognize.py 单独识别序号题名。附录
    # 全篇只出现一次（2026-09-19 起，见 recognize.py::_scan_backward）。
    appendix_first_content_seen = False

    for i in range(len(paragraphs)):
        entry = ordinals.get(i)
        if entry is None or not entry.part or i in skip_indices:
            continue
        part = entry.part
        slot = entry.meta.get("slot")

        if part == "题目":
            apply_spec(paragraphs[i], styles, "S06")
        elif part == "副标题":
            apply_spec(paragraphs[i], styles, "S29")
        elif part == "英文题目":
            apply_spec(paragraphs[i], styles, "S11")
        elif part == "英文副标题":
            # S32 跟中文副标题（S29）对称，但西文字体不同——S29 是
            # `east=黑体 ascii=黑体`，英文拿它排出来是黑体的英文字
            # （2026-09-15 查预设实测），不能复用。
            apply_spec(paragraphs[i], styles, "S32")
        elif part in ("一级标题", "二级标题", "三级标题", "四级标题"):
            apply_spec(paragraphs[i], styles, _LEVEL_STYLE_IDS[entry.meta["level"]])
        elif part in _ANCHOR_STYLE_IDS and slot == "label":
            # 引言/结论/参考文献/致谢/附录——标签本身套锚点样式。
            apply_spec(paragraphs[i], styles, _ANCHOR_STYLE_IDS[part])
            _apply_label_break(paragraphs[i], part, label_page_break)
        elif part in _LABEL_CONTENT_STYLE_IDS and slot == "label":
            label_style_id, _content_style_id = _LABEL_CONTENT_STYLE_IDS[part]
            apply_spec(paragraphs[i], styles, label_style_id)
            _apply_label_break(paragraphs[i], part, label_page_break)
        elif slot == "content" and part in _LABEL_CONTENT_STYLE_IDS:
            _label_style_id, content_style_id = _LABEL_CONTENT_STYLE_IDS[part]
            if content_style_id is not None:
                apply_spec(paragraphs[i], styles, content_style_id)
        elif slot == "content" and part == "附录":
            # 规范 [226]："附录内容文字与图表编排样式可参照正文"——第一段是
            # 序号题名（S20），第二段起当正文（S01）。
            if not appendix_first_content_seen:
                appendix_first_content_seen = True
                apply_spec(paragraphs[i], styles, "S20")
            else:
                apply_spec(paragraphs[i], styles, "S01")
        elif slot == "content" and part in _ANCHOR_CONTENT_STYLE_IDS:
            apply_spec(paragraphs[i], styles, _ANCHOR_CONTENT_STYLE_IDS[part])
            if part == "参考文献" and renumber_references and paragraphs[i].text.strip():
                # 空段落跳过：标题和历史条目之间常常夹着空行，挂上编号就会
                # 多出一个孤零零的"[3]"（Word 实测过）。有字的才当条目。
                if reference_numbering_id is None:
                    reference_numbering_id = ensure_reference_numbering(
                        document, reference_style, existing_ids=reference_existing_num_ids)
                _make_auto_numbered_reference(
                    paragraphs[i], reference_numbering_id, document)


# 目录条目的层级 → 内置 toc 样式（预设 S15/S16/S17）。
# 判据用**条目文字开头的编号**：Word 生成目录时会把标题自己的编号原样抄进条目
# （"一、""（一）""1."）；无编号的（引言/结论/参考文献/附录/致谢）按规范 [90]
# 与章标题同一层级，也算第一档。规范 [202]：章标题居左、一级节标题缩进 1 字符、
# 二级节标题缩进 2 字符。
_TOC_ENTRY_LEVEL_PATTERNS: tuple[tuple[re.Pattern, str], ...] = (
    (re.compile(r"^\s*[（(]\s*[〇零一二三四五六七八九十百廿卅]+\s*[）)]"), "S16"),
    (re.compile(r"^\s*[〇零一二三四五六七八九十百廿卅]+\s*[、.]"), "S15"),
)
# 点分方案（"1"/"1.1"/"1.1.1"……，见 `recognize.ArabicDotHeadingGrammar`）：
# 各级标题都是"数字+点"的写法，只看有没有点分不出层级——以前一律当成三级
# （S17），点分稿下一二三级标题在目录里全部挤成同一档，不是"分级分错"，是
# 根本没分（2026-09-22 审计揪出的真 bug）。改成数一数编号里有几段数字（用
# 点分开），1 段＝一级、2 段＝二级、3 段及以上＝三级——跟上面中文方案"一、"＝
# 一级、"（一）"＝二级、纯数字＝三级的分档原则一致，只是这里三档全用数字。
_TOC_DOT_NUMBER_RE = re.compile(r"^\s*(\d{1,2}(?:\.\d{1,2}){0,3})\.?(?:\s|$)")
_DOT_LEVEL_TO_STYLE = {1: "S15", 2: "S16"}


def toc_entry_style_id(text: str) -> str:
    """目录条目该挂哪一档目录样式（S15=目录一级 / S16=二级 / S17=三级）。"""
    for pattern, style_id in _TOC_ENTRY_LEVEL_PATTERNS:
        if pattern.match(text):
            return style_id
    dot_match = _TOC_DOT_NUMBER_RE.match(text)
    if dot_match:
        depth = dot_match.group(1).count(".") + 1
        return _DOT_LEVEL_TO_STYLE.get(depth, "S17")
    return "S15"           # 没编号的（引言/结论/参考文献/致谢/附录）＝第一档


def apply_existing_toc_entry_styles(document: DocumentObject,
                                    styles: StyleRegistry) -> int:
    """给**已有**目录域渲染出来的那些条目套上 toc 1/2/3（预设 S15/S16/S17）。

    用户 2026-09-16 实拍 `民总.docx`：那份目录是 Word 的"自动目录"内容控件，
    **条目段落的 `pPr` 完全是空的**（既没有 `pStyle`、也没有缩进，见
    `tools/_probe_toc_entry_fmt.py`）——而我们的做法只是把预设落到内置
    `toc 1/2/3` 样式上，条目根本不挂那个样式 ⇒ 一条都没被格式化（用户原话：
    "条目格式设置不成功"）。Word 自己刷新目录时会给条目补上 toc N，可那份稿子是
    WPS 系工具生成的、打开也未必刷新，所以只能我们自己按条目文字判断层级挂上去。

    只碰**目录域区间**里的段落（`doc_index.toc_field_span_indices`）：域指令
    那几段（没文字）跳过，已经有 toc 样式的也不动。返回挂了几段。
    """
    paragraphs = document.paragraphs
    span = toc_field_span_indices(paragraphs, {"目录"})
    if not span:
        return 0
    touched = 0
    for i in sorted(span):
        paragraph = paragraphs[i]
        text = paragraph.text.strip()
        if not text:
            continue                     # 域指令/结束那几段没有文字，不用套
        if is_toc_paragraph(paragraph):
            continue                     # 已经挂着 toc N 了，样式已经对它生效
        apply_spec(paragraph, styles, toc_entry_style_id(text))
        touched += 1
    return touched


def apply_long_quotes(document: DocumentObject, quote_indices: list[int],
                      styles: StyleRegistry) -> None:
    """给标成"长引文"的段落套 S18 样式（楷体四号、两端对齐、首行缩进 2 字符、
    左缩进 2 字符、固定行距 20 磅——具体数值全部来自预设 Excel/JSON 里的
    S18 那一行，这里不写死）。

    长引文没有可靠的自动识别信号（不像标题/图表题能从编号、"图/表+数字"这类
    文字模式判断），所以完全靠用户在"校对和调整识别结果"窗口里手动勾选
    ☐引用——`quote_indices` 只会来自那份人工调整结果（`ManualStructure.
    quote_paragraphs`），这里只管套格式，不做任何自动探测。
    """
    for i in quote_indices:
        apply_spec(document.paragraphs[i], styles, "S18")


def apply_appendix_entries(document: DocumentObject, entries, styles: StyleRegistry) -> int:
    """给附录区间里 `附录A …` 这些段落套 S20（附录序号题名）。

    用户 2026-09-15："附录A-Z识别要做"。识别在 `detect.detect_appendix_entries`
    （范围限制很死），这里只管套样式。**故意放在 `apply_anchor_content_styles`
    之后**：那一步会把附录整区的正文套成 S01，先套 S20 会被它盖掉。只碰段落
    下标，不插入/删除段落。
    """
    count = 0
    for index, _label in entries:
        if index < len(document.paragraphs):
            apply_spec(document.paragraphs[index], styles, "S20")
            count += 1
    return count


def apply_table_text_styles(document: DocumentObject, styles: StyleRegistry,
                           skip_tables: frozenset[int] | set[int] = frozenset()) -> None:
    """给**表格单元格里的文字**套格式（规范 [218]：表单元格中文字未体五号、
    单倍行距、段前段后各 3 磅）。

    为什么要单独一个函数：整条管线原来只走 `document.paragraphs`，而那个属性
    **看不到表格里的段落**——表格是块级元素，不是段落。所以表格文字一直没被
    格式化过（只有 cleanup 碰过表格，那是清格式不是套格式）。

    ！注意：合并单元格会被 `row.cells` **重复返回**，所以这里只能做幂等的
    操作（套格式），绝对不能插入/删除段落。

    `skip_tables`：用户指定的封面区间里的表格（`cleanup.tables_in_prefix()`
    算出来的**元素对象本身**，不是 id——原因见那个函数）——封面不动。

    **样式是「题注/Caption/图题/表题」的单元格段落跳过**（2026-09-15）：题注
    经常写在表格上方/下方的单元格里，统一套 S26 会把它的题注样式盖掉
    （实测 `图表压力测试.docx` 表 7 首格那个「图 1」）。
    """
    for table in _iter_tables(document.tables):
        if table._tbl in skip_tables:
            continue
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    style = paragraph.style
                    if style is not None and str(style.name).strip().lower() \
                            in _CAPTION_LIKE_STYLES:
                        continue
                    apply_spec(paragraph, styles, "S26")


def repeat_table_header_rows(document: DocumentObject,
                            skip_tables: frozenset[int] | set[int] = frozenset()) -> int:
    """给每个表格的第一行加"标题行重复"（`w:tblHeader`）。返回处理了几个表。
    规范要求"如果某个表需要转页接排……续表均应重复表头和关于单位的陈述"——
    把标题行带到下一页，正是 Word 的"重复标题行"（`w:tblHeader`）干的事。

    **这是"续表"这件事里唯一能纯 python 完全自动化的部分**。"表1（续）"
    那一行做不到，原因见 README：跨不跨页是排版结果，而插入那一行本身又会
    反过来改变分页，两边互相影响、天然需要迭代。

    只给第一行加：Word 允许把前几行都设成标题行，但"第一行是表头"是绝大多数
    论文表格的情形，多设反而会让没有表头的表平白多重复一行。
    """
    count = 0
    for table in _iter_tables(document.tables):
        if table._tbl in skip_tables:
            continue
        rows = table.rows
        if not rows:
            continue
        trPr = rows[0]._tr.get_or_add_trPr()
        if trPr.find(qn("w:tblHeader")) is not None:
            count += 1
            continue

        tbl_header = OxmlElement("w:tblHeader")
        # CT_TrPr 里的顺序：tblHeader 排在 trHeight 之后、tblCellSpacing 之前。
        before_tbl_header = {
            qn(tag)
            for tag in (
                "w:cnfStyle",
                "w:divId",
                "w:gridBefore",
                "w:gridAfter",
                "w:wBefore",
                "w:wAfter",
                "w:cantSplit",
                "w:trHeight",
            )
        }
        index = 0
        for position, child in enumerate(trPr):
            if child.tag in before_tbl_header:
                index = position + 1
            else:
                break
        trPr.insert(index, tbl_header)
        count += 1
    return count


# 表格遍历（含嵌套）**只有一份实现**：`cleanup.iter_tables`——清直接格式那边
# 也得按同一批表格走（2026-09-16 审计报告 Bug#3：嵌套表原来"只套样式、
# 不清直接格式"）。留个薄壳，免得两个调用点都改名。
_iter_tables = iter_tables


# 「题注/图题/表题」这几类样式名——**认不出来的题注段落一律不许动**。
# 用户 2026-09-15：题注系统里什么写法都可能出现（`图表 2`、`公式 1`、空文字的
# 图片题注……），工具只认 图/表 两类，**其余的不是错、是"不归我管"**。
# 但管线里有一堆"按范围给正文套格式"的步骤（引言内容、附录正文、表格文字），
# 它们会把这类段落的样式直接换成正文——实测：`图表压力测试.docx` 的 `公式 1`
# 和那个只有图片的题注段，样式就是这样没的（pStyle 被清空 ⇒ 掉成正文）。
# 所以把它们的段号算成一个"别碰名单"，挂到那几步上。
_CAPTION_LIKE_STYLES = ("题注", "caption", "图题", "表题")


def caption_style_indices(document: DocumentObject) -> set[int]:
    """样式是「题注/Caption/图题/表题」的段落下标 —— 给"范围套正文"那些步骤当
    跳过名单用。

    ⚠️ **绝不能拿它去拦 `apply_captions()`**：识别出来的题注正是要挂上
    「图题/表题」样式的，一拦，清单域（按样式名收）就空了。
    """
    found: set[int] = set()
    for i, paragraph in enumerate(document.paragraphs):
        style = paragraph.style
        if style is not None and str(style.name).strip().lower() in _CAPTION_LIKE_STYLES:
            found.add(i)
    return found


def apply_captions(
    document: DocumentObject,
    captions: list[CaptionMatch],
    styles: StyleRegistry,
) -> None:
    r"""给图题/表题套格式（规范 [217] 图序与图名、[218] 表序与表名，两者格式
    一模一样：宋体五号居中、单倍行距、段前 6 磅段后 12 磅）。

    **这里只做两件事：挂样式、套格式。文字一个字不改、域一个不碰、编号一律不
    动**（用户 2026-09-16：“我们不管域自己，我们只管样式……我们也不对用户的域
    重新编号”）。以前那套“手打编号→SEQ 域”的题注化（`apply_caption_numbering`）
    已整体删除。

    **中英两行分开处理**：中文行挂真样式（S24/S25），英文行（`Fig 1 …` /
    `Table 1 …`，见 CaptionMatch.is_english）只把同样的格式**直接写上去**。
    原因很实在："插图清单/附表清单"域是按**样式名**收集题注的，英文行要是也
    挂着图题/表题样式，中英对照的文档里一条图会变成两条目。直接格式看起来
    一模一样，但域收集看不见它。改这段前先想清楚这一点。

    **续行（`表 1（续）`）走英文行同一路**（用户 2026-09-15 裁决）："续图和续表
    只识别出表题，一律不计列表"。套了样式就会在附表清单里多出一条"表 1（续）"，
    所以只写直接格式。其余（中文图题/表题，**含分图**）都挂样式 ⇒ 进清单——
    挂不挂样式就是“进不进清单”本身（用户 2026-09-16：分图不区分，每张都进列表）。
    """
    for caption in captions:
        style_id = "S24" if caption.kind == "图" else "S25"
        paragraph = document.paragraphs[caption.paragraph_index]
        if not caption.wears_caption_style:
            write_spec_onto_paragraph(paragraph, styles.spec(style_id))
        else:
            apply_spec(paragraph, styles, style_id)


def _run_is_inline_picture(run_el) -> bool:
    """这个 run 是不是"嵌入型图片"（有 `wp:inline`、没有文字、没有浮动锚点）。"""
    if run_el.find(qn("w:drawing")) is None:
        return False
    if any(run_el.iter(qn("wp:anchor"))):
        return False
    return not "".join(t.text or "" for t in run_el.iter(qn("w:t"))).strip()


def split_pictures_out_of_captions(paragraphs: list[Paragraph], styles: StyleRegistry) -> int:
    r"""图题/表题段落里夹着的**嵌入型图片**挪到它前面单独一段，返回拆了几段。

    原因：清单域（`TOC \t "图题,1"`）按样式收**整段**，图和图题写在同一段里时，
    图片会被一起搬进清单条目（用户 2026-09-19 实测）。图题段本身只留文字。
    图片段按图片段规矩排：居中、单倍行距、缩进清零，用正文样式。
    图片后面紧跟的、只装一个换行的 run 一并去掉（那是"图片+换行+图题"写法的
    换行，图题段里留着就会多出一个空行）。浮动图不动（位置由作者负责）。

    ⚠️ 只该在**重建之后、按重建后的段落**调用（插段会让后面的段落序号后移，
    所以内部从后往前处理）。
    """
    split = 0
    for paragraph in reversed(paragraphs):
        pictures = [r for r in paragraph._p.findall(qn("w:r")) if _run_is_inline_picture(r)]
        if not pictures:
            continue
        holder = Paragraph(OxmlElement("w:p"), paragraph._parent)
        paragraph._p.addprevious(holder._p)
        apply_spec(holder, styles, "S01")
        for run_el in pictures:
            holder._p.append(run_el)             # 移动（不复制）
        # 图片后面残留的纯换行 run / 开头的空 run 清掉
        for run_el in list(paragraph._p.findall(qn("w:r"))):
            has_text = "".join(t.text or "" for t in run_el.iter(qn("w:t"))).strip()
            only_break = (run_el.find(qn("w:br")) is not None and not has_text
                          and run_el.find(qn("w:drawing")) is None)
            if only_break:
                paragraph._p.remove(run_el)
            elif has_text:
                break
        holder.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        holder.paragraph_format.line_spacing = 1.0
        holder.paragraph_format.keep_with_next = True   # 图和图题别被分页拆开
        _zero_paragraph_indent(holder)
        split += 1
    return split


# 再画一个，变成 `[1][1] 鲁晓明：……`。
_MANUAL_REF_NUMBER = re.compile(r"^\s*[\[［]\s*\d+\s*[\]］]\s*")


def _strip_paragraph_prefix(paragraph: Paragraph, pattern: re.Pattern[str]) -> None:
    """把段落开头匹配 `pattern` 的那一段文字删掉。

    按 run 逐个消耗，只改文字、不删 run 元素：Word 里 `[1]` 很可能被拆成好几个
    run（改过字体/拼写检查都会拆），只清第一个 run 是删不干净的；而直接把
    后面的 run 全删又会丢掉它们身上的格式，跟 `normalize.py` 里"尽量少重排
    run"那条规矩相反。

    ⚠️ **超链接段落必须走 `iter_inner_content()`**（2026-09-22 审计揪出的真
    bug）：`paragraph.text` 包含超链接的可见文字（`w:hyperlink` 里的 run），
    但 `paragraph.runs` **不包含**——只算 `w:p` 直属的 `w:r`。以前用
    `paragraph.text` 算前缀长度、却按 `paragraph.runs` 消耗字符，手打序号
    整个或部分落在超链接里时两边对不上号：序号剥不干净（残留半截）或者
    多吃掉正文开头几个字。知网复制的参考文献条目常把整条做成超链接，最容易
    踩上。改成把超链接内的 run 当普通 run 一样纳入同一条消耗序列，跟
    `paragraph.text` 的定义（含超链接文字）保持一致。
    """
    match = pattern.match(paragraph.text)
    if match is None:
        return
    remaining = match.end()
    runs: list[Run] = []
    for item in paragraph.iter_inner_content():
        runs.extend(item.runs if isinstance(item, Hyperlink) else [item])
    for run in runs:
        text = run.text
        if not text:
            continue
        if remaining >= len(text):
            remaining -= len(text)
            set_run_text(run, "")
        else:
            set_run_text(run, text[remaining:])
            remaining = 0
        if remaining == 0:
            break


def reference_display_text(text: str, index: int, style: str | None = None) -> str:
    """参考文献条目**在成品里会显示成什么**（结构页显示用，只算字符串）。

    成品里条目序号是 Word 自动编号画的（编号定义按预设写法建，见
    `numbering.reference_level_text`），手打的 `[1]`/`［1］` 会被剥掉
    （`_MANUAL_REF_NUMBER`）。所以显示 ＝ 编号 + 空格 + 剥掉前缀后的正文。
    `style` ＝ 预设 `citation.reference_numbering_style`。
    ⚠️ Word 那边序号后跟的是**制表符**（编号定义没写 `w:suff`，Word 默认用
    制表符），这里用空格近似——结构页只要让人看出"序号会被换成第几条"。
    """
    prefix = reference_level_text(style).replace("%1", str(index))
    return f"{prefix} {_MANUAL_REF_NUMBER.sub('', text)}"


def _make_auto_numbered_reference(
    paragraph: Paragraph,
    numbering_id: int,
    document: DocumentObject | None = None,
) -> None:
    """把一条参考文献条目改成 Word 自动编号。

    三种情况：
      * 本来没有自动编号：剥掉手打的序号，挂上 `w:numPr`；
      * 已经有、而且指的就是我们那个定义：不动；
      * **已经有、但指的是别的定义**（学生自己用 Word 自动编号排的，
        `民总-排好1版.docx` 实测是 `%1.` 即 "1. 2. 3."）：**改指到我们的
        定义**——这样预设选 `[1][2][3]` 才会真的变成方括号（用户 2026-09-16：
        "当用户直接用自动编号排参考文献时，重新编号失效，没有调整为指定编号
        格式"）。改指之前把旧定义里的悬挂缩进搬过来（`numbering.copy_level_indent`），
        否则条目会顶着左边距、转行不对齐。

    为什么要自动编号：参考文献是全文最常增删的一节，序号手打的话每加一条都
    得把后面重编一遍；挂上 numPr 之后用户在 Word 里回车接着写就会自己排。
    """
    pPr = paragraph._p.get_or_add_pPr()
    existing = pPr.find(qn("w:numPr"))
    if existing is not None:
        current = paragraph_num_id(paragraph)
        if current is None or current == numbering_id:
            return
        if document is not None:
            copy_level_indent(document, current, numbering_id)
        num_id = existing.find(qn("w:numId"))
        if num_id is None:
            num_id = OxmlElement("w:numId")
            existing.append(num_id)
        num_id.set(qn("w:val"), str(numbering_id))
        # 改指的同时把层级也钉回第 0 层——我们建的参考文献编号定义
        # （`_build_reference_abstract`）只有一层（ilvl=0）。这一段原来的
        # `w:numPr` 如果层级不是 0（比如原来是别的多级列表里嵌套的一项），
        # 只改 numId、不碰 ilvl 的话，段落会指向我们定义里根本不存在的那一
        # 层，Word 那边这个号可能就不出来了（2026-09-22 审计揪出的：疑似，
        # 未在 Word 里实机复核）。
        ilvl = existing.find(qn("w:ilvl"))
        if ilvl is None:
            ilvl = OxmlElement("w:ilvl")
            existing.insert(0, ilvl)
        ilvl.set(qn("w:val"), "0")
        return

    _strip_paragraph_prefix(paragraph, _MANUAL_REF_NUMBER)

    numPr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), "0")
    num_id = OxmlElement("w:numId")
    num_id.set(qn("w:val"), str(numbering_id))
    numPr.append(ilvl)
    numPr.append(num_id)

    # numPr 在 pPr 里必须排在 pStyle 之后、spacing/ind/jc 之前（CT_PPrBase
    # 的顺序），不能直接 append——那会排到最后，Word 当成无效属性。
    before_numpr = {
        qn(tag)
        for tag in (
            "w:pStyle",
            "w:keepNext",
            "w:keepLines",
            "w:pageBreakBefore",
            "w:framePr",
            "w:widowControl",
        )
    }
    index = 0
    for position, child in enumerate(pPr):
        if child.tag in before_numpr:
            index = position + 1
        else:
            break
    pPr.insert(index, numPr)


# 规范里这几个部件出现的先后顺序，以及三个"清单"部件各自要收哪种题注——
# 这两份数据 2026-09-18 挪到 `doc_index.py`（`FRONT_ORDER`/`BACK_ORDER`/
# `LIST_PART_KINDS`），好让 `doc_index.auto_ops()` 和这个模块共用同一份
# 顺序知识，不用互相 import；这里直接从上面 import 进来的 `LIST_PART_KINDS`
# 使用即可（`build_part_paragraphs()` 下面要用它判断"清单要收哪种"）。

# 英文摘要页那两个标签读不到"插入文字"时的兜底写法（旧预设只有一格装两行）。
_ENGLISH_ABSTRACT_FALLBACK_LABELS = ("【ABSTRACT】", "【KEYWORDS】")

# 占位语（往缺失部件的位置插的那些提示语）的**原文集中在
# `essay_tool/engine/placeholders.py`**：识别阶段（`detect`）要认得出"这行是
# 工具自己插的"，两边必须同一份定义。下面是本地别名，用法完全没变。
_PLACEHOLDER_PROMPTS = PLACEHOLDER_PROMPTS
_CONTENT_PLACEHOLDER_PROMPTS = CONTENT_PLACEHOLDER_PROMPTS
_ENGLISH_PLACEHOLDER_PROMPTS = ENGLISH_PLACEHOLDER_PROMPTS
_ANCHOR_CONTENT_PROMPTS = ANCHOR_CONTENT_PROMPTS
_REFERENCE_CONTENT_PROMPT_PLAIN = REFERENCE_CONTENT_PROMPT_PLAIN
_APPENDIX_BODY_PROMPT = APPENDIX_BODY_PROMPT

# 引言/结论/参考文献/附录/致谢缺失时，标签下面也要占一段位——不能只补标签、
# 不补内容，那样用户很容易忘了这里还得自己写。跟摘要/关键词一样，内容是
# 论文的实质部分，没法自动生成，只能提示占位；用各自内容本来该用的样式
# （参考文献条目、附录序号题名、致谢正文），引言/结论没有专门的内容样式，
# 就是普通正文。
_ANCHOR_CONTENT_STYLE_IDS = {
    "引言": "S01",
    "结论": "S01",
    "参考文献": "S21",
    "附录": "S20",
    "致谢": "S22",
}


def _make_paragraph(document: DocumentObject, text: str, styles: StyleRegistry,
                    style_id: str) -> Paragraph:
    """**造**一段（**只造、不定位**）。

    构造方式跟 `Paragraph.insert_paragraph_before(text)` 一样（建 `w:p` + 一个装
    文字的 run），只是**不挂进文档**——挂到哪儿由调用方决定。样式不走 python-docx
    的 `style=`，而是 `apply_spec`（有的角色根本不建 Word 样式、直接把格式写到
    段落上，见 `styles._DIRECT_FORMAT_BASE`）。
    """
    paragraph = Paragraph(OxmlElement("w:p"), document)
    if text:
        paragraph.add_run(text)
    apply_spec(paragraph, styles, style_id)
    return paragraph


def _make_field_paragraph(document: DocumentObject, kind: str) -> Paragraph:
    """**造**一个"清单域"段落（**只造、不定位**）；有意不套样式（同旧实现）。"""
    paragraph = Paragraph(OxmlElement("w:p"), document)
    paragraph._p.append(make_list_field(kind))
    return paragraph


def _english_labels(insert_text: dict[str, str]) -> tuple[str, str]:
    """英文摘要页两个标签的字面文字：（摘要标签, 关键词标签）。

    2026-09-14 起"英文关键词"是独立部件、有自己的一格 `insert_text`；旧预设
    里只有"英文摘要"一格、里面装两行（上行摘要标签、下行关键词标签），所以
    关键词那儿保留"回头读英文摘要第二行"的兼容写法。
    """

    def lines(part: str) -> list[str]:
        raw = insert_text.get(part) or ""
        return [line.strip() for line in raw.splitlines() if line.strip()]

    ab = lines("英文摘要")
    kw = lines("英文关键词")
    abstract = ab[0] if ab else _ENGLISH_ABSTRACT_FALLBACK_LABELS[0]
    keywords = kw[0] if kw else (
        ab[1] if len(ab) > 1 else _ENGLISH_ABSTRACT_FALLBACK_LABELS[1])
    return abstract, keywords


def build_part_paragraphs(
    document: DocumentObject,
    part: str,
    insert_text: dict[str, str],
    styles: StyleRegistry,
    *,
    slot_col: str = "label",
    list_kinds: frozenset[str] | set[str] = frozenset(),
    renumber_references: bool = True,
    reference_style: str | None = None,
) -> list[Paragraph]:
    """部件 → **它由哪几段构成**（按成品里的先后顺序）。**不含任何位置判断。**

    这是索引法里"生成指令"展开的那一半：索引里一行 `生成 目录 的标签位`，执行器
    用它展开成"标签段 + 目录域段"，再按索引顺序往下追加。
    ⚠️ **这里不许出现 anchor／fallback／"紧贴上一个部件"这类词**——位置由
    `Op.out_index` 决定，本函数看不到也不该看到。

    `slot_col`：摘要/关键词这类部件在面板上有**两行**空位（标签位 + 内容位）。
    `"label"`＝造整个部件；`"text"`＝只造"内容"那几段（标签已经在文档里，
    对应 `_insert_content_placeholder` 那条路）；`"label_only"`＝只造标签那一段
    （同一部件另有"内容位"指令造内容，两条合起来才是完整部件——不这样分，内容占位语
    会被造两遍，2026-09-20 用户报的 bug）。
    """
    if slot_col == "label_only":
        # 标签永远是"整个部件"的第一段（摘要/关键词/英文摘要/英文关键词/附录/
        # 引言·结论·参考文献·致谢皆如此），所以取整个部件的第一段即可
        return build_part_paragraphs(document, part, insert_text, styles, slot_col="label",
                                     list_kinds=list_kinds, renumber_references=renumber_references,
                                     reference_style=reference_style)[:1]

    if part in _PLACEHOLDER_PROMPTS:
        style_id = "S06" if part == "题目" else "S29"
        return [_make_paragraph(document, _PLACEHOLDER_PROMPTS[part], styles, style_id)]

    if part in ("摘要", "关键词"):
        label_style_id, content_style_id = _LABEL_CONTENT_STYLE_IDS[part]
        content = _make_paragraph(document, _CONTENT_PLACEHOLDER_PROMPTS[part],
                                  styles, content_style_id)
        if slot_col == "text":
            return [content]
        label = _make_paragraph(document, insert_text.get(part) or part,
                                styles, label_style_id)
        return [label, content]

    if part == "英文摘要":
        abstract, _keywords = _english_labels(insert_text)
        content = _make_paragraph(document, _ENGLISH_PLACEHOLDER_PROMPTS["英文摘要"],
                                  styles, "S13")
        if slot_col == "text":
            return [content]
        return [_make_paragraph(document, abstract, styles, "S12"), content]

    if part == "英文关键词":
        _abstract, keywords = _english_labels(insert_text)
        content = _make_paragraph(document, _ENGLISH_PLACEHOLDER_PROMPTS["英文关键词"],
                                  styles, "S31")
        if slot_col == "text":
            return [content]
        return [_make_paragraph(document, keywords, styles, "S30"), content]

    if part == "英文题目":
        return [_make_paragraph(document, _ENGLISH_PLACEHOLDER_PROMPTS["英文题目"],
                                styles, "S11")]

    if part == "英文副标题":
        return [_make_paragraph(document, _ENGLISH_PLACEHOLDER_PROMPTS["英文副标题"],
                                styles, "S32")]

    if part == "目录":
        return [_make_paragraph(document, insert_text.get("目录") or "目录", styles, "S14")]

    if part in LIST_PART_KINDS:
        # "要插哪几个域"由调用方算好传进来（pipeline 里先走一遍
        # alerts.check_list_parts）：空集合＝这行没启用、或没有题注可收。
        styles.ensure("S15")          # 条目走 toc 1 样式，缩进等配置靠它生效
        out: list[Paragraph] = []
        if slot_col != "text":
            out.append(_make_paragraph(document, insert_text.get(part) or part,
                                       styles, "S14"))
        if "图" in list_kinds:        # 规范 [91] 分开列时插图清单在前
            out.append(_make_field_paragraph(document, "图"))
        if "表" in list_kinds:
            out.append(_make_field_paragraph(document, "表"))
        return out

    if part == "附录":
        # 三段：附录标题（S19，带大纲级别、进目录靠它）＋ 附录序号题名（S20）
        # ＋ 正文占位（S01）。
        title = _make_paragraph(document, _ANCHOR_CONTENT_PROMPTS["附录"], styles, "S20")
        body = _make_paragraph(document, _APPENDIX_BODY_PROMPT, styles, "S01")
        if slot_col == "text":
            return [title, body]
        return [_make_paragraph(document, insert_text.get(part) or part,
                                styles, _ANCHOR_STYLE_IDS[part]), title, body]

    # 引言/结论/参考文献/致谢：标签 ＋ 一段内容占位提示（内容是论文的实质部分，
    # 没法自动生成，只能提示用户自己往里面写）。
    content_text = _ANCHOR_CONTENT_PROMPTS[part]
    if part == "参考文献" and not renumber_references:
        # 没开自动编号时不能还说"序号已经是自动编号"（用户 2026-09-15）。
        content_text = _REFERENCE_CONTENT_PROMPT_PLAIN
    content = _make_paragraph(document, content_text, styles,
                              _ANCHOR_CONTENT_STYLE_IDS[part])
    if part == "参考文献" and renumber_references:
        # 占位这一条也要挂上自动编号：不挂的话用户看到的是"没有序号的第一条"，
        # 而提示语却写着"序号已经是自动编号"，他回车接着写也接不上这个列表。
        _make_auto_numbered_reference(
            content, ensure_reference_numbering(document, reference_style), document)
    if slot_col == "text":
        return [content]
    return [_make_paragraph(document, insert_text.get(part) or part,
                            styles, _ANCHOR_STYLE_IDS[part]), content]


def build_generated_elements(document: DocumentObject, op, insert_text: dict[str, str],
                             styles: StyleRegistry, *,
                             list_kinds: frozenset[str] | set[str] = frozenset(),
                             renumber_references: bool = True,
                             reference_style: str | None = None,
                             label_page_break: dict[str, bool] | None = None) -> list:
    """把一条**生成指令**展开成要追加的元素（按成品里的先后顺序）。

    索引法里"生成指令"的那一半：一条指令可能是一段（题目占位语）、两段
    （摘要＝标签＋内容）、三段（附录＝标题＋序号题名＋正文占位），也可能是一个
    **域**（目录域／清单域——域是独立的指令，不是标签的一部分）。

    ⚠️ **只造、不定位**：返回的元素还没挂进文档，挂到哪儿由执行器按
    `Op.out_index` 顺序往下追加。
    ⚠️ 索引法路径下 `list_kinds` 传空（清单域由独立指令给）；旧"缺件"路径仍
    传 kinds（那边标签和域是一次造完的）。
    """
    kind = getattr(op, "kind", "placeholder")
    if kind == "toc_field":
        paragraph = Paragraph(OxmlElement("w:p"), document)
        paragraph._p.append(make_toc_field(styles))
        return [paragraph._p]
    if kind == "list_field":
        paragraph = Paragraph(OxmlElement("w:p"), document)
        paragraph._p.append(make_list_field(getattr(op, "text", "") or "图"))
        return [paragraph._p]
    part = getattr(op, "part", "") or ""
    known = (set(_PLACEHOLDER_PROMPTS) | {"摘要", "关键词", "英文题目", "英文副标题",
             "英文摘要", "英文关键词", "目录", "附录"} | set(_ANCHOR_CONTENT_PROMPTS)
             | set(LIST_PART_KINDS))       # 三个清单的标签段（域是独立指令）
    if part not in known:
        return []                      # 认不出的部件：宁可什么都不造，也不要造错
    slot_col = getattr(op, "slot_col", "") or "label"
    paragraphs = build_part_paragraphs(
        document, part, insert_text, styles,
        slot_col=slot_col,
        list_kinds=list_kinds, renumber_references=renumber_references,
        reference_style=reference_style)
    if paragraphs and slot_col in ("label", "label_only"):
        _apply_label_break(paragraphs[0], part, label_page_break)     # 标签永远是第一段
    return [p._p for p in paragraphs]

