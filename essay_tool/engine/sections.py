"""分节、页码、页眉、目录域。

现在只处理"前置部分"（摘要起至目录，罗马页码）和"正文部分"（引言起，
阿拉伯页码）两节——封面（cover.py）是在全部套版完成后才拼到文档最前面的
独立一节，封面节不放页眉，不归这里管，所以没有第三节要处理。

Excel"页眉页脚"小节那几行与节的对应关系：**从最后一节往前配**
（`apply_header_footer`，2026-09-16 修）——源文件里的分节数可能跟预设的
行数对不上（比如整篇只有一个节），从前往后配会把"前置部分"的页眉页码
套到全篇（实测过：单节文档整篇被套上前置页眉/罗马页码）。表格里的顺序
仍然是纸质文档上的节顺序，只是配对方向反过来。
"""

from __future__ import annotations

from docx.document import Document as DocumentObject
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from essay_tool.config.schema import HeaderFooterRule
from essay_tool.engine.styles import (
    StyleRegistry,
    set_header_bottom_border,
    set_header_line_extent,
)

_TOC_INSTR = 'TOC \\o "1-3" \\h \\z \\u'
_TOC_PLACEHOLDER = "（目录内容由 Word 自动生成，打开后右键选择「更新域」，或全选后按 F9 刷新）"


def split_body_section(document: DocumentObject, body_start_paragraph: Paragraph | None) -> None:
    """在 body_start_paragraph（引言所在段落，不论是识别出来的还是缺失后
    自动补上的）前面插入分节符，把文档切成"前置部分"和"正文部分"两节。

    用段落对象、找它在 XML 里的"上一个兄弟段落"来定位，而不是拿段落下标
    做减法——缺失部件可能会在这之前插入新段落，下标会变，但段落对象和
    它在 XML 树里的位置是稳的。body_start_paragraph 是 None，或者引言
    就是文档第一段（前面没有前置部分内容可分），就什么都不做，文档仍是
    单节。
    """
    if body_start_paragraph is None:
        return
    # 往回找最近的一个**段落**兄弟：分节符只能挂在段落的 pPr 上，而正文首标题
    # 前面那一个兄弟可能是表格（表格挂不了 sectPr）。以前遇到不是 w:p 就直接
    # return，结果是整篇只剩一节——前置部分的罗马页码、正文部分的页眉全都不
    # 生效，而且没有任何提示（审计 2026-09-16 报的）。
    prev_p = body_start_paragraph._p.getprevious()
    while prev_p is not None and prev_p.tag != qn("w:p"):
        prev_p = prev_p.getprevious()
    if prev_p is None:
        return
    body = document._element.body
    sentinel_sectPr = body.get_or_add_sectPr()
    front_matter_sectPr = sentinel_sectPr.clone()
    # `clone()` 连正文那一节的 `w:headerReference`/`w:footerReference` 也整个
    # 复制过去了——它们指的是**同一个** r:id，也就是同一份物理页眉/页脚部件。
    # `apply_header_footer` 后面会对这一节调 `is_linked_to_previous = False`，
    # 但 python-docx 的这个 setter 一看"这一节已经有引用了（哪怕是抄来的）"
    # 就当"早就不是链接状态"直接跳过、不会另建一份独立定义——于是前置节和
    # 正文节全程共用同一份页眉/页脚部件，改哪节的文字都会影响另一节，"这节
    # 不要页码"也压不住（2026-09-22 审计揪出的真 bug，实测过：两节 header
    # part id 完全相同）。这里清掉这几个引用，前置节回到真正"链接到前一节"
    # 的状态，后面 `apply_header_footer` 一断链接就会各自建一份独立定义。
    for tag in (qn("w:headerReference"), qn("w:footerReference")):
        for ref in front_matter_sectPr.findall(tag):
            front_matter_sectPr.remove(ref)
    prev_p.set_sectPr(front_matter_sectPr)


def _set_page_number_format(sectPr, fmt: str | None, start: int = 1) -> None:
    existing = sectPr.find(qn("w:pgNumType"))
    if existing is not None:
        sectPr.remove(existing)
    pgNumType = OxmlElement("w:pgNumType")
    if fmt:
        pgNumType.set(qn("w:fmt"), fmt)
    pgNumType.set(qn("w:start"), str(start))
    pgMar = sectPr.get_or_add_pgMar()
    pgMar.addnext(pgNumType)


def _is_page_field(instr: str) -> bool:
    """域指令是不是**页码**域。

    ⚠️ 不能写成 `"PAGE" in instr.upper()`：`PAGEREF`（交叉引用）里也含 PAGE，
    会被一起摘掉。域指令的第一个词就是域名，只认它。
    """
    tokens = (instr or "").strip().split()
    return bool(tokens) and tokens[0].upper() == "PAGE"


# 文本框/图形的宿主元素。Word 2007+ 写成 `w:drawing`，兼容形态是 `w:pict`
# （VML），**同一个框在文件里出现两遍**，成对裹在 `mc:AlternateContent` 里。
_MC_ALTERNATE = ("{http://schemas.openxmlformats.org/markup-compatibility/2006}"
                 "AlternateContent")
_VML_IMAGEDATA = "{urn:schemas-microsoft-com:vml}imagedata"


def _shape_host(node, root):
    """从 node 往上找"这个文本框整体"的最外层元素（AlternateContent/drawing/pict）。"""
    host = None
    cur = node
    while cur is not None and cur is not root:
        if cur.tag == _MC_ALTERNATE:
            return cur
        if cur.tag in (qn("w:drawing"), qn("w:pict")):
            host = cur
        cur = cur.getparent()
    return host


def _txbx_empty(host) -> bool:
    """这个（或这几个）文本框里还剩下看得见的东西吗（文字 / 图片）。"""
    for box in host.iter(qn("w:txbxContent")):
        if "".join(t.text or "" for t in box.iter(qn("w:t"))).strip():
            return False
        if list(box.iter(qn("a:blip"))):
            return False
        for img in box.iter(_VML_IMAGEDATA):
            if img.get(qn("r:id")):
                return False
    return True


def _strip_page_fields(root) -> int:
    """摘掉 root 底下所有**页码域**（fldChar 三段式），返回摘掉的个数。

    ⚠️ **必须按文档顺序跨 run 认域**：Word 写域时 `begin` 字符、`instrText`、
    `separate`、结果、`end` 字符分别住在**不同的 run** 里（民总页脚实测就是
    `begin` 一个 run、`PAGE \\* MERGEFORMAT` 另一个 run）。老写法只在"装着
    `begin` 的那个 run"里找指令 ⇒ 一条都认不出来 ⇒ 文本框里的旧页码原样留着
    （用户 2026-09-16："民总能稳定复现导出物有两个重合的页码"）。

    嵌套域（比如域结果里还有 PAGEREF）不开新层：内层 run 同时记进外层的区间，
    这样整个外层区间能被一起摘掉。
    """
    removed = 0
    stack: list[dict] = []          # 未闭合的域：{"instr": 指令原文, "runs": [元素]}
    for run in list(root.iter(qn("w:r"))):
        types = [c.get(qn("w:fldCharType")) for c in run.findall(qn("w:fldChar"))]
        instr = "".join(t.text or "" for t in run.findall(qn("w:instrText")))
        if "begin" in types:
            stack.append({"instr": "", "runs": [run]})
        else:
            if instr:
                for level in stack:
                    level["instr"] += " " + instr
            for level in stack:
                level["runs"].append(run)
        if "end" in types and stack:
            level = stack.pop()
            if _is_page_field(level["instr"]):
                for element in level["runs"]:
                    parent = element.getparent()
                    if parent is not None:
                        parent.remove(element)
                removed += 1
    return removed


def _has_page_field(p) -> bool:
    """这一段里有没有页码域（含拆分写法：只要有一段 instrText 里出现 PAGE 这个词）。"""
    for t in p.iter(qn("w:instrText")):
        tokens = [tok.strip("\\*") for tok in (t.text or "").split()]
        if "PAGE" in [tok.upper() for tok in tokens]:
            return True
    for fld in p.iter(qn("w:fldSimple")):
        if _is_page_field(fld.get(qn("w:instr")) or ""):
            return True
    return False


def _drop_page_number_fields(part) -> None:
    """摘掉页脚里已有的页码域（`w:fldSimple PAGE` 与 fldChar 三段式两种写法）。

    页脚里可能本来就有页码（用户自己插的，或者上次导出留下的）。不摘掉的话
    我们再插一个就变成**两个页码**（审计 2026-09-16 报的）。只动页码域，
    页脚里的其他文字原样留着。

    ⚠️ **必须扫整份页脚、包括文本框里面**（用户 2026-09-16 实拍民总："稳定
    复现导出物有两个重合的页码"）：那一份的页码是作者放在页脚里的**浮动文本框**
    （`w:drawing`/`w:pict` → `w:txbxContent` → 段落 → 域）。摘完之后，**空掉的
    文本框整体删掉**（不然留个空框，框上有边框的话会显出一个空方框）；框里还有
    别的文字或图片的不动。
    """
    root = part._element
    # 先记下"哪些文本框里装着页码域"——摘完要把空掉的框整个删掉
    hosts = set()
    for p in root.iter(qn("w:p")):
        if _has_page_field(p):
            host = _shape_host(p, root)
            if host is not None:
                hosts.add(host)
    _strip_page_fields(root)
    for fld in list(root.iter(qn("w:fldSimple"))):
        if _is_page_field(fld.get(qn("w:instr")) or ""):
            parent = fld.getparent()
            if parent is not None:
                parent.remove(fld)
    for host in hosts:
        parent = host.getparent()
        if parent is not None and _txbx_empty(host):
            parent.remove(host)


def _add_page_number_field(paragraph, prefix: str, suffix: str) -> None:
    if prefix:
        paragraph.add_run(prefix)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    run = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "1"
    run.append(t)
    fld.append(run)
    paragraph._p.append(fld)
    if suffix:
        paragraph.add_run(suffix)


# settings.xml 里必须排在 w:updateFields **之后**的元素（顺序照 python-docx
# 的 oxml/settings.py 里 CT_Settings._tag_seq 抄）。插 updateFields 时要插在
# 这里面第一个出现的元素前面，不能直接 append 到末尾——append 会把它排到
# w:compat/w:rsids/m:mathPr 后面，那个位置手调 Word 是绝不会产生的。
_AFTER_UPDATE_FIELDS = (
    "w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat", "w:docVars",
    "w:rsids", "m:mathPr", "w:attachedSchema", "w:themeFontLang", "w:clrSchemeMapping",
    "w:doNotIncludeSubdocsInStats", "w:doNotAutoCompressPictures", "w:forceUpgrade",
    "w:captions", "w:readModeInkLockDown", "w:smartTagType", "sl:schemaLibrary",
    "w:shapeDefaults", "w:doNotEmbedSmartTags", "w:decimalSymbol", "w:listSeparator",
)


def enable_field_update_on_open(document: DocumentObject) -> None:
    """让 Word 一打开文档就自动刷新域（主要是目录）。

    生成 docx 时，TOC 域里缓存的是我们写的占位提示语"（目录内容由 Word 自动
    生成，打开后右键选择「更新域」…）"。Word 不主动刷新域，于是用户打开文件
    看到的**就是这句提示语**，而不是目录条目——2026-09-12 用户因此以为"工具
    跑出来的目录根本没有引言结论"。在 settings.xml 里置
    <w:updateFields w:val="true"/> 之后，Word 打开时会自己把域刷一遍。
    """
    settings = document.settings.element
    update = settings.find(qn("w:updateFields"))
    if update is None:
        update = OxmlElement("w:updateFields")
        for tag in _AFTER_UPDATE_FIELDS:
            anchor = settings.find(qn(tag))
            if anchor is not None:
                anchor.addprevious(update)
                break
        else:
            settings.append(update)
    update.set(qn("w:val"), "true")


SESSION_TOKEN = "{届数}"


def expand_header_tokens(text: str, session: str = "") -> str:
    """页眉文字里的占位替换：`{届数}` → 用户填的届数，**每字间空一格**（"二十四"→
    "二 十 四"，跟页眉"每字间空一格"的写法一致；十星预设的页眉文字里写的是
    `第 {届数} 届`）。届数没填就把占位换成空。没有占位的页眉原样返回。"""
    if SESSION_TOKEN not in text:
        return text
    spaced = " ".join(ch for ch in session.strip() if not ch.isspace())
    return text.replace(SESSION_TOKEN, spaced)


def apply_header_footer(
    document: DocumentObject,
    rules: list[HeaderFooterRule],
    styles: StyleRegistry,
    header_bottom_line: bool = True,
    header_line_extend_cm: tuple[float, float] = (0.0, 0.0),
) -> None:
    header_style = None
    # 配对方向：**从最后一节往前配**（用户 2026-09-16 拍板"你看着办"）。
    # 预设「页眉页脚」那两行的顺序就是分节后的节顺序（末行＝正文部分），所以
    # 正文部分永远对着**最后一节**。以前是 `zip(sections, rules)` 按位置从前往后
    # 配：文档只剩一节（没有引言、也没识别出正文首标题）时，那唯一一节会套上
    # "前置部分"那一行——整篇变成罗马页码 + 前置页眉，而且没有任何提示。
    # 现在从后往前配，单节文档自然落到"正文部分"那一行上。
    pairs = list(zip(document.sections[::-1], rules[::-1]))[::-1]
    for section, rule in pairs:
        if rule.header_text:
            section.header.is_linked_to_previous = False
            p = section.header.paragraphs[0]
            header_style = styles["S27"]
            p.style = header_style
            p.text = rule.header_text

        if rule.page_number_enabled:
            section.footer.is_linked_to_previous = False
            p = section.footer.paragraphs[0]
            p.style = styles["S28"]
            # 页脚里原来可能已经有页码（用户自己插的域、文本框里的号，或上一次
            # 导出留下的）：先把**整份页脚**里的页码域摘干净再插我们自己的，
            # 否则成品上会出现两个页码（详见 `_drop_page_number_fields`）。
            _drop_page_number_fields(section.footer)
            _add_page_number_field(p, rule.page_number_prefix, rule.page_number_suffix)
            _set_page_number_format(section._sectPr, rule.page_number_format, start=1)
        else:
            # 明确"不启用页码"的那一节：**不能只"不写"**——页脚链接还连着前一节
            # 的话，这一节会把前一节的页码原样显示出来（比如正文部分顶着罗马
            # 数字）。用户 2026-09-16 裁决："断开页脚链接 + 摘掉页码域"，
            # 跟预设表里"未启用页码时不写入任何内容"那句对齐。
            # （没配到规则的那一节，例如自动封面，本来就不进这个循环、不碰。）
            section.footer.is_linked_to_previous = False
            _drop_page_number_fields(section.footer)

    if header_style is not None:
        # 页眉文字下面那条横线的总开关（预设「页眉线」小节）。封面节不放页眉，
        # 所以封面永远不受影响——本开关只管“有页眉的地方画不画线”。
        set_header_bottom_border(header_style, header_bottom_line)
        # 线向栏外伸出（十星）：只在画线时有意义；(0, 0)＝不碰，老行为。
        if header_bottom_line:
            set_header_line_extent(header_style, *header_line_extend_cm)


def make_toc_field(styles: StyleRegistry) -> OxmlElement:
    """**造**一个目录域元素（`w:fldSimple`），返回给调用方塞进段落里。

    与 `make_list_field()` 同一套做法：只**写下域代码**，生成时不需要打开 Word，
    用户打开文档时 Word 自己收集条目、自己算页码（settings 里开着 updateFields）。

    顺带把 toc 1/2/3 三个内置样式按预设建出来（预设 S15/S16/S17 映射到
    "TOC 1/2/3"）——规范 [202] 要求的“章标题行居左、一级节标题缩进 1 字符、
    二级节标题缩进 2 字符”就靠它们生效。目录**条目**的格式是 Word 用它自己的
    内置样式画的，把样式套到这一条上没用（生成目录时内容会被整个替掉）。

    ⚠️ 2026-09-17 从 `insert_toc_field_after()` 里拆出来（拆法与
    `apply_structure` 里缺件那一套一致）："造"归这里，"放"归调用方——
    索引法里位置由索引（`Op.out_index`）给，本函数看不到也不该看到位置。
    """
    for style_id in ("S15", "S16", "S17"):
        styles.ensure(style_id)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), _TOC_INSTR)
    run = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = _TOC_PLACEHOLDER
    run.append(t)
    fld.append(run)
    return fld


def insert_toc_field_after(label_paragraph: Paragraph, styles: StyleRegistry) -> Paragraph:
    """在 label_paragraph（"目录"标签段落，不论是识别出来的还是缺失后自动
    补上的）后面插一段 TOC 域，返回新插入的段落。目录内容从来不是用户手打
    的，每次都由工具生成，不走"缺了才补"那套逻辑——找到就在后面插，不用
    先判断"有没有内容"。

    （域本体由 `make_toc_field()` 造；这一层只负责"放"。）
    """
    new_p = OxmlElement("w:p")
    label_paragraph._p.addnext(new_p)
    paragraph = Paragraph(new_p, label_paragraph._parent)
    new_p.append(make_toc_field(styles))
    return paragraph


# "插图清单/附表清单"的域：按**样式名**收集题注（`\t "图题,1"` = 收"图题"
# 样式、条目定到目录级别 1），页码由 Word 自己算。2026-09-12 用真实 Word
# 实测过：自定义的中文样式名认，条目走 toc 1 样式、页码正确。
# 注意收集范围：**中英文对照时只有中文行挂着这两个样式**（英文行走直接格式，
# 见 detect.CaptionMatch.is_english），所以清单里一图一条、不会中英重复。
_LIST_FIELD_INSTRS = {
    "图": 'TOC \\h \\z \\t "图题,1"',
    "表": 'TOC \\h \\z \\t "表题,1"',
}
_LIST_PLACEHOLDER = "（清单内容由 Word 自动生成，打开后按 F9 刷新）"


def make_list_field(kind: str) -> OxmlElement:
    """造一个清单域元素（`w:fldSimple`），返回给调用方塞进段落里。

    跟目录域一样只是**写下域代码**——生成时不需要打开 Word，用户打开文档时
    Word 自己收集、自己算页码（settings 里开着 updateFields）。
    """
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), _LIST_FIELD_INSTRS[kind])
    run = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = _LIST_PLACEHOLDER
    run.append(t)
    fld.append(run)
    return fld
