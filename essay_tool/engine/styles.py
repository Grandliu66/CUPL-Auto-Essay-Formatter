"""把 ParagraphStyleSpec 写成 Word 真正的样式（Styles 面板里看得到、导航窗格/目录能用）。

"正文/一级标题/二级标题/三级标题/四级标题" 对应 Word 内置的
Normal/Heading 1-4（这样目录域、导航窗格、大纲级别才认得出来）；页眉/页脚/
脚注文本/目录一级~三级也各有 Word 内置名（见 BUILTIN_STYLE_NAMES）；
长引文、参考文献条目这类高频角色按 spec.name 新建自定义样式。
**名单之外的角色（摘要、关键词、各种标签……）不建样式**——走下面的
`_DIRECT_FORMAT_BASE`，把格式直接写在段落上。
"""

from __future__ import annotations

from docx.document import Document as DocumentObject
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, RGBColor
from docx.styles.style import BaseStyle
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from essay_tool.config.schema import ParagraphStyleSpec
from essay_tool.config.units import Indent, IndentKind

# 单倍/双倍/1.5 倍：比例固定，只设 `line_spacing_rule` 就够（python-docx 的
# setter 自己会把 w:line 写成对应值，见 docx.text.parfmt.line_spacing_rule）。
# 固定值/最小值（磅）/多倍行距（倍）都要另外算 `line_value`，走别的分支。
_FIXED_RATIO_LINE_RULES = frozenset({
    WD_LINE_SPACING.SINGLE, WD_LINE_SPACING.DOUBLE, WD_LINE_SPACING.ONE_POINT_FIVE,
})

# spec.name -> Word 内置样式名（英文 style id，跨语言 Word 都认得）。
#
# 能对上 Word 内置的就一定要对上：源文档里**本来就有**这些内置样式（比如学生
# 用官方模板写的论文，一定有 header/footer/footnote text），对上的话就是"把
# 格式覆盖到源文件已有的样式上"，而不是往文档里再加一个新的（用户 2026-09-12：
# "不要直接增加新预设"）。
#
# 页眉/页脚/脚注文本这三条还有一个必须对上的理由：**名字撞车**。预设里叫
# "脚注文本"的自定义样式会跟 Word 内置的「脚注文本」同名，Word 会把我们建的
# 那个自动改名成「脚注文本 1」，用户打开样式面板会看到一个莫名其妙的带" 1"
# 的样式。
# spec.name -> Word 内置样式的**内部名**（styles.xml 里 w:name 存的那个写法）。
#
# ！注意这里要填小写的内部名，不能填界面上那个大写名（"TOC 1"、"Heading 1"）。
# 因为 python-docx 查样式时会把名字过一道 BabelFish 别名表，而那张表**只有
# 12 条**（Caption/Footer/Header/Heading 1~9），根本没有 TOC 1、Footnote Text。
# 填大写名的话它会原样拿去当 w:name 找，而文档里存的是小写，结果找不到、
# 又新建一个同名（但界面显示一模一样）的样式——2026-09-12 就是这么把
# `TOC 1/2/3` 重复建了一遍。小写内部名过了 BabelFish 不会被改，正好命中。
BUILTIN_STYLE_NAMES: dict[str, str] = {
    "正文": "Normal",
    "一级标题": "heading 1",
    "二级标题": "heading 2",
    "三级标题": "heading 3",
    "四级标题": "heading 4",
    "页眉": "header",
    "页脚页码": "footer",
    "脚注文本": "footnote text",
    # 目录条目（引言/一、/（一）/1. 那些行）：它们**不是用户写的**，是 TOC 域
    # 生成的，Word 用它自己的内置样式 toc 1/2/3 画。预设里 S16/S17 的缩进值
    # （一级节标题缩进1字符、二级缩进2字符，规范 [202]）本来就是为它们准备的，
    # 只是从来没接到内置样式上——接上之后缩进自动生效，不需要给任何段落套样式。
    "目录一级": "toc 1",
    "目录二级": "toc 2",
    "目录三级": "toc 3",
}

# Heading 1-4 -> 大纲级别（0 起）。真正在 Word 里点过"标题 1"这些样式的文档，
# 这个值早就在样式定义里了；但源文件如果从没用过 Word 内置的标题样式（比如
# 全篇都是直接格式敲出来的、从没点过样式面板），`styles["Heading 1"]` 会找
# 不到，_get_or_add_style() 就会新建一个"看起来叫这个名字"但没有大纲级别的
# 空样式——名字对了，但目录域/导航窗格根本不认，TOC 域会更新出"未找到目录
# 项"。所以这里不管样式是原来就有的还是刚建的，都把大纲级别显式补一遍。
_HEADING_OUTLINE_LEVELS: dict[str, int] = {
    "heading 1": 0,
    "heading 2": 1,
    "heading 3": 2,
    "heading 4": 3,
    # 「附录标题」不是 Word 内置样式，是照着规范 [225] 现建的（原因见
    # `_DIRECT_FORMAT_BASE` 上面那段）。它不是 Heading 1，得单独把大纲级别
    # 补上，否则目录域收不到它——而规范 [90] 明确要求附录要进目录。
    "附录标题": 0,
}


def _set_fonts(rPr, ascii_font: str, east_asian_font: str) -> None:
    # 用 get_or_add_rFonts() 按 schema 顺序插入，别手工 insert(0, …)：rPr 里
    # w:rFonts 必须排在 w:rStyle **之后**（run 上挂着字符样式时 rStyle 在最前），
    # 插到 0 号位会得到一个手调 Word 绝不会产生、Word 也不保证认的顺序。
    rFonts = rPr.get_or_add_rFonts()
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rFonts.get(qn(attr)) is not None:
            del rFonts.attrib[qn(attr)]
    rFonts.set(qn("w:ascii"), ascii_font)
    rFonts.set(qn("w:hAnsi"), ascii_font)
    rFonts.set(qn("w:eastAsia"), east_asian_font)
    rFonts.set(qn("w:cs"), ascii_font)


def _apply_indent(pPr, first_line: Indent, left: Indent) -> None:
    """把预设里的"首行缩进/左缩进"写进 pPr。

    ⚠️ 预设写"无"的时候**必须显式写 0**，不能什么都不写（写成"无"就等于让
    段落去继承样式的缩进）。理由：预设里绝大多数样式（一级标题、题目、各种
    标签、图题表题、页眉页脚……）的底座就是"正文"，而正文带 2 字符首行缩进；
    不写 0 就会继承那 2 字符，而 Word 的「居中」是在**扣掉首行缩进之后的区域**
    里居中 ⇒ 整行看起来往右偏一个字（用户 2026-09-16 实拍民总：正文里居中的
    一级标题"没有真正居中，而是基于首行缩进居中的"，"不光是正文，其他标签也
    差不多是这样"）。写全 8 个属性是为了同时盖掉继承来的 `firstLineChars`
    与 `leftChars`——只写 `firstLine=0` 时，字符数版本（chars）在 Word 里优先。
    跟 `apply_structure._zero_paragraph_indent`（图片段落）是同一个口径。
    """
    # 用 get_or_add_ind() 而不是自己 append：pPr 的子元素有 schema 顺序
    # （ind 必须排在 jc 前面），位置错了 Word 可能报"内容有问题"。样式和段落
    # 两条路都走这里，所以必须按规矩插。
    ind = pPr.get_or_add_ind()
    for attr in ("w:firstLine", "w:firstLineChars", "w:hanging", "w:hangingChars",
                 "w:left", "w:leftChars"):
        if ind.get(qn(attr)) is not None:
            del ind.attrib[qn(attr)]

    if first_line.kind == IndentKind.CHARS:
        ind.set(qn("w:firstLineChars"), str(int(round(first_line.chars * 100))))
    elif first_line.kind == IndentKind.LENGTH:
        ind.set(qn("w:firstLine"), str(first_line.length.twips))
    else:
        ind.set(qn("w:firstLine"), "0")
        ind.set(qn("w:firstLineChars"), "0")

    if left.kind == IndentKind.CHARS:
        ind.set(qn("w:leftChars"), str(int(round(left.chars * 100))))
    elif left.kind == IndentKind.LENGTH:
        ind.set(qn("w:left"), str(left.length.twips))
    else:
        ind.set(qn("w:left"), "0")
        ind.set(qn("w:leftChars"), "0")

    if not ind.attrib:
        pPr.remove(ind)


# 上面那批内置样式“常见”的 styleId。**这只是个保险**：Word 在很多文档里
# 并不用 "Heading1" 这种名字，而是给 styleId 编数字（实测 `民总.docx` 里
# heading 1 的 styleId 就是 "2"、toc 1 是 "9"）——那种情况下这里根本找不到，
# 会自动退回按名字查找，所以写错/写漏都无害，真正起作用的始终是名字那一路。
# （注：样式名被 `_RENAME_TO_SPEC_NAME` 改掉后同一次运行内也不会再按名字
# 查第二遍——StyleRegistry 会缓存已建好的样式。）
_BUILTIN_STYLE_IDS: dict[str, str] = {
    "Normal": "Normal",
    "heading 1": "Heading1",
    "heading 2": "Heading2",
    "heading 3": "Heading3",
    "heading 4": "Heading4",
    "header": "Header",
    "footer": "Footer",
    "footnote text": "FootnoteText",
}

# 要把界面名改成预设里的名字的角色（Word 内置的本地化名字跟预设表对不上：
# 中文 Word 把 Heading 1 显示成"标题 1"，而预设里叫"一级标题"；Footer 显示成
# "页脚"，预设里叫"页脚页码"）。没列出来的（正文/页眉/脚注文本）Word 本来就
# 显示成那个中文名，不用动，免得节外生枝。
#
# **TOC 1/2/3 也故意不列进来**：Word 生成目录时要按这个名字去认那两个样式，
# 改了名有可能让它找不到、退回默认值。而标题能改是因为目录靠**大纲级别**
# 收集条目，跟样式名无关。
_RENAME_TO_SPEC_NAME = {"heading 1", "heading 2", "heading 3", "heading 4", "footer"}


def _find_style_by_id(document: DocumentObject, style_id: str) -> BaseStyle | None:
    for style in document.styles:
        if style.element.get(qn("w:styleId")) == style_id:
            return style
    return None


def set_header_bottom_border(style: BaseStyle, enabled: bool = True) -> None:
    """给“页眉”样式加/摘那条下边框线——「页眉线」开关（预设「页眉线」小节）。

    官方模板的“页眉”样式自带 <w:pBdr><w:bottom w:val="single" w:color="auto"
    w:sz="6" w:space="1"/></w:pBdr>，所以每页页眉文字下面都有一道横线；
    封面没有页眉，自然也没有线（2026-09-12 做封面时才发现我们一直漏了这个
    属性——生成稿的页眉光秃秃的）。

    enabled=False 时把已有的下边框摘掉——**源文档自带的线也算**，因为用户
    填“不要”的意思就是不要。挂在**样式**上而不是段落上：用该样式的页眉
    段落自动带线。调用方：sections.apply_header_footer（拿到预设开关值）。
    """
    pPr = style.element.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if not enabled:
        if pBdr is not None:
            bottom = pBdr.find(qn("w:bottom"))
            if bottom is not None:
                pBdr.remove(bottom)
            if len(pBdr) == 0:
                pPr.remove(pBdr)
        return
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        # pBdr 在 pPr 子元素序列里排得比较靠前（shd/tabs 之前），不能无脑 append
        for tag in (
            "w:shd", "w:tabs", "w:spacing", "w:ind", "w:jc",
            "w:textAlignment", "w:outlineLvl", "w:rPr", "w:sectPr",
        ):
            anchor = pPr.find(qn(tag))
            if anchor is not None:
                anchor.addprevious(pBdr)
                break
        else:
            pPr.append(pBdr)
    bottom = pBdr.find(qn("w:bottom"))
    if bottom is None:
        bottom = OxmlElement("w:bottom")
        pBdr.append(bottom)
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:color"), "auto")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")


def set_header_line_extent(style: BaseStyle, left_cm: float, right_cm: float) -> None:
    """让“页眉”样式下面那条线向正文栏外伸出（左、右各 N 厘米）。

    段落边框画在段落的缩进范围内，所以给页眉样式设**负**的左右缩进，线就伸出
    正文栏（学术十星的页眉线比正文栏宽，样例实测左伸 1.5cm、右伸 1.35cm，取对称的 1.5/1.5）。页眉
    文字是居中的，会跟着缩进范围一起挪一点点（左右伸出量之差的一半），十星
    样张里也是这样。

    两个量都为 0 时**什么都不碰**——保持样式原有的缩进（老预设的行为）。
    `w:ind` 上如果带着按字符计的 `leftChars`/`rightChars`，Word 会优先认它们、
    忽略我们写的厘米值，所以一并清掉。
    """
    if not left_cm and not right_cm:
        return
    fmt = style.paragraph_format
    fmt.left_indent = Cm(-left_cm)
    fmt.right_indent = Cm(-right_cm)
    ind = style.element.get_or_add_pPr().find(qn("w:ind"))
    if ind is not None:
        for attr in ("leftChars", "rightChars", "startChars", "endChars"):
            ind.attrib.pop(qn(f"w:{attr}"), None)


def _get_or_add_style(document: DocumentObject, spec: ParagraphStyleSpec) -> BaseStyle:
    builtin_name = BUILTIN_STYLE_NAMES.get(spec.name)
    target_name = builtin_name or spec.name
    styles = document.styles

    style = None
    builtin_id = _BUILTIN_STYLE_IDS.get(target_name)
    if builtin_id is not None:
        style = _find_style_by_id(document, builtin_id)
    if style is None:
        try:
            style = styles[target_name]
        except KeyError:
            style = styles.add_style(target_name, WD_STYLE_TYPE.PARAGRAPH)

    if builtin_name is not None and target_name in _RENAME_TO_SPEC_NAME:
        style.name = spec.name

    outline_level = _HEADING_OUTLINE_LEVELS.get(target_name)
    if outline_level is not None:
        pPr = style.element.get_or_add_pPr()
        outlineLvl = pPr.find(qn("w:outlineLvl"))
        if outlineLvl is None:
            outlineLvl = OxmlElement("w:outlineLvl")
            # outlineLvl 在 pPr 子元素序列里靠后，但必须排在 w:rPr / w:sectPr /
            # w:pPrChange 前面，不能无脑 append。
            for tag in ("w:rPr", "w:sectPr", "w:pPrChange"):
                anchor = pPr.find(qn(tag))
                if anchor is not None:
                    anchor.addprevious(outlineLvl)
                    break
            else:
                pPr.append(outlineLvl)
        outlineLvl.set(qn("w:val"), str(outline_level))

    return style


def _pin_font_color(style: BaseStyle) -> None:
    """把样式字色钉死为黑色。

    规范不配色（预设里也没有"颜色"这一项），但**复用源文档已有的 Word 内置
    样式**（heading 1～4 这些）时，会把源样式自带的颜色一起带出来——python-docx
    自带模板的标题样式是蓝色的，凡是"平地起"的文档（如空白模板）都会漏蓝；
    用户 2026-09-13 在生成的空白模板里抓到。这里显式钉黑，顺手删掉主题色
    残留（themeColor/themeTint/themeShade）。
    """
    style.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    rPr = style.element.get_or_add_rPr()
    color = rPr.find(qn("w:color"))
    if color is not None:
        for attr in ("w:themeColor", "w:themeTint", "w:themeShade"):
            if color.get(qn(attr)) is not None:
                del color.attrib[qn(attr)]


def _toc_tab_position(document) -> int | None:
    """版心宽度（twips）——目录条目那页码的右对齐制表位该放的位置。

    默认取**最后一节**（正文那一节）：封面节是后来拼上去的，页边距另算。
    样式建好之后会有一次 `refresh_toc_tab_stops()`，按目录**实际所在的那一节**
    再校正一遍——前言用 3.17 厘米、正文用 2 厘米的稿子（官方模板就是）两节
    版心不一样宽，取错了会让页码戳到右边界外面。
    """
    for section in reversed(list(document.sections)):
        width = section.page_width
        left = section.left_margin
        right = section.right_margin
        if width is not None and left is not None and right is not None:
            return width.twips - left.twips - right.twips
    return None


def set_toc_entry_tabs(document, style, spec: ParagraphStyleSpec) -> None:
    """给目录条目样式补上"页码贴右边界 + 虚线引导点"的制表位。

    为什么必须补（用户 2026-09-16：“那个目录还是没改好”）：Word 生成的目录
    条目末尾是一个制表符，页码靠**样式上的制表位**拉到右边界；而官方模板里
    **只有 `toc 1` 有**（`right + dot + pos=8296`），`toc 2`/`toc 3` 一个都没有
    ⇒ 二级三级条目的页码直接跟在文字后面（一级有虚线、二三级没有，很难看）。

    位置就是**本文档的版心宽度**：模板那个 8296 本来就是模板自己的版心宽
    （左右 3.17 厘米），我们预设是左右 2 厘米、版心 17 厘米＝9638 twips。

    ⚠️ **三级用同一个 pos，不能再加上各级自己的左缩进。** 2026-09-16 走过
    一次弯路：当时真 Word 的 COM 量法（`Information(5)` 取末字符）看着像是
    “落点＝左页边距＋pos−左缩进”，于是把左缩进加回了 pos；结果用户拿 PNG
    一眼看出“左侧缩进越多的，右边突出也越多”。改用**像素**重新量那张图
    （tools/_probe_toc_pixels.py，2.4 倍图、右边界应在 1292px）得到：一级
    1289~1291（对）、二级 1323、三级 1357——两级分别往右突出了约 13/27 磅，
    正好就是加回去的那点缩进。结论：`pos` 是**相对左页边距的绝对位置**。
    """
    width = _toc_tab_position(document)
    if width is None:
        return
    pos = width
    pPr = style.element.get_or_add_pPr()
    tabs = pPr.find(qn("w:tabs"))
    if tabs is None:
        tabs = OxmlElement("w:tabs")
        # schema 顺序：tabs 排在下 shd 之后、spacing 之前
        for tag in ("w:spacing", "w:ind", "w:jc"):
            anchor = pPr.find(qn(tag))
            if anchor is not None:
                anchor.addprevious(tabs)
                break
        else:
            pPr.append(tabs)
    for existing in list(tabs):
        tabs.remove(existing)
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:leader"), "dot")
    tab.set(qn("w:pos"), str(pos))
    tabs.append(tab)


# 需要补制表位的样式（＝目录一/二/三级 → 内置 toc 1/2/3）
_TOC_STYLE_IDS = frozenset({"S15", "S16", "S17"})

# 目录一/二/三级对应的内置样式名（跟 BUILTIN_STYLE_NAMES 同口径）
_TOC_STYLE_NAMES = ("toc 1", "toc 2", "toc 3")


def refresh_toc_tab_stops(document, anchor) -> None:
    """把目录条目样式上的制表位改成按**目录实际所在那一节**的版心宽来写。

    为什么还要来这么一趟（`apply_style_spec` 里不是算过吗）：样式是文档级的，
    而“版心”是**节的**属性。建样式那会儿节的划分还没定（分节在管线后段做），
    所以那时只能按最后一节（正文）估。前言用 3.17 厘米、正文用 2 厘米的稿子
    （官方模板就是），两节版心差着两千多 twips：拿正文那节的宽度写 pos，
    落到前言那一节里就是“页码戳到右边界外面”。

    三个层级用**同一个** pos，不加各级自己的左缩进——理由见
    `set_toc_entry_tabs` 的注释（加回去会让二级三级往右突出）。

    `anchor` 是目录域（或第一条 `toc N` 条目）的段落。
    """
    if anchor is None:
        return
    body = document.element.body
    section_index = 0
    for child in body:
        if child is anchor._p:
            break
        if child.tag == qn("w:p"):
            pPr = child.find(qn("w:pPr"))
            if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
                section_index += 1
        elif child.tag == qn("w:sectPr"):
            section_index += 1
    sections = list(document.sections)
    if section_index >= len(sections):
        return
    section = sections[section_index]
    if None in (section.page_width, section.left_margin, section.right_margin):
        return
    pos = (section.page_width.twips - section.left_margin.twips
           - section.right_margin.twips)
    for name in _TOC_STYLE_NAMES:
        try:
            style = document.styles[name]
        except KeyError:
            continue
        pPr = style.element.find(qn("w:pPr"))
        tabs = pPr.find(qn("w:tabs")) if pPr is not None else None
        if tabs is None:
            continue                    # 没给我们建过制表位的样式不动它
        for tab in tabs:
            tab.set(qn("w:pos"), str(pos))


def apply_style_spec(document: DocumentObject, spec: ParagraphStyleSpec) -> BaseStyle:
    style = _get_or_add_style(document, spec)

    style.font.size = spec.size
    style.font.bold = spec.bold
    # 斜体同样要"压平"：内置标题样式自带斜体（python-docx 的"标题 4"），
    # 不显式写 False 的话会漏进成品（用户 2026-09-13 在四级标题上抓到）。
    style.font.italic = False
    _set_fonts(style.element.get_or_add_rPr(), spec.font_ascii, spec.font_east_asian)
    _pin_font_color(style)

    pf = style.paragraph_format
    pf.alignment = spec.alignment
    pf.space_before = spec.space_before
    pf.space_after = spec.space_after
    pf.page_break_before = spec.page_break_before

    if spec.line_rule in _FIXED_RATIO_LINE_RULES:
        # 单倍/双倍/1.5 倍：比例固定，只设 rule——python-docx 的 setter 会自己把
        # 对应的 w:line 一并写好（见 docx.text.parfmt），不经过 `line_value`
        # （2026-09-22 修：这三种以前压根没在预设里收，见 `units.LINE_RULES`）。
        pf.line_spacing_rule = spec.line_rule
    elif spec.line_rule == WD_LINE_SPACING.MULTIPLE:
        pf.line_spacing = spec.line_value
    else:
        pf.line_spacing = spec.line_value
        pf.line_spacing_rule = spec.line_rule

    _apply_indent(style.element.get_or_add_pPr(), spec.first_line_indent, spec.left_indent)
    if spec.id in _TOC_STYLE_IDS:
        # 目录条目：页码靠样式上的右对齐制表位拉到右边界（见那个函数的注释）
        set_toc_entry_tabs(document, style, spec)
    return style


# 不建独立样式、改用"基础样式 + 段落直接格式"落地的角色 -> 拿哪个样式当底座。
#
# 判断依据是"这个格式会被多少个段落用"：只用一两次的（各种标签、题目、附属
# 段落），要是每种都在 Word 里建一个样式，用户打开文件会看到一堆"只出现过
# 一次"的样式名，很不自然（用户 2026-09-12 的反馈："谁会为一个只用一次的
# 东西设预设呢"）。这一类改成把格式**直接写在段落上**，样式面板里只留真正
# 会被反复使用的"角色样式"。
#
# **附录标题（S19）故意不在这张表里**，它要建一个真样式「附录标题」。
# 规范 [225] 要求它"小四号黑体字加粗……格式设置为一级标题"，也就是说它是
# 唯一一个"挂着标题身份、却要换一套字体"的角色。以前的做法是"套 S02 底座
# + 把黑体直接写在段落上"，结果**目录里那一条也变成黑体**（用户实测）——
# Word 生成目录条目时会把标题段落的**直接字符格式**一起复制进去。改成一个
# 独立样式之后字体落在样式层、目录条目不再继承它，同时靠
# `_HEADING_OUTLINE_LEVELS` 里补的大纲级别照样进目录。
# 注意：字体写进样式、而段落套样式，所以**不能**再把 S19 放回这张表里。
_DIRECT_FORMAT_BASE: dict[str, str] = {
    "S06": "S01",  # 中文题目
    "S07": "S01",  # 摘要标签
    "S08": "S01",  # 摘要正文
    "S09": "S01",  # 关键词标签
    "S10": "S01",  # 关键词内容
    "S11": "S01",  # 英文题目
    "S12": "S01",  # 英文摘要标签
    "S13": "S01",  # 英文摘要正文
    "S14": "S01",  # 目录标题
    "S20": "S01",  # 附录序号题名
    "S22": "S01",  # 致谢正文
    "S29": "S01",  # 副标题
    "S30": "S01",  # 英文关键词标签
    "S31": "S01",  # 英文关键词内容
    # 英文副标题（用户 2026-09-15）：跟中文副标题一样"只用一两次"，不建独立
    # 样式；字体要写 Times New Roman，所以不能复用 S29（S29 的西文字体是黑体）。
    "S32": "S01",
}

# 不在这张表里、也不在上面内置名单（BUILTIN_STYLE_NAMES）里的角色，真被
# 用到时按普通独立样式建（StyleRegistry 的兜底行为）。**预设表里一行都没删**
# ——"能不能建"这项能力不受影响，见用户 2026-09-12 的提醒。

# 对外只读别名：GUI（主窗口「格式」页签）拿它给"直接格式"的行做标记，
# 免得界面上再抄一份名单。
DIRECT_FORMAT_IDS = frozenset(_DIRECT_FORMAT_BASE)
# 同上：假样的父样式——界面排表时把假样式排到它后面。
DIRECT_FORMAT_BASE = dict(_DIRECT_FORMAT_BASE)


def write_spec_onto_paragraph(paragraph: Paragraph, spec: ParagraphStyleSpec) -> None:
    """把一套格式**直接写到段落上**，不经过 Word 样式。

    段落级格式（对齐、缩进、间距、段前分页）写进 pPr；字体字号加粗必须写进
    **每一个 run**——Word 里给整段改字体就是这个效果，只写段落标记的话文字
    本身不会变。
    """
    pf = paragraph.paragraph_format
    pf.alignment = spec.alignment
    pf.space_before = spec.space_before
    pf.space_after = spec.space_after
    pf.page_break_before = spec.page_break_before
    if spec.line_rule in _FIXED_RATIO_LINE_RULES:
        pf.line_spacing_rule = spec.line_rule
    else:
        pf.line_spacing = spec.line_value
        pf.line_spacing_rule = spec.line_rule

    pPr = paragraph._p.get_or_add_pPr()
    _apply_indent(pPr, spec.first_line_indent, spec.left_indent)

    # 段落标记（pPr/rPr）这里**故意不设**：它只在段落为空时才看得出效果
    # （比如装 TOC 域那一段——域内容不算 run，paragraph.runs 是空的），而那一段
    # 的可见文字由 Word 的 TOC 样式决定，设了也看不出差别。反过来，手工往
    # pPr/rPr 里塞 w:sz/w:b 还得自己维持 schema 顺序，塞错位置 Word 会报
    # "内容有问题"。段落不空时，字体已经写在下面每个 run 上了。
    #
    # 遍历**后代** run 而不是 paragraph.runs：英文行里用户自己插的 Word 题注
    # 编号在 `w:fldSimple` 里，那个结果 run 不是段落的直接子元素——不一起上
    # 字体的话，数字会按段落本来的样式渲染成另一个字号（中文题注行不受影响，
    # 它走真样式）。
    for run_el in paragraph._p.iter(qn("w:r")):
        run = Run(run_el, paragraph)
        run.font.size = spec.size
        run.font.bold = spec.bold
        _set_fonts(run._element.get_or_add_rPr(), spec.font_ascii, spec.font_east_asian)


def apply_spec(paragraph: Paragraph, registry: "StyleRegistry", style_id: str) -> None:
    """把一个角色（Sxx）落到段落上——**全项目唯一决定"建不建样式"的地方**。

    该角色在 `_DIRECT_FORMAT_BASE` 里：只套基础样式，格式直接写在段落上；
    否则：按需建出独立样式，段落套上去。
    """
    base_id = _DIRECT_FORMAT_BASE.get(style_id)
    if base_id is None:
        paragraph.style = registry[style_id]
        return
    paragraph.style = registry[base_id]
    write_spec_onto_paragraph(paragraph, registry.spec(style_id))


class StyleRegistry:
    """按需建样式：**只有真正被用到的样式才会建进文档**。

    以前是一次性把预设表里三十多个样式全建出来，结果成品文档的样式面板里堆着
    一堆从头到尾没用上的样式（用户 2026-09-12 提的问题）。现在改成：谁在管线
    上被 `styles["S07"]` 这样取用一次，才现建一次；没被取用的永远不建，也就
    不会出现在样式面板里。

    对调用方来说用法跟 dict 一样（`styles["Sxx"]`），区别只是"取值"这个动作
    本身会触发创建。想知道这次到底用了哪些样式，看 `created_ids`。
    """

    def __init__(self, document: DocumentObject, specs: dict[str, ParagraphStyleSpec]) -> None:
        self._document = document
        self._specs = specs
        self._built: dict[str, BaseStyle] = {}

    def __getitem__(self, style_id: str) -> BaseStyle:
        style = self._built.get(style_id)
        if style is None:
            spec = self._specs.get(style_id)
            if spec is None:
                raise KeyError(f"预设里没有编号 {style_id!r} 的样式")
            style = apply_style_spec(self._document, spec)
            self._built[style_id] = style
        return style

    def __contains__(self, style_id: object) -> bool:
        return style_id in self._specs

    def ensure(self, style_id: str) -> None:
        """把某个样式按预设建出来，但**不套给任何段落**。

        目录一/二/三级就是这种情况：它们是 Word 生成目录时自己拿去用的
        （内置 toc 1/2/3），不需要、也没法套到某个具体段落上。
        """
        _ = self[style_id]

    def spec(self, style_id: str) -> ParagraphStyleSpec:
        """取某个角色的完整格式定义——"直接写到段落上"那条路要用它。"""
        spec = self._specs.get(style_id)
        if spec is None:
            raise KeyError(f"预设里没有编号 {style_id!r} 的样式")
        return spec

    @property
    def created_ids(self) -> list[str]:
        """已经真正建出来的样式编号，按创建顺序。"""
        return list(self._built)
