"""脚注（`word/footnotes.xml`）。

python-docx **没有脚注 API**——脚注内容住在 `word/footnotes.xml` 这个**独立部件**
里，不在 `document.paragraphs` 里，包里也没有对应的部件类（`docx.oxml` 里只有
样式枚举常量）。所以只能自己解析 XML、改完再写回部件的 blob。

规范对脚注的两条要求（[214][215]）：
  * 脚注内容用宋体小五号、两端对齐、单倍行距、段前段后均 0；
  * 序号与正文之间空半个汉字符、序号字体是"正文"不是"上标"。

这里的做法**不往脚注段落上直接写格式**，而是：
  1. 把脚注段落上的**直接格式清掉**（pPr 只留一个指向「脚注文本」样式的
     pStyle，rPr 整个去掉）；
  2. 让那个样式去决定格式——它由预设的 S23 定义，已经映射到 Word 内置的
     `footnote text`。
这样格式只有一处真相，跟正文那套"先清直接格式、再套样式"的思路一致。

**规范 [215] 那句话管两件事（用户 2026-09-17 的原话）**："Front 意义上的字体
当然要写，要写正文的字体，一般是宋体；其次，shape 意义上的字体也要设，那就
是写成非上标或者上标。"
⇒ 预设「引注」里的 `footnote_marker_font` **不是字体名，是"要不要上标"的
开关**（界面上两个取值就是 `正文(非上标)` / `上标`）：
  * **字体**：一律写**正文的字体**（预设 S01 的 `font_east_asian` / `font_ascii`，
    中文一般宋体）；
  * **字形**：`正文(非上标)` = 不写 `w:vertAlign`（而且清 rPr 已经把 Word 默认
    那个上标/缩小去掉了）；`上标` = 写 `w:vertAlign val="superscript"`。

[215] 前半条"序号与内容之间空半个汉字符"：预设「引注」那一格填的是
**半角空格**，而宋体的半角空格宽度正好是半个汉字符，所以直接插 `U+0020`
就是对的，不需要用字间距去凑。Word 新建脚注时会自己插一个，这里只是兜底
补齐（幂等，已有就不动）。
"""

from __future__ import annotations

from lxml import etree

from docx.oxml.ns import qn

FOOTNOTES_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
)


def get_footnotes_part(document):
    """拿到 footnotes.xml 那个部件；文档里没有脚注（或没这个部件）时返回 None。"""
    try:
        return document.part.part_related_by(FOOTNOTES_REL)
    except KeyError:
        return None


def _load(document):
    part = get_footnotes_part(document)
    if part is None or not part.blob:
        return None, None
    return part, etree.fromstring(part.blob)


def count_real_footnotes(document) -> int:
    """数一下真正的脚注有几条。

    `footnotes.xml` 里除了正文用的脚注，还有几个 `w:type` 是 separator /
    continuationSeparator 的"装饰性"条目——它们不算脚注，要排掉。
    """
    _, root = _load(document)
    if root is None:
        return 0
    return sum(
        1 for note in root.findall(qn("w:footnote")) if note.get(qn("w:type")) is None
    )


# 预设「引注」那几格是**人写的说明**，不是枚举值，得翻译成 OOXML 的
# 取值。认不出来时**不猜**：按规范 [231] 的硬要求走（带圈阿拉伯数字），并打
# 一条警告——免得用户在 Excel 里改了却发现没生效。
_NUMBER_FORMATS = {
    "①②③": "decimalEnclosedCircle",
    "①②③……": "decimalEnclosedCircle",
    "①②③…": "decimalEnclosedCircle",
    "带圈阿拉伯数字": "decimalEnclosedCircle",
    "1.2.3.": "decimal",
    "123": "decimal",
    "阿拉伯数字": "decimal",
}
_DEFAULT_NUMBER_FORMAT = "decimalEnclosedCircle"

_TRAILING_SPACES = {
    "半角空格": " ",
    "半个汉字符": " ",  # 宋体里半角空格就是半个汉字符宽，两者等价
    "全角空格": "\u3000",
    "不空": "",
    # 界面上的写法是"不空格"（`main_window` 的 `fn_space` 下拉），而这里只认
    # "不空"/"无" ⇒ 用户选了"不空格"会被当成认不出来、静默退成半角空格
    # （2026-09-16 审计报告 Bug#2）。两种写法都认。
    "不空格": "",
    "无": "",
}

# CT_FtnProps 的子元素顺序（pos, numFmt, numStart, numRestart）。
#
# **位置坑（实测踩过）**：`w:footnotePr` 的合法位置**只有 `w:sectPr` 里面**
# （它在 schema 里属于 EG_SectPrContents）。`word/footnotes.xml` 的根下只允许
# `w:footnote`，把 footnotePr 塞在那里 Word **不报错也不生效**——文档能正常
# 打开，但 NumberingRule 读出来还是 0。旧的“手工塞 XML 位置不对”又添一例。
_FTN_PROPS_ORDER = ("w:pos", "w:numFmt", "w:numStart", "w:numRestart")


def _lookup(mapping: dict[str, str], raw: str, what: str, fallback: str) -> str:
    value = str(raw or "").strip()
    if value in mapping:
        return mapping[value]
    print(f"！预设的{what}写作「{value}」，不认识，按规范默认值处理（{fallback!r}）")
    return fallback


def _set_ftn_prop(props, tag: str, value: str) -> None:
    """在 w:footnotePr 里设置一项，没有就按 schema 顺序新建。"""
    element = props.find(qn(tag))
    if element is None:
        element = etree.Element(qn(tag))
        anchor = None
        for later in _FTN_PROPS_ORDER[_FTN_PROPS_ORDER.index(tag) + 1 :]:
            found = props.find(qn(later))
            if found is not None:
                anchor = found
                break
        if anchor is None:
            props.append(element)
        else:
            anchor.addprevious(element)
    element.set(qn("w:val"), value)


def _footnote_ref_run(paragraph):
    for run in paragraph.findall(qn("w:r")):
        if run.find(qn("w:footnoteRef")) is not None:
            return run
    return None


# 预设「引注」→ 脚注序号字体：界面上只有两个取值，它们是**开关**，不是字体名
# （规范 [215]：序号的字体是"正文"，不是"上标"）。详见模块 docstring。
_MARKER_SUPERSCRIPT = {
    "正文(非上标)": False,
    "正文": False,
    "非上标": False,
    "上标": True,
}

# 预设里没有 S01（正文）时的兜底字体。正常预设都有，走不到这里。
_DEFAULT_BODY_FONTS = ("宋体", "宋体")


def _set_marker_font(run, value: str,
                     body_fonts: tuple[str, str] = _DEFAULT_BODY_FONTS) -> None:
    """给"脚注序号"那一个 run 写**正文字体 + 字形**（预设「引注」那一格）。

    `body_fonts` ＝ （中文正文字体, 西文正文字体），由 pipeline 从预设 S01
    （正文）里取。

    * `正文(非上标)` ⇒ 写正文字体、**不写** `w:vertAlign`（＝非上标；上面清
      掉整段 rPr 时已经把 Word 默认那个上标去掉了）；
    * `上标` ⇒ 写正文字体 + `w:vertAlign val="superscript"`；
    * 认不出的取值（老预设/Excel 里填的真字体名，如"楷体"）⇒ 当**字体名**
      覆盖字体，字形按非上标（向后兼容，别让老预设变得没反应）。
    """
    if run is None or not value:
        return
    east_asian, ascii_font = body_fonts
    superscript = _MARKER_SUPERSCRIPT.get(value)
    if superscript is None:
        east_asian = ascii_font = value
        superscript = False
    rPr = run.find(qn("w:rPr"))
    if rPr is None:
        rPr = etree.Element(qn("w:rPr"))
        run.insert(0, rPr)          # rPr 必须是 run 的第一个子元素
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = etree.Element(qn("w:rFonts"))
        rPr.insert(0, rFonts)       # rFonts 排在 rStyle 之后、其余之前
    rFonts.set(qn("w:eastAsia"), east_asian)
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rFonts.set(qn(attr), ascii_font)
    if superscript:
        # rPr 是刚清过重建的：rFonts 已就位，vertAlign 接在它后面就合法。
        vert = rPr.find(qn("w:vertAlign"))
        if vert is None:
            vert = etree.SubElement(rPr, qn("w:vertAlign"))
        vert.set(qn("w:val"), "superscript")


_KNOWN_MARKER_SPACES = (" ", "　")   # 半角/全角空格——本函数认识、会管的两种


def _ensure_trailing_space(paragraph, space: str) -> None:
    """序号后面该有的空格状态：`space` 非空就确保有（幂等，已有就不动）；
    `space == ""`（预设选"不空格"）就把**我们自己认识的那一种**先摘掉——
    以前这里只会加、从不摘，选了"不空格"改了也白改，上一次导出（或 Word
    新建脚注时自带）插的空格会一直留着（2026-09-22 审计揪出的真 bug）。

    只认半角/全角空格这两种（`_KNOWN_MARKER_SPACES`）：序号后紧跟着的这个
    位置，按 [215] 的规矩本来就只该是"空格或没有"，不会是别的正文字符，
    摘的时候只摘开头这一个字符、不动这个 run 后面本来就有的其余文字（比如
    这个 run 恰好把空格和第一个字连在了一起）、也不删 run 元素本身
    （项目一贯做法：只改文字，不删元素）。
    """
    ref = _footnote_ref_run(paragraph)
    if ref is None:
        return
    later_runs = ref.itersiblings(qn("w:r"))
    for later in later_runs:
        text = later.find(qn("w:t"))
        if text is None:
            continue  # 不是文字 run（比如另一个域），继续往后找
        current = text.text or ""
        if space:
            if current.startswith(space):
                return        # 已经是想要的样子，幂等不动
            for known in _KNOWN_MARKER_SPACES:
                if current.startswith(known):
                    # 是我们认识的另一种空格（半角/全角来回切换时宽度不对）：
                    # 先摘掉，下面再插入想要的那种——以前不摘就直接往下插
                    # 新的一个，旧的那个还留着，序号后面就变成两个空格宽
                    # （2026-09-22 审计揪出的真 bug）。
                    text.text = current[len(known):]
                    break
            break     # 没有想要的这种空格：走到下面去补一个
        for known in _KNOWN_MARKER_SPACES:
            if current.startswith(known):
                text.text = current[len(known):]
                return
        return            # 序号后紧跟的不是空格（本来就没有）：不空格，什么都不用做
    if not space:
        return

    new_run = etree.Element(qn("w:r"))
    text = etree.SubElement(new_run, qn("w:t"))
    # 半角空格会被 XML 解析器当空白抹掉，必须显式声明保留
    text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text.text = space
    ref.addnext(new_run)


def normalize_footnotes(document, style_id: str, citation,
                       body_fonts: tuple[str, str] | None = None) -> int:
    """把一份文档里的脚注该做的都做掉——一次解析 footnotes.xml、一次写回。

    `style_id` 是「脚注文本」样式在**这份文档里**的 styleId（可能是
    `FootnoteText`，也可能是 Word 编的数字，所以必须由调用方传进来）。

    citation 是预设里的 `CitationConfig`（Excel「引注」表那几行），四项设置：

    | 预设字段 | 作用 | 规范 |
    |---|---|---|
    | `footnote_marker_style` | 序号用带圈阿拉伯数字 ①②③ | [231] |
    | `footnote_restart_each_page` | 每页重新编号 | [231] |
    | `footnote_marker_trailing_space` | 序号与内容之间空半格 | [215] |
    | `footnote_marker_font` | 序号的字体（正文字体）＋上标/非上标 | [215] |

    `body_fonts`：正文字体（中文, 西文），由 pipeline 从预设 S01 取；不传就用
    `_DEFAULT_BODY_FONTS`（宋体）兜底。序号写的就是它——[215] 说序号"字体是
    正文"，而不是写一个叫"正文(非上标)"的字体名（那是 2026-09-17 修的 bug）。

    返回改了几段；文档里没有真脚注时返回 0。
    """
    part, root = _load(document)
    if root is None:
        return 0

    number_format = _lookup(
        _NUMBER_FORMATS,
        citation.footnote_marker_style,
        "脚注序号样式",
        _DEFAULT_NUMBER_FORMAT,
    )
    space = _lookup(
        _TRAILING_SPACES,
        citation.footnote_marker_trailing_space,
        "序号后空格",
        " ",
    )
    marker_font = str(getattr(citation, "footnote_marker_font", "") or "").strip()
    fonts = tuple(body_fonts) if body_fonts else _DEFAULT_BODY_FONTS

    props = None
    for section in document.sections:
        # 分节级的属性就写在每一节的 sectPr 上（不是 footnotes.xml，见上面那条坑）
        sect_pr = section._sectPr  # noqa: SLF001 —— python-docx 只有这个入口
        props = sect_pr.find(qn("w:footnotePr"))
        if props is None:
            props = etree.Element(qn("w:footnotePr"))
            # EG_SectPrContents 的顺序是 headerReference*、footerReference*、
            # footnotePr?、endnotePr?、……——footnotePr 排在页眉页脚引用
            # 之后，不是最前面。以前无条件 insert(0, ...)：这一节本来就有
            # 页眉/页脚引用时，写完顺序就反了（Word 有可能因此不认这条
            # 脚注设置，2026-09-22 审计发现的）。这里改成插在最后一个
            # headerReference/footerReference 之后（没有就还是最前面，
            # 跟以前没有页眉页脚引用时的行为一致）。
            insert_pos = 0
            for i, child in enumerate(sect_pr):
                if child.tag in (qn("w:headerReference"), qn("w:footerReference")):
                    insert_pos = i + 1
            sect_pr.insert(insert_pos, props)
        _set_ftn_prop(props, "w:numFmt", number_format)
        if citation.footnote_restart_each_page:
            _set_ftn_prop(props, "w:numRestart", "eachPage")
        else:
            # ⚠️ 以前这里只在"勾选"时写 `w:numRestart`，取消勾选时从不摘掉——
            # 上一次导出（或源文件本来）写过的"每页重新编号"会一直留着，
            # 复选框看着是取消了，成品里其实还在按页重新编号
            # （2026-09-22 审计揪出的真 bug）。缺省（不写这个元素）在 OOXML
            # 里就是"连续编号、不按页重来"，直接删掉即可，不需要另写一个
            # "continuous"占位。
            existing = props.find(qn("w:numRestart"))
            if existing is not None:
                props.remove(existing)

    changed = 0
    for note in root.findall(qn("w:footnote")):
        if note.get(qn("w:type")) is not None:
            continue  # 分隔符那些，不是真脚注，不碰
        for paragraph in note.findall(qn("w:p")):
            old_pPr = paragraph.find(qn("w:pPr"))
            if old_pPr is not None:
                paragraph.remove(old_pPr)

            # pPr 必须是 w:p 的第一个子元素，所以 insert(0, ...) 而不是 append。
            new_pPr = etree.Element(qn("w:pPr"))
            pStyle = etree.SubElement(new_pPr, qn("w:pStyle"))
            pStyle.set(qn("w:val"), style_id)
            paragraph.insert(0, new_pPr)

            for run in paragraph.findall(qn("w:r")):
                rPr = run.find(qn("w:rPr"))
                if rPr is not None:
                    # 这一下同时完成两件事：格式交给样式管（[214]），
                    # 以及把脚注处的序号从"上标"变回正常大小（[215]）。
                    run.remove(rPr)

            # 序号字体与字形：清完 rPr 之后单独给"装着序号的那个 run"补上
            # 正文字体（中文/西文各一栏）＋上标开关（预设「引注」那一格）。
            _set_marker_font(_footnote_ref_run(paragraph), marker_font, fonts)

            # ⚠️ 以前这里有 `if space:` 守着，`space == ""`（预设选"不空格"）
            # 时函数压根不会被调用，摘除旧空格的逻辑写了也没用——改成一律调，
            # 让函数自己按 `space` 是不是空串决定"补"还是"摘"。
            _ensure_trailing_space(paragraph, space)
            changed += 1

    if changed:
        part._blob = etree.tostring(  # noqa: SLF001 —— 通用 Part 只能这样写回
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
    return changed
