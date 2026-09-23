r"""拆包读取器：把一份 docx 读成"逐段记录"。

为什么不用 python-docx 的属性：规范管得最细的那些（`leftChars`、`w:tabs/@pos`、
`w:numPr`、`outlineLvl`、`vertAlign`）python-docx 大多读不到，只能自己拆包看 XML。

这是差分比对与期望对比**共用的唯一"实际值"来源**。两条使用路径：
  * 两个 docx 互相差（不需要期望值）
  * docx vs 期望值（期望值由规则表/预设算出）

口径说明（很重要，踩过坑）：
  * 段号 = `w:body` 的**直接子元素**里数出来的 `w:p` 下标（表格算一个块，不算段）。
    这与管线"拆完内容控件之后"的段落下标一致；若产物里还残留 `w:sdt`，会在
    `warnings` 里报出来（那种情况下段号不可比）。
  * 表格单元格里的段落**不在**这个序列里，单独按 (表序, 行, 列) 收集。

只读，不改任何东西。
"""

from __future__ import annotations

import hashlib
import pathlib
import zipfile
from dataclasses import dataclass, field

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def qn(tag: str) -> str:
    """`qn("w:jc")` → `{…}jc`；也接受已经带命名空间的写法。"""
    if tag.startswith("{"):
        return tag
    prefix, _, local = tag.partition(":")
    return f"{{{NS[prefix]}}}{local}" if prefix in NS else local


def local_name(element) -> str:
    """元素的本地名（不带命名空间），如 `{…}jc` → `jc`。"""
    tag = element.tag
    return tag.split("}", 1)[1] if isinstance(tag, str) and "}" in tag else str(tag)


def attrs_of(element) -> dict[str, str]:
    """元素的属性 → `{本地名: 值}`（`w:val` 直接叫 `val`）。"""
    return {
        (key.split("}", 1)[1] if "}" in key else key): value
        for key, value in element.attrib.items()
    }


def signature(element) -> str:
    """整棵子树的紧凑签名（属性 + 递归子元素）。

    为什么需要：很多东西**不是靠属性表达的**，而是靠子元素——页眉那条装饰线
    是 `w:pPr/w:pBdr/w:bottom`，制表位是 `w:tabs/w:tab`。只看属性的话，
    「加/去一条页眉线」在比较里完全看不出来（踩过：页眉装饰线开关报"无变化"）。
    """
    if element is None:
        return ""
    name = local_name(element)
    attrs = ";".join(f"{k}={v}" for k, v in sorted(attrs_of(element).items()))
    children = "".join(signature(child) for child in element
                       if isinstance(child.tag, str))
    return f"{name}({attrs})" + (f"<{children}>" if children else "")


def child_attrs(element, only_w: bool = True) -> dict[str, dict[str, str]]:
    """元素的直接子元素 → `{本地名: 属性表 + #sig}`。同名子元素只留最后一个。

    `#sig` 是整棵子树的签名（见 `signature()`），比较时必须一起看。
    """
    out: dict[str, dict[str, str]] = {}
    if element is None:
        return out
    for child in element:
        if not isinstance(child.tag, str):
            continue
        if only_w and not child.tag.startswith("{" + W + "}"):
            continue
        cell = attrs_of(child)
        cell["#sig"] = signature(child)
        out[local_name(child)] = cell
    return out


# ------------------------------------------------------------------ 记录


@dataclass
class Para:
    """一个段落在产物里的"实际值"。"""

    index: int | None                     # body 里的段号（表格内段落为 None）
    text: str
    style_id: str = ""
    style_name: str = ""
    ppr: dict[str, dict[str, str]] = field(default_factory=dict)        # pPr 子元素原始属性
    marker_rpr: dict[str, dict[str, str]] = field(default_factory=dict)  # 段落标记的 rPr
    runs: list[tuple[str, dict[str, dict[str, str]]]] = field(default_factory=list)
    fields: list[str] = field(default_factory=list)                      # 域指令原文
    has_drawing: bool = False
    has_pict: bool = False
    in_table: bool = False
    table_path: tuple[int, int, int] | None = None
    # Word 自动编号的引用：编号不写在文字里，要根据这两个值反算可见编号
    num_id: str | None = None
    num_level: int = 0

    # ---- 便捷读取（差分报告里只打印用得上的那几项）----
    def prop(self, name: str) -> str:
        """某个 pPr 子项的显示文本（`jc` / `ind` / `spacing` / `numPr` …）。"""
        got = self.ppr.get(name)
        if got is None:
            return ""
        return ";".join(f"{k}={v}" for k, v in sorted(got.items())
                        if not k.startswith("#"))

    def label(self) -> str:
        if self.index is not None:
            return f"段{self.index}"
        if self.table_path:
            t, r, c = self.table_path
            return f"表{t}-{r}-{c}"
        return "?"


@dataclass
class DocShape:
    """一份产物的整体形状。"""

    path: pathlib.Path
    parts: dict[str, str] = field(default_factory=dict)      # 部件名 → md5
    paras: list[Para] = field(default_factory=list)
    header_paras: list[Para] = field(default_factory=list)
    footer_paras: list[Para] = field(default_factory=list)
    footnote_paras: list[Para] = field(default_factory=list)
    tables: list[list[list[Para]]] = field(default_factory=list)   # [表][行][列]
    table_repeat_header: list[bool] = field(default_factory=list)
    sections: list[dict[str, str]] = field(default_factory=list)
    styles: dict[str, dict] = field(default_factory=dict)     # styleId → {name, basedOn, pPr, rPr}
    doc_defaults: dict[str, dict[str, dict[str, str]]] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)


# ------------------------------------------------------------------ 读取


def read_package(path: pathlib.Path) -> tuple[dict[str, bytes], dict[str, str]]:
    with zipfile.ZipFile(path) as package:
        raw = {name: package.read(name) for name in package.namelist()}
    digests = {name: hashlib.md5(data).hexdigest()[:12] for name, data in raw.items()}
    return raw, digests


def read_paragraph(element, *, index: int | None, in_table: bool,
                   table_path=None) -> Para:
    """把一个 `w:p` 元素读成 Para。"""
    para = Para(index=index, text="", in_table=in_table, table_path=table_path)
    ppr_element = element.find(qn("w:pPr"))
    if ppr_element is not None:
        para.ppr = child_attrs(ppr_element)
        # 段里挂的分节属性（分节符）**剔掉**：同一个东西在「节」的比较里已经有
        # 一份，留着会让改页边距看起来像"某一段正文变了"（看着像误伤）。
        para.ppr.pop("sectPr", None)
        marker = ppr_element.find(qn("w:rPr"))
        para.marker_rpr = child_attrs(marker)
        numpr = ppr_element.find(qn("w:numPr"))
        if numpr is not None:
            numid = numpr.find(qn("w:numId"))
            ilvl = numpr.find(qn("w:ilvl"))
            if numid is not None:
                # OOXML：`w:numId="0"` 表示**取消编号**（不受样式里的编号影响）
                # ⇒ 等同于没有自动编号，不能当成"反算不出的编号"。
                value = attrs_of(numid).get("val")
                para.num_id = value if value != "0" else None
            if ilvl is not None:
                para.num_level = int(attrs_of(ilvl).get("val") or 0)
        style = ppr_element.find(qn("w:pStyle"))
        if style is not None:
            para.style_id = attrs_of(style).get("val", "")

    for node in element.iter():
        if not isinstance(node.tag, str):
            continue
        name = local_name(node)
        if name == "t":
            para.text += node.text or ""
        elif name == "tab":
            # 制表符是判“手打目录条目”的关键（页码靠制表位拉过去），不能丢
            para.text += "\t"
        elif name in ("br", "cr"):
            # 软换行/回车符：python-docx（引擎那一侧）读成 "\n"，跟着读——
            # 用户 2026-09-20 裁决“带软换行的段不认标题”，判据侧靠“标题正则
            # 不跨行”自然实现（此前这里把 w:br 丢了，判据就多认了）。
            para.text += "\n"
        elif name == "instrText":
            para.fields.append(node.text or "")
        elif name == "fldSimple":
            para.fields.append(attrs_of(node).get("instr", ""))
        elif name == "drawing":
            para.has_drawing = True
        elif name == "pict":
            para.has_pict = True
    return para


def read_styles(root) -> tuple[dict[str, dict], dict[str, dict[str, dict[str, str]]]]:
    """styles.xml → (按 styleId 索引的样式表, docDefaults)。"""
    styles: dict[str, dict] = {}
    defaults: dict[str, dict[str, dict[str, str]]] = {"pPr": {}, "rPr": {}}
    if root is None:
        return styles, defaults
    doc_defaults = root.find(qn("w:docDefaults"))
    if doc_defaults is not None:
        for key, tag in (("rPr", "w:rPrDefault"), ("pPr", "w:pPrDefault")):
            holder = doc_defaults.find(qn(tag))
            inner = holder.find(qn(f"w:{key}")) if holder is not None else None
            defaults[key] = child_attrs(inner)
    for style in root.findall(qn("w:style")):
        attrs = attrs_of(style)
        style_id = attrs.get("styleId", "")
        name_element = style.find(qn("w:name"))
        styles[style_id] = {
            "type": attrs.get("type", ""),
            "name": (attrs_of(name_element).get("val", "")
                     if name_element is not None else ""),
            "basedOn": attrs_of(style.find(qn("w:basedOn"))).get("val", "")
            if style.find(qn("w:basedOn")) is not None else "",
            "pPr": child_attrs(style.find(qn("w:pPr"))),
            "rPr": child_attrs(style.find(qn("w:rPr"))),
        }
    return styles, defaults


def read_relationships(raw: dict[str, bytes]) -> dict[str, str]:
    """`word/_rels/document.xml.rels` → `{rId: 目标部件名}`（如 `rId12`→`word/header2.xml`）。

    为什么需要：节里的页眉/页脚引用存的是 **rId**，而 rId 只是个流水号——
    中间插一段就会让它们整体挪位（`rId12→rId11`）。直接比 rId 会把这种
    "编号挪位但语义没变"报成差异（实测踩过）。解析成部件名之后就稳了。
    """
    name = "word/_rels/document.xml.rels"
    if name not in raw:
        return {}
    out: dict[str, str] = {}
    for rel in etree.fromstring(raw[name]):
        rid = rel.get("Id")
        target = rel.get("Target")
        if not rid or not target:
            continue
        if target.startswith("/"):
            out[rid] = target.lstrip("/")
        else:
            out[rid] = "word/" + target.lstrip("./")
    return out


def read_sections(root, rels: dict[str, str] | None = None) -> list[dict[str, str]]:
    """所有的 `w:sectPr`（段落里的分节符 + 正文末尾那一个），按文档顺序。"""
    rels = rels or {}
    out: list[dict[str, str]] = []
    if root is None:
        return out
    for sect in root.iter(qn("w:sectPr")):
        info: dict[str, str] = {}
        sizing = sect.find(qn("w:pgSz"))
        if sizing is not None:
            info.update({f"pgSz.{k}": v for k, v in attrs_of(sizing).items()})
        margin = sect.find(qn("w:pgMar"))
        if margin is not None:
            info.update({f"pgMar.{k}": v for k, v in attrs_of(margin).items()})
        # 页码格式 / 起始页码写在节属性里，不在页脚部件里——很容易漏看
        num_type = sect.find(qn("w:pgNumType"))
        if num_type is not None:
            info.update({f"pgNumType.{k}": v for k, v in attrs_of(num_type).items()})
        title_page = sect.find(qn("w:titlePg"))
        info["titlePg"] = "1" if title_page is not None else ""
        if sect.get(qn("w:type")):
            info["type"] = sect.get(qn("w:type"))
        for ref in sect.findall(qn("w:headerReference")):
            rid = attrs_of(ref).get("id", "")
            info[f"header.{attrs_of(ref).get('type', '')}"] = rels.get(rid, rid)
        for ref in sect.findall(qn("w:footerReference")):
            rid = attrs_of(ref).get("id", "")
            info[f"footer.{attrs_of(ref).get('type', '')}"] = rels.get(rid, rid)
        out.append(info)
    return out


def read_docx(path: pathlib.Path) -> DocShape:
    """读一份 docx，返回它的形状。"""
    path = pathlib.Path(path)
    raw, digests = read_package(path)
    shape = DocShape(path=path, parts=digests)

    document_root = etree.fromstring(raw["word/document.xml"])
    styles_root = (etree.fromstring(raw["word/styles.xml"])
                   if "word/styles.xml" in raw else None)
    shape.styles, shape.doc_defaults = read_styles(styles_root)
    shape.sections = read_sections(document_root, read_relationships(raw))

    body = document_root.find(qn("w:body"))
    if body is None:
        shape.warnings.append("没有 w:body")
        return shape
    if body.find(qn("w:sdt")) is not None:
        shape.warnings.append(
            "body 里还有 w:sdt（内容控件）——段号与管线口径可能不一致，"
            "先跑 cleanup.unwrap_section_break_sdt 再比")

    index = 0
    for child in body:
        if not isinstance(child.tag, str):
            continue
        name = local_name(child)
        if name == "p":
            para = read_paragraph(child, index=index, in_table=False)
            para.style_name = shape.styles.get(para.style_id, {}).get("name", "")
            shape.paras.append(para)
            index += 1
        elif name == "tbl":
            table_rows: list[list[Para]] = []
            for row_index, row in enumerate(child.findall(qn("w:tr"))):
                row_cells: list[Para] = []
                for cell in row.findall(qn("w:tc")):
                    for inner in cell.findall(qn("w:p")):
                        para = read_paragraph(
                            inner, index=None, in_table=True,
                            table_path=(len(shape.tables), row_index, len(row_cells)))
                        para.style_name = shape.styles.get(
                            para.style_id, {}).get("name", "")
                        row_cells.append(para)
                table_rows.append(row_cells)
            shape.tables.append(table_rows)
            header = child.find(qn("w:tr"))
            repeat = False
            if header is not None:
                trpr = header.find(qn("w:trPr"))
                repeat = trpr is not None and trpr.find(qn("w:tblHeader")) is not None
            shape.table_repeat_header.append(repeat)

    # 页眉/页脚/脚注：差分时也要看（页码、页眉线都在那里）
    for part_name, bucket in (("header", shape.header_paras),
                              ("footer", shape.footer_paras),
                              ("footnotes.xml", shape.footnote_paras)):
        if part_name == "footnotes.xml":
            candidates = ["word/footnotes.xml"]
        else:
            candidates = sorted(n for n in raw if n.startswith(f"word/{part_name}"))
        for name in candidates:
            if name not in raw:
                continue
            root = etree.fromstring(raw[name])
            for position, element in enumerate(root.iter(qn("w:p"))):
                para = read_paragraph(element, index=None, in_table=False)
                para.style_name = shape.styles.get(para.style_id, {}).get("name", "")
                bucket.append(para)
                del position
    return shape


# ------------------------------------------------------------------ 有效格式


PROP_ORDER = ("rFonts", "sz", "szCs", "b", "i", "u", "vertAlign", "color")


def _merge(target: dict[str, str], cell: dict[str, dict[str, str]]) -> None:
    """把一层（docDefaults / 样式 / 直接格式）盖到结果上。

    `#sig` 也一起盖：它是元素子树的签名，专门用来让「子元素层面的差异」
    可见（页眉线、制表位…），而属性表本身是看不见这些的。
    """
    for key, value in (cell or {}).items():
        target[key] = ";".join(f"{k}={v}" for k, v in sorted(value.items()))


def effective_run_format(shape: DocShape, para: Para) -> dict[str, str]:
    """段落的"有效字符格式"：docDefaults ← 样式链 ← 段落标记的直接格式。

    ⚠️ 只算"样式 + 段落标记"这一层；**单个 run 上的直接格式没算进来**
    （那是逐 run 的事，报告里单列）。这样同一段里 run 之间的差异不会互相淹没。
    """
    out: dict[str, str] = {}
    _merge(out, shape.doc_defaults.get("rPr", {}))
    chain: list[str] = []
    style_id = para.style_id
    seen: set[str] = set()
    while style_id and style_id in shape.styles and style_id not in seen:
        seen.add(style_id)
        chain.append(style_id)
        style_id = shape.styles[style_id].get("basedOn", "")
    for style_id in reversed(chain):                 # 从最底层往上盖
        _merge(out, shape.styles[style_id].get("rPr", {}))
    _merge(out, para.marker_rpr)
    return {k: v for k, v in out.items() if k in PROP_ORDER or k == "tabs"}


def effective_para_format(shape: DocShape, para: Para) -> dict[str, str]:
    """段落的"有效段落格式"：docDefaults ← 样式链 ← 直接格式。"""
    out: dict[str, str] = {}
    _merge(out, shape.doc_defaults.get("pPr", {}))
    chain: list[str] = []
    style_id = para.style_id
    seen: set[str] = set()
    while style_id and style_id in shape.styles and style_id not in seen:
        seen.add(style_id)
        chain.append(style_id)
        style_id = shape.styles[style_id].get("basedOn", "")
    for style_id in reversed(chain):
        _merge(out, shape.styles[style_id].get("pPr", {}))
    for key, value in para.ppr.items():
        if key.startswith("_"):
            continue
        out[key] = ";".join(f"{k}={v}" for k, v in sorted(value.items()))
    return out


if __name__ == "__main__":                                   # 手工看一眼
    import sys
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    target = pathlib.Path(sys.argv[1])
    shape = read_docx(target)
    print(f"{target.name}：段 {len(shape.paras)}、表 {len(shape.tables)}、"
          f"页眉段 {len(shape.header_paras)}、页脚段 {len(shape.footer_paras)}、"
          f"脚注段 {len(shape.footnote_paras)}、节 {len(shape.sections)}")
    for warning in shape.warnings:
        print(f"  ⚠️ {warning}")
    for para in shape.paras[:8]:
        fmt = effective_para_format(shape, para)
        print(f"  {para.label():<6} 样式={para.style_id!r}/{para.style_name!r} "
              f"jc={fmt.get('jc', '')} ind={fmt.get('ind', '')[:40]!r} "
              f"文字={para.text[:26]!r}")
