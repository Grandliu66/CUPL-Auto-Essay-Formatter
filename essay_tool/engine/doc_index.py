# -*- coding: utf-8 -*-
"""索引法 · 第一层：把文档读成一份**完整清单**（"索引"）。

做法由用户在 2026-09-17 定，名字也由用户定：**索引法**。三句话：

  1. **索引** ＝ 成品的完整清单（每一行：照抄原文档的哪一项 / 生成哪一项），
     覆盖原文档的**所有**顶层对象；
  2. **加工** ＝ 拿原文档的副本、清空正文，**按索引顺序逐行往后追加**——
     不存在"插进既存内容中间"这个动作，所以"插入点算错"这类 bug 从根上消失
     （2026-09-17 的副标题 bug 就是这个：拿原文档的旧段号去当插入点，而那时
     前面已经插进 4 段，旧段号早就失效了）；
  3. **保障** ＝ 导完对账：**只增不减 + 位置参数/对象一字不改**。

本模块只做"读"：`document` → `list[Entry]`。不 import `pipeline`／`apply_structure`，
可以单独测。

**覆盖面**：body 层的**全部顶层子元素**（`w:p` / `w:tbl` / `w:sdt` / 末尾的
`w:sectPr`），**包括结构面板故意不显示的那些**——"面板可以显示得少，但索引不能少"。
2026-09-17 实拍（`tools/_probe_image_rows.py`，20 份夹具）：图片独占一段的段落
（`paragraph.text` 读不到图片，被当空段跳过）、挂分节符的段落、表格里的段落
（不在 `document.paragraphs` 里）在面板上**都没有行**。这些一旦漏掉，"按索引重建"
就等于"丢内容"。

⚠️ 调用方要先 `unwrap_section_break_sdt(document, keep_content=True)`，跟
`adjust_preview.build_rows()`／`pipeline.run()` 拿到手动结构时的第一步一致——
Word 的目录常裹在内容控件里，拆不拆包会改变"哪些段落看得见"。
"""
from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field

from docx.oxml.ns import qn
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
VML = "{urn:schemas-microsoft-com:vml}"

# 对象的"种类"标记（`Entry.traits`）。带这些标记的条目，加工层**一律不许删**。
T_PICTURE = "picture"       # 有图片（a:blip）
T_INLINE = "inline"         # 嵌入型图片
T_ANCHOR = "anchor"         # 浮动型图片（位置参数是相对的，见 anchor_ref）
T_VML = "vml"               # VML 兜底图形（跟 drawing 成对出现）
T_TEXTBOX = "textbox"       # 文本框（w:txbxContent）
T_FIELD = "field"           # 域（目录/清单/SEQ/PAGE…）
T_TOC_FIELD = "tocField"    # 目录/清单域
T_REVISION = "revision"     # 未接受的修订（w:ins/w:del）
T_SECTPR = "sectPr"         # 这一段挂着分节符（页眉/页码的分界）
T_TABLE = "table"           # 表格
T_SDT = "sdt"               # 内容控件
T_EMPTY_TEXT = "emptyText"  # 文字为空——**面板会把它当空段跳过**（这才是关键标记）
T_EMPTY = "empty"           # 文字为空、而且里面真的什么对象也没有
T_BODY_SECTPR = "bodySectPr"  # body 末尾那一个 sectPr


def _sig(element) -> str:
    """子树指纹（SHA1 前 16 位）。lxml 的序列化是确定的，可以拿它当"一字不改"的凭据。"""
    if element is None:
        return ""
    return hashlib.sha1(etree.tostring(element)).hexdigest()[:16]


def _count(node, tag: str) -> int:
    return len(node.findall(f".//{tag}"))


@dataclass
class Entry:
    """索引里的一行。"""

    top_index: int                     # 它在 body 里是第几个顶层子元素
    kind: str                          # paragraph / table / sdt / sectPr
    origin: str = "source"             # source＝照抄原文档某一项；generated＝生成（占位语/域）
    text: str = ""                     # 正文文字（给人看，也当身份指纹的一部分）
    traits: tuple[str, ...] = ()       # 对象种类标记（见上面的 T_*）
    counts: dict = field(default_factory=dict)   # {"picture":1,"inline":1,"row":3,...}
    pos_sig: str = ""                  # 位置参数指纹（pPr / tblPr / sdtPr）
    content_sig: str = ""              # 身份指纹（文字 + 对象计数）
    anchor_ref: tuple = ()             # 浮动对象：(inline|anchor, relativeFrom, 参数摘要)
    part: str | None = None            # 身份（部件名）——由识别层填，本模块留空
    must_exist: bool = True            # 一条都不许丢（源条目恒为 True）
    keep_bytes: bool = False           # **目标**标记：索引法落地后要求产物里逐字节相同
    note: str = ""                     # 人看的补充说明（比如"图片独占一段"）
    # 识别层的附加信息（2026-09-18 起，见 engine.recognize）：标题层级/角色/
    # 自动编号固化文字/标签是"标签位"还是"内容位"……本模块不关心具体键名，
    # 只负责搬运——识别层往这里填，加工层从这里读，两层通过同一份 Entry 对接。
    meta: dict = field(default_factory=dict)

    def label(self) -> str:
        who = self.part or ("（空段）" if T_EMPTY_TEXT in self.traits else "")
        head = f"[{self.top_index:>3}] {self.kind:<9}"
        tags = ",".join(self.traits) or "-"
        return f"{head} {tags:<28} {who:<8} {self.text[:34]!r}"


def _traits_of(node) -> list[str]:
    traits: list[str] = []
    if node.findall(f".//{A}blip"):
        traits.append(T_PICTURE)
    if node.findall(f".//{WP}inline"):
        traits.append(T_INLINE)
    if node.findall(f".//{WP}anchor"):
        traits.append(T_ANCHOR)
    if node.findall(f".//{VML}shape"):
        traits.append(T_VML)
    if node.findall(f".//{W}txbxContent"):
        traits.append(T_TEXTBOX)
    if node.findall(f".//{W}ins") or node.findall(f".//{W}del"):
        traits.append(T_REVISION)
    if (node.findall(f".//{W}fldChar") or node.findall(f".//{W}fldSimple")
            or node.findall(f".//{W}instrText")):
        traits.append(T_FIELD)
    return traits


def _counts_of(node) -> dict:
    return {
        "picture": _count(node, f"{A}blip"),
        "inline": _count(node, f"{WP}inline"),
        "anchor": _count(node, f"{WP}anchor"),
        "shape": _count(node, f"{VML}shape"),
        "textbox": _count(node, f"{W}txbxContent"),
        "field": len(node.findall(f".//{W}fldChar")) + len(node.findall(f".//{W}fldSimple")),
        "tbl": _count(node, f"{W}tbl"),
        "tr": _count(node, f"{W}tr"),
        "tc": _count(node, f"{W}tc"),
    }


def _anchor_refs(node) -> tuple:
    """浮动对象的位置参数摘要：相对谁定位（relativeFrom）＋偏移/对齐。

    ⚠️ `relativeFrom="paragraph"`／`"line"`／`"column"` 的浮动对象**跟着锚点段落走**
    ——锚点段落在文档流里后移几段，它就在版面上后移几段。这是"躲不掉"的那一类
    （用户在 Word 里插段落也一样），索引里记下来、导出后如实告知。
    `relativeFrom="page"`／`"margin"` 的不受宿主段落位置影响。
    """
    out = []
    for kind, tag in (("inline", f"{WP}inline"), ("anchor", f"{WP}anchor")):
        for item in node.findall(f".//{tag}"):
            for axis in ("positionH", "positionV"):
                pos = item.find(f"{WP}{axis}")
                if pos is None:
                    continue
                ref = pos.get("relativeFrom", "")
                align = pos.find(f"{WP}align")
                offset = pos.find(f"{WP}posOffset")
                out.append((kind, axis, ref,
                            (align.text if align is not None else None)
                            or (offset.text if offset is not None else None) or ""))
    return tuple(out)


def _paragraph_entry(node, index: int) -> Entry:
    ppr = node.find(f"{W}pPr")
    text = "".join(t.text or "" for t in node.findall(f".//{W}t"))
    traits = _traits_of(node)
    if ppr is not None and ppr.find(f"{W}sectPr") is not None:
        traits.append(T_SECTPR)
    counts = _counts_of(node)
    contents = sum(counts[k] for k in ("picture", "shape", "field", "textbox"))
    if not text.strip():
        # ⚠️ 两个标记要分开：`emptyText` 是"面板会跳过它"（图片独占一段的段落正是
        # 这一类：`paragraph.text` 读不到图片）；`empty` 才是真的什么都没有。
        # 2026-09-17 实拍：PaperTest[51]、test2[43]、课程考核封面[2][32]、
        # 圧力测试_Agent版[62]、图表压力测试[26]、图表脚注测试[51]、民总-排好版[2]
        # 全部是 emptyText（面板无行）——它们是"按索引重建"时最容易丢的那批。
        traits.append(T_EMPTY_TEXT)
        if not contents:
            traits.append(T_EMPTY)
    content = "|".join([text] + [f"{k}={v}" for k, v in counts.items() if v])
    note = ""
    if set(traits) & {T_PICTURE, T_ANCHOR, T_INLINE, T_VML, T_TEXTBOX}:
        note = ("图片/图形独占一段（面板上看不见：paragraph.text 读不到它）"
                if T_EMPTY_TEXT in traits else "图与文字同段")
    return Entry(top_index=index, kind="paragraph", text=text, traits=tuple(traits),
                 counts=counts, pos_sig=_sig(ppr),
                 content_sig=hashlib.sha1(content.encode()).hexdigest()[:16],
                 anchor_ref=_anchor_refs(node), note=note)


def build_index(document) -> list[Entry]:
    """把文档读成索引（body 层**全部**顶层子元素，一条不落）。"""
    body = document.element.body
    out: list[Entry] = []
    for index, child in enumerate(body):
        if not isinstance(child.tag, str):          # 注释/处理指令
            out.append(Entry(top_index=index, kind="other",
                             traits=("other",), note="注释/处理指令（原样保留）"))
            continue
        if child.tag == f"{W}p":
            out.append(_paragraph_entry(child, index))
        elif child.tag == f"{W}tbl":
            counts = _counts_of(child)
            text = "".join(t.text or "" for t in child.findall(f".//{W}t"))
            out.append(Entry(
                top_index=index, kind="table", text=text, traits=(T_TABLE,),
                counts=counts, pos_sig=_sig(child.find(f"{W}tblPr")),
                content_sig=hashlib.sha1(
                    ("|".join([text] + [f"{k}={v}" for k, v in counts.items() if v])
                     ).encode()).hexdigest()[:16],
                anchor_ref=_anchor_refs(child),
                note=f"{counts['tr']} 行 × {counts['tc']} 格"
                     + (f"、内含 {counts['picture']} 张图" if counts["picture"] else "")))
        elif child.tag == f"{W}sdt":
            text = "".join(t.text or "" for t in child.findall(f".//{W}t"))
            counts = _counts_of(child)
            out.append(Entry(
                top_index=index, kind="sdt", text=text, traits=(T_SDT,),
                counts=counts, pos_sig=_sig(child.find(f"{W}sdtPr")),
                content_sig=hashlib.sha1(text.encode()).hexdigest()[:16],
                anchor_ref=_anchor_refs(child),
                note="内容控件（目录常裹在这里面）"))
        elif child.tag == f"{W}sectPr":
            out.append(Entry(top_index=index, kind="sectPr",
                             traits=(T_BODY_SECTPR, T_SECTPR),
                             pos_sig=_sig(child),
                             content_sig=_sig(child),
                             note="body 末尾的分节定义"))
        else:
            # 2026-09-19 补：书签范围（w:bookmarkStart/End）、批注范围
            # （w:commentRangeStart/End）这类贴在段落之间、不构成段落/表格/
            # 内容控件的顶层子元素——以前这四个分支之外的一律不追加 Entry，
            # `top_index` 序列因此出现空洞，这些元素也就从没进过索引，
            # `rebuild_body()` 自然也就没法把它们复制到成品里（实测：中国
            # 政法大学硕士、博士学位论文写作规范（更新）.docx 里有 19 个
            # 顶层 w:bookmarkEnd 一直被这样漏掉），违反本模块开头"body 层
            # 全部顶层子元素、一条不落"的承诺。跟"注释/处理指令"那条分支
            # 一样，原样保留、不作判断。
            out.append(Entry(top_index=index, kind="other",
                             traits=("other",), pos_sig=_sig(child),
                             content_sig=_sig(child),
                             note=f"未识别的顶层元素（{child.tag.split('}')[-1]}，原样保留）"))
    return out


# ============================================================ TOC 域 / 样式：结构性事实
#
# 这几个函数只回答"文档里这里有没有一个 TOC 家族的域/条目"，不做任何识别
# 判断——2026-09-18 从 `detect.py` 搬过来（那边、`cleanup.py`、`alerts.py`
# 三处各自 import 或各自抄一份，是典型的"功能重复"）。搬到这里是因为
# `doc_index.py` 本来就是"读文档结构性事实、不做判断"的那一层，`recognize.py`
# 和 `cleanup.py` 都只需要事实，不需要 detect.py 那一整套识别架构。

_TOC_INSTR_RE = re.compile(r"^\s*TOC(\s|$)", re.IGNORECASE)
_LIST_FIELD_KIND_RE = re.compile(r"\\c\s+\"?([^\"\s\\]+)\"?")
_LIST_FIELD_STYLE_RE = re.compile(r'\\t\s+"([^"]*)"')
_STYLE_NAME_KINDS = {"图题": "图", "表题": "表"}


def paragraph_field_codes(paragraph) -> list[str]:
    """段落里出现的域指令原文（`w:instrText` 的文本 + `w:fldSimple/@w:instr`）。"""
    codes: list[str] = []
    for el in paragraph._p.iter(qn("w:instrText")):
        if el.text:
            codes.append(el.text)
    for el in paragraph._p.iter(qn("w:fldSimple")):
        instr = el.get(qn("w:instr"))
        if instr:
            codes.append(instr)
    return codes


def has_toc_field(paragraph) -> bool:
    """这一段里是不是目录/清单域（TOC 家族）。"""
    return any(_TOC_INSTR_RE.match(code) for code in paragraph_field_codes(paragraph))


def list_kinds_in_instr(code: str) -> set[str]:
    r"""一条域指令里"收了哪几种题注"——`\c "图"` 与 `\t "图题,1"` 两种写法都认
    （后者是我们自己写的清单域，只认 `\c` 会把它判成"目录"）。"""
    kinds = {m.group(1).strip() for m in _LIST_FIELD_KIND_RE.finditer(code)}
    for match in _LIST_FIELD_STYLE_RE.finditer(code):
        for item in match.group(1).split(","):
            kind = _STYLE_NAME_KINDS.get(item.strip())
            if kind is not None:
                kinds.add(kind)
    return kinds


def list_field_kinds(paragraph) -> set[str]:
    kinds: set[str] = set()
    for code in paragraph_field_codes(paragraph):
        if not _TOC_INSTR_RE.match(code):
            continue
        kinds |= list_kinds_in_instr(code)
    return kinds


def toc_field_part(paragraph) -> str:
    """这个域属于哪个部件：目录还是哪张清单。"""
    kinds = list_field_kinds(paragraph)
    if "图" in kinds and "表" in kinds:
        return "插图和附表清单"
    if "图" in kinds:
        return "插图清单"
    if "表" in kinds:
        return "附表清单"
    return "目录"


def is_toc_paragraph(paragraph) -> bool:
    """目录条目（样式叫 `toc 1`/`toc 2`…）。"""
    return str(paragraph.style.name).strip().lower().startswith("toc ")


def is_list_entry_paragraph(paragraph) -> bool:
    """图表目录/清单的条目段落（Word 内置样式 `table of figures`）。"""
    return str(paragraph.style.name).strip().lower() == "table of figures"


def _fld_char_kinds(paragraph) -> list[str]:
    out: list[str] = []
    for el in paragraph._p.iter(qn("w:fldChar")):
        out.append(el.get(qn("w:fldCharType")) or "")
    return out


def toc_field_span_indices(paragraphs, parts: set[str] | None = None) -> set[int]:
    """与"TOC 家族域"的区间相交的段落下标（含域渲染出来的条目段）。

    不闭合的域一律不算区间（未闭合时深度永远不归零，会把余下全文当成域内、
    静默吞掉——2026-09-17 AAT 穷举测出来的"回炉不收敛"根因之一）。
    """
    marked: set[int] = set()
    stack: list[tuple[int, bool]] = []
    for index, paragraph in enumerate(paragraphs):
        owned = has_toc_field(paragraph) and (parts is None or toc_field_part(paragraph) in parts)
        for kind in _fld_char_kinds(paragraph):
            if kind == "begin":
                stack.append((index, False))
            elif kind == "end":
                if stack:
                    start, is_toc = stack.pop()
                    if is_toc:
                        marked.update(range(start, index + 1))
        if owned and stack:
            start, _was_toc = stack[-1]
            stack[-1] = (start, True)
    return marked


def _body_instr_texts(document) -> list[str]:
    body = document._element.body
    found = [el.text or "" for el in body.iter(qn("w:instrText"))]
    for el in body.iter(qn("w:fldSimple")):
        instr = el.get(qn("w:instr"))
        if instr:
            found.append(instr)
    return found


def body_toc_field_parts(document) -> set[str]:
    """整篇文档里已经存在哪些目录/清单域（含内容控件里的）。"""
    parts: set[str] = set()
    for code in _body_instr_texts(document):
        if not _TOC_INSTR_RE.match(code):
            continue
        kinds = list_kinds_in_instr(code)
        if "图" in kinds and "表" in kinds:
            parts.add("插图和附表清单")
        elif "图" in kinds:
            parts.add("插图清单")
        elif "表" in kinds:
            parts.add("附表清单")
        else:
            parts.add("目录")
    return parts


def body_list_field_kinds(document) -> set[str]:
    kinds: set[str] = set()
    for code in _body_instr_texts(document):
        if not _TOC_INSTR_RE.match(code):
            continue
        kinds |= list_kinds_in_instr(code)
    return kinds


def summarize(index: list[Entry]) -> str:
    """一句话统计（打印用）。"""
    def n(trait: str) -> int:
        return sum(1 for e in index if trait in e.traits)

    return (f"条目 {len(index)}：段落 {sum(1 for e in index if e.kind == 'paragraph')}"
            f"／表格 {n(T_TABLE)}／内容控件 {n(T_SDT)}／body-sectPr {n(T_BODY_SECTPR)}"
            f"；图片 {n(T_PICTURE)}（浮动 {n(T_ANCHOR)}）／文本框 {n(T_TEXTBOX)}"
            f"／域 {n(T_FIELD)}／修订 {n(T_REVISION)}／分节符载体 {n(T_SECTPR)}"
            f"／文字为空的段 {n(T_EMPTY_TEXT)}（真空白 {n(T_EMPTY)}）")


def compare(before: list[Entry], after: list[Entry], *,
            strict_bytes: bool = False) -> list[str]:
    """对账：源索引 `before` 对 产物索引 `after`，返回问题清单（空＝通过）。

    判据三条：
      1. **只增不减**：`before` 里每一条都要在 `after` 里按顺序找得到
         （用身份指纹匹配；找不到就是"丢了/串了"）；
      2. **对象总数不减**：图片 / 浮动图 / 文本框 / 域 / 表格 / 分节符载体，
         合计不能比源少；
      3. `strict_bytes=True` 时，标了 `keep_bytes` 的条目在产物里必须**逐字节相同**
         （位置参数一字不改）。⚠️ 现在**不要开**这条：眼下加工层是"在原文上改"
         （清直接格式会扫过每一段），字节级不变要等索引法落地（改成"按索引重建"）
         才成立——本参数是给那一步准备的目标契约。
    """
    problems: list[str] = []
    if len(after) < len(before):
        problems.append(f"条目少了 {len(before) - len(after)} 条"
                        f"（源 {len(before)} → 产物 {len(after)}）")
    # 贪心按序匹配（顺序一致才认得出来；对不上就是丢/串）
    cursor = 0
    for entry in before:
        hit = None
        for i in range(cursor, len(after)):
            if after[i].content_sig == entry.content_sig \
                    and after[i].kind == entry.kind:
                hit = i
                break
        if hit is None:
            problems.append(f"找不到这一条（源第 {entry.top_index} 项 / {entry.kind}"
                            f" / {entry.text[:20]!r} / {','.join(entry.traits) or '-'}）")
            continue
        cursor = hit + 1
        if strict_bytes and entry.keep_bytes and entry.content_sig != after[hit].content_sig:
            problems.append(f"标了『原样保留』的条目被改动了：源第 {entry.top_index} 项")
        if strict_bytes and entry.keep_bytes and entry.pos_sig != after[hit].pos_sig:
            problems.append(f"标了『原样保留』的条目，位置参数被改了："
                            f"源第 {entry.top_index} 项")
    for trait, name in ((T_PICTURE, "图片"), (T_ANCHOR, "浮动图"), (T_TEXTBOX, "文本框"),
                        (T_FIELD, "域"), (T_TABLE, "表格"), (T_SECTPR, "分节符载体"),
                        (T_REVISION, "修订")):
        was = sum(1 for e in before if trait in e.traits)
        now = sum(1 for e in after if trait in e.traits)
        if now < was:
            problems.append(f"{name}少了：源 {was} → 产物 {now}")
    return problems


def dump(index: list[Entry], *, limit: int | None = None) -> str:
    """人看的清单（调试/给用户过目用）。"""
    head = summarize(index)
    lines = [head, "-" * 96]
    for entry in index[:limit]:
        lines.append(entry.label())
    if limit is not None and len(index) > limit:
        lines.append(f"...（还有 {len(index) - limit} 条）")
    return "\n".join(lines)


def open_and_index(path) -> tuple:
    """读一份 docx 并建索引（**先拆内容控件，跟管线同一口径**）。"""
    from docx import Document

    from essay_tool.engine.cleanup import unwrap_section_break_sdt

    document = Document(str(path))
    unwrap_section_break_sdt(document, keep_content=True)
    return document, build_index(document)


# ============================================================ 第二层：执行器
#
# 用户 2026-09-17 定的做法：**按索引顺序逐个往后追加**——每一步都只有"已经生成的
# 上面部分"，没有"下面"，所以不存在"插进既存内容中间"这个动作，
# "插入点算错"这类 bug（2026-09-17 副标题那个）从根上没有生存空间。
#
# 关键实现技巧：**不能从空白文档起手**。图片本体、页眉页脚、脚注、样式表、编号
# 定义、批注……都靠"关系"指向文档里别的 part，从空白起手就全断了。做法是拿
# **原文档的副本**、只把正文清空，再把每一项搬回去——关系表原封不动，
# 搬一项不会丢任何引用。


@dataclass
class Op:
    """加工层的一条指令（第二层的输入）＝**索引里的一行**。

    * `action="copy"`：照抄源文档 `entry.top_index` 那一项（原样搬过去）；
    * `action="generate"`：**在这里生成**一段新内容（占位语/域），身份是 `part`。

    `out_index`＝**它在成品里是第几条**（由 `build_ops()` 统一填好）。用户
    2026-09-17："不光是有每一段原来的信息，而且有新文档应该在哪的信息"——
    每行的**出处**是 `entry.top_index`，**新位置**就是 `out_index`。
    """

    action: str
    entry: Entry | None = None
    text: str = ""
    part: str | None = None
    style_id: str | None = None
    # 生成什么：`placeholder`＝占位语（默认）／`toc_field`＝目录域／
    # `list_field`＝清单域（`text` 里写 "图"/"表"）。
    kind: str = "placeholder"
    # 空位行是“标签位”还是“内容位”（`Row.slot_col`）。摘要/关键词这类部件在面板上
    # 有**两行**空位（标签 + 内容）⇒ 两条生成指令都叫“摘要”，执行器必须分得清
    # 哪一条该生成【摘要】标签、哪一条该生成内容占位语（2026-09-17 实测：真实面板
    # 上确实会出现“第 5 条：生成 摘要”连着“第 6 条：生成 摘要”）。
    #   "label"      ＝造整个部件（标签＋内容占位；面板上没有单独"内容位"时用）
    #   "label_only" ＝只造标签（同一部件另有"内容位"那条指令负责内容，避免占位语重复）
    #   "text"       ＝只造内容
    slot_col: str = ""
    out_index: int = -1

    def label(self) -> str:
        if self.action == "copy" and self.entry is not None:
            who = f"身份：{self.entry.part}" if self.entry.part else self.entry.kind
            tail = f"  {self.entry.text[:20]!r}" if self.entry.text else ""
            return (f"第 {self.out_index} 条：照抄 原文档第 {self.entry.top_index} 项"
                    f"（{who}）{tail}")
        where = "标签位" if self.slot_col in ("label", "label_only") else (
            "内容位" if self.slot_col == "text" else "位置")
        what = {"toc_field": "目录域", "list_field": f"{self.text}清单域"}.get(
            self.kind, "占位语")
        return f"第 {self.out_index} 条：生成 {self.part or '未命名'} 的{where}（{what}）"


def copy_all_ops(index: list[Entry]) -> list[Op]:
    """1:1 照抄（**零失真基线**：这样重建出来的文档应当与原文逐条一致）。"""
    ops = [Op("copy", entry=entry) for entry in index]
    for k, op in enumerate(ops):
        op.out_index = k
    return ops


def paragraph_ordinals(index: list[Entry]) -> dict:
    """`{段落序号: Entry}`。

    ⚠️ 两套下标**不是一套数**：`Entry.top_index` 是 **body 顶层序号**（段落、表格、
    分节定义都算），而结构面板的 `Row.doc_index` 是 **段落序号**（`document.paragraphs`
    的下标，表格与分节定义不算）。换算只放在这一个地方。
    """
    out: dict[int, Entry] = {}
    ordinal = 0
    for entry in index:
        if entry.kind == "paragraph":
            out[ordinal] = entry
            ordinal += 1
    return out


def compute_span_ownership(rows) -> tuple[dict[int, str], list[tuple[int, int, str]]]:
    """按 `rows` **当前顺序**现场划一遍"块"——不看 `Row.block` 这个字段本身
    是不是准确、有没有被更新过，只看"这一行现在排在哪个标签/整行部件后面"。

    2026-09-19 用户拍板的模型："块"不是贴在段落上、需要维护的持久标签，
    是"当前排在哪个标签行后面"这个位置关系本身——标签一挪，跟着它的范围
    立刻跟着变，用户把内容行手动挪到别的标签下面之后，归属立刻改口，不用
    等谁去回填一个 `.block` 字段（例：摘要下两段内容，用户把第二段拖到了
    关键词下面，那一段现在就该算关键词的，不管它当初是识别成摘要的）。

    每碰到一个"标签行"（`kind` 是 label/title/subtitle/entitle 之一）或者
    "进入正文区的第一行"，就开一个新的块——正文区（标题+其下内容）整体
    当**一整块**（部件名固定叫"正文"，用户原话"正文整个算一块"），内部的
    标题/正文各自的身份走它们自己那套逻辑（层级列/引言结论锚点），不需要
    这个函数插手，所以 `owner` 里不包含正文区的行。`slot`/`zone`（区带）
    这两种行会打断当前块但不会开新块——它们没有"内容"这个概念。

    返回 `(owner, spans)`：`owner` 是 `{rows 下标: 部件名}`，只覆盖前/后区
    的内容行，`build_ops()` 用它给这些行定身份；`spans` 是
    `[(start, end, 部件名), ...]`半开区间列表（按当前顺序排），"拖动整块"
    的手势用它找"这一行现在属于哪一段连续区间"。
    """
    owner: dict[int, str] = {}
    spans: list[tuple[int, int, str]] = []
    cur_part: str | None = None
    cur_start: int | None = None
    prev_zone: str | None = None
    for i, row in enumerate(rows):
        kind = getattr(row, "kind", "")
        zone = getattr(row, "zone", None)
        if kind in ("zone", "slot"):
            if cur_part is not None:
                spans.append((cur_start, i, cur_part))
                cur_part = None
            continue
        entering_body = zone == "body" and prev_zone != "body"
        prev_zone = zone
        if kind in ("label", "title", "subtitle", "entitle") or entering_body:
            if cur_part is not None:
                spans.append((cur_start, i, cur_part))
            cur_part = "正文" if entering_body else (
                getattr(row, "block", None) or getattr(row, "label", None) or kind)
            cur_start = i
            if cur_part != "正文":
                owner[i] = cur_part
            continue
        # 非边界行：只有"曾经被认成某个部件的标签/内容"的行（`block`/`label`
        # 不为空）才跟着当前打开的这一段走；压根没有归属过的游离正文
        # （`block`/`label` 都是 None）不该被顺手划进旁边随便哪个块——
        # 2026-09-19 实测抓到：一份没有任何可识别标题、正文整段被当"前置区
        # 无主段落"铺出来的文档，会被这样整段错划成"目录"的内容，`alerts.
        # check_duplicate_parts` 报出几百段的假"重复的部件：目录"。这种
        # 游离行直接跳过（既不加入当前段、也不打断它——它只是"看不见"）。
        has_prior_claim = bool(getattr(row, "block", None) or getattr(row, "label", None))
        if cur_part is not None and cur_part != "正文" and has_prior_claim:
            owner[i] = cur_part
    if cur_part is not None:
        spans.append((cur_start, len(rows), cur_part))
    return owner, spans


def build_ops(rows, index: list[Entry], *, enabled=None,
              insert_toc: bool = True, list_kinds=None) -> list[Op]:
    """结构面板的行序 → 指令表（**索引法里"调整结果 → 索引"的那一步**）。

    `rows`：结构面板的行（`adjust_preview.Row`，这里只按鸭子类型取
    `kind` / `doc_index` / `slot_for` / `slot_col` / `block` / `label`），
    **顺序即成品顺序**。
    `enabled`：`{部件名: 是否启用}`——空位行只有在**启用**时才生成占位语；
    没启用＝"这里什么都不输出"。
    `insert_toc`：源文件里没有真目录域时才在"目录"那一行后面跟一条**目录域**指令
    （有真域时不插第二份，用户 2026-09-15 的裁决）。
    `list_kinds`：`{清单部件名: {"图","表"}}`——同样在那一行后面跟清单域指令。
    域是**独立的指令**（一个部件可能由"标签段 + 域段"组成，域不是标签的一部分）。

    三条硬规矩（都是 2026-09-17 用户口径的直接后果）：

    1. **一条不落**：面板没铺的源条目（图片独占的段落、空段、表格、分节定义、
       目录条目……）**在这里按原顺序补成 copy 指令**——"面板可以显示得少，
       但索引不能少"；
    2. **位置＝行序**：空位行排第几，成品里就在第几条（`Op.out_index`）；
    3. **body 末尾的分节定义**：它本来在最后时**始终排最后**（否则在它后面追加
       生成条目会让文档不合规）；本来就不在最后（`平时作业.docx` 实测在第 0 位）
       时按原顺序原位保留——索引法第一原则是照抄，不替用户"纠正"文件。
    """
    enabled = enabled or {}
    list_kinds = list_kinds or {}
    p_slots = paragraph_ordinals(index)
    # 标题列（l1~l4）自成一套：身份不是 `row.block`（结构页的标题行从来不填
    # 这个字段），是**列本身**——跟 `adjust_preview.shown_text()` 算显示文字
    # 用的是同一条规矩（层级永远按当前列算，`row.heading` 只提供编号/是否
    # 自动编号这些跟层级无关的原材料，见该函数注释）。`_HEADING_PART_NAMES`
    # 在这里单独放一份是因为本模块不 import `recognize`（层次相反，见上面
    # `CONTENT_OWNER_PARTS` 那条同类注释）。
    _HEADING_PART_NAMES = {1: "一级标题", 2: "二级标题", 3: "三级标题", 4: "四级标题"}
    last_entry = index[-1] if index else None
    defer_sectpr = bool(last_entry is not None and T_BODY_SECTPR in last_entry.traits)
    kept = [r for r in rows
            if getattr(r, "kind", "") == "slot"
            or getattr(r, "doc_index", None) is not None]
    owner, _spans = compute_span_ownership(kept)

    ops: list[Op] = []
    emitted: set[int] = set()

    # 2026-09-19 修：哪些原始段落**面板上有自己的行**（不管这一行现在被拖到了
    # `rows` 里的第几位）——`flush_before` 只该替"面板压根没铺、永远不会轮到
    # 自己回合"的条目代劳（图片独占的段落、空段、表格、分节定义、目录条目……，
    # 这批还没有自己的行，仍然要靠它兜底，是下一步要补的），不能连"有行、
    # 只是这一行在新顺序里排到后面才轮到"的条目也捎带着按原文顺序倒出来——
    # 那样等真正轮到那一行时，它的 `top_index` 已经在 `emitted` 里，会被当成
    # "这一行不对应真实段落"直接跳过，拖拽重排在这种情况下对最终导出的文档
    # **完全没有效果**：面板显示已经换了顺序，成品却还是原文顺序（2026-09-19
    # 实测抓到的真实事故，不是"文档极度错乱"才会触发的边界情况，随便换两个
    # 部件的位置就能复现）。
    row_top_indices = {
        p_slots[di].top_index
        for r in kept
        for di in (getattr(r, "doc_index", None),)
        if di is not None and di in p_slots
    }
    # 折进域行的缓存条目段（`Row.attached`，段落序号）：也算"有归属"，
    # 不能被 flush_before 提前倒出；等它们的域行轮到时一起发。
    attached_tops = {
        p_slots[a].top_index
        for r in kept
        for a in (getattr(r, "attached", ()) or ())
        if a in p_slots
    }
    row_top_indices |= attached_tops

    def flush_before(limit_top: int | None) -> None:
        """把"还没发、且在原文档里排在 `limit_top` 之前、且面板上没有自己
        那一行"的条目按原顺序补上——有自己那一行的条目留给它自己的回合处理，
        不管那一行现在排在 `kept` 里的第几位。"""
        for entry in index:
            if limit_top is not None and entry.top_index >= limit_top:
                break
            if entry.top_index in emitted:
                continue
            if entry.top_index in row_top_indices:
                continue                       # 有自己的行，留给它自己的回合
            if defer_sectpr and T_BODY_SECTPR in entry.traits:
                continue                       # 留到最后再发
            ops.append(Op("copy", entry=entry))
            emitted.add(entry.top_index)

    fields_emitted: set[str] = set()

    def add_fields(part: str | None) -> None:
        """某一行的后面要不要跟域指令（目录／清单）。

        同一个部件在面板上占了两行时（重复的标签行，或"插入空位"让空位
        继承了同一个部件），两行各自的回合都会调用这里——以前没有"这个
        部件的域已经发过"这道判断，域指令会跟着发两遍，成品里就出现两个
        【目录】/两条目录域（2026-09-22 审计揪出的）。一个部件的域只该
        在整份成品里出现一次，用 `part` 去重。
        """
        if not part or part in fields_emitted:
            return
        fields_emitted.add(part)
        if part == "目录" and insert_toc:
            ops.append(Op("generate", part=part, kind="toc_field"))
        for kind in ("图", "表"):        # 规范 [91]：分开列时插图清单在前
            if kind in (list_kinds.get(part) or ()):
                ops.append(Op("generate", part=part, kind="list_field", text=kind))

    # 2026-09-20 修（用户：所有"插入标签后正文"的地方，插入语会插两遍）：
    # "标签位"那条生成指令默认把**整个部件**（标签＋内容占位）造出来；而面板上
    # 摘要/关键词/英文摘要/英文关键词/参考文献/附录/致谢同时还有一个"内容位"（空位，
    # 或用户已经拖进去的内容行），"内容位"那条指令又会再造一遍内容 ⇒ 占位语出现两次。
    # 所以：同一部件在面板上另有"内容列"的行时，标签位只造标签（`label_only`）。
    text_col_parts = {
        (getattr(r, "slot_for", None) or getattr(r, "block", None))
        for r in kept
        if (getattr(r, "kind", "") == "slot" and getattr(r, "slot_col", "") == "text")
        or getattr(r, "kind", "") == "content"
    } - {None}

    for i, row in enumerate(kept):
        kind = getattr(row, "kind", "")
        doc_index = getattr(row, "doc_index", None)
        if kind == "slot":
            nxt = next((r.doc_index for r in kept[i + 1:]
                        if getattr(r, "doc_index", None) is not None), None)
            flush_before(p_slots[nxt].top_index if nxt in p_slots else None)
            part = getattr(row, "slot_for", None)
            if part and enabled.get(part, False):
                col = getattr(row, "slot_col", "")
                if col == "label" and part in text_col_parts:
                    col = "label_only"
                ops.append(Op("generate", part=part, slot_col=col))
                add_fields(part)
            continue
        entry = p_slots.get(doc_index)
        flush_before(entry.top_index if entry is not None else None)
        if entry is None or entry.top_index in emitted:
            continue                           # 面板这一行不对应任何真实段落
        if kind == "heading" and getattr(getattr(row, "heading", None), "role", None):
            # 固定名称锚点（引言/结论）：结构页仍然把它们画在一级标题这一栏
            # 里（纯 UI 布局延续），但身份不是"一级标题"——是 `row.heading.role`
            # 本身，跟数字编号标题的处理分开，别把 level/number 这些标题专属
            # 的 meta 也写上去。
            entry.part = row.heading.role
        elif kind == "heading":
            # 数字编号标题：level 按当前所在列算（用户随时可能把它拖到别的
            # 级别），number/marker/title/canonical/auto_numbered 这些跟层级
            # 无关的原材料从 `row.heading`（鸭子类型，等价于旧
            # `detect.HeadingMatch`）读——它是 None 时（一段普通正文被拖进
            # 标题列，识别时从没认出过编号）就当"没有编号的标题"处理。
            level = (int(row.col[1]) if getattr(row, "col", "") in ("l1", "l2", "l3", "l4")
                    else None)
            entry.part = _HEADING_PART_NAMES.get(level)
            heading = getattr(row, "heading", None)
            entry.meta.update(
                level=level,
                number=getattr(heading, "number", None) if heading else None,
                marker=getattr(heading, "marker", "") if heading else "",
                title=getattr(heading, "title", "") if heading else "",
                canonical=getattr(heading, "canonical", None) if heading else None,
                auto_numbered=bool(getattr(heading, "auto_numbered", False)) if heading else False,
            )
        elif kind == "toc_field":
            # 目录/清单域本体（`adjust_preview.py` 里 `row.block` 借用部件名
            # 只是为了让它跟着所属的块一起重排，不代表它是那个部件的"标签"）
            # ——不写 `entry.part`，跟自动识别路径（`recognize.py` 的内容
            # 扫描本来就跳过域段落）保持一致。写了的话 `alerts.check_
            # duplicate_parts()` 会把"标签段 + 域段"数成两次"标签"，报出
            # 假的"重复的部件"（2026-09-19 实测抓到：目录段落 21/22 各算一次）。
            pass
        else:
            # 2026-09-19 改：不再直接信任 `row.block`（可能是一个过期的
            # 标签，参见 `compute_span_ownership` 的说明）——身份现场按
            # "当前排在哪个标签/整行部件后面"算，只有正文区之前完全没有
            # 任何标签（游离正文）这种 `owner` 算不出结果的情况，才退回读
            # `row.block`/`row.label`（兜底，正常情况下两者应该一致）。
            entry.part = owner.get(i) or getattr(row, "block", None) or getattr(row, "label", None)
        # 图题/表题：结构页勾的"图/表/续"写回索引（`Entry.meta["caption"]`），
        # 导出的告警与套样式都读这一份——手动改过的题注不再被现场重新识别
        # 冲掉。标题行/域行不可能是题注，一并清掉识别遗留的标记。
        if kind in ("caption", "body", "content"):
            special = getattr(row, "special", None)
            if special in ("图", "表"):
                entry.meta["caption"] = {
                    "kind": special, "role": getattr(row, "caption_role", None) or "main"}
            else:
                entry.meta.pop("caption", None)
        elif kind in ("heading", "toc_field"):
            entry.meta.pop("caption", None)
        ops.append(Op("copy", entry=entry))
        emitted.add(entry.top_index)
        for a in (getattr(row, "attached", ()) or ()):
            follower = p_slots.get(a)
            if follower is not None and follower.top_index not in emitted:
                ops.append(Op("copy", entry=follower))
                emitted.add(follower.top_index)
        add_fields(entry.part)

    flush_before(None)                        # 收尾：后面没铺的条目
    if defer_sectpr and last_entry.top_index not in emitted:
        ops.append(Op("copy", entry=last_entry))
        emitted.add(last_entry.top_index)
    for k, op in enumerate(ops):
        op.out_index = k
    return ops


# 规范里这几个部件出现的先后顺序（原是 `apply_structure._FRONT_ORDER`/
# `_BACK_ORDER`，2026-09-18 搬到这里，好让 `auto_ops()` 和 GUI 的
# `build_ops()` 共用同一份顺序知识，不必互相 import 对方模块）。分成前置/
# 后置两段是因为"一个都没找到时插在哪"前置和后置要的答案相反——见
# `apply_structure.py` 历史注释（"从头到尾一个部件都没有"那次真实事故）。
FRONT_ORDER = [
    "题目", "副标题", "摘要", "关键词", "英文题目", "英文副标题", "英文摘要", "英文关键词", "目录",
    "插图清单", "附表清单", "插图和附表清单", "引言",
]
BACK_ORDER = ["结论", "参考文献", "附录", "致谢"]

# 三个"清单"部件各自要收哪种题注（合并形态一页两个域）。
LIST_PART_KINDS: dict[str, frozenset[str]] = {
    "插图清单": frozenset({"图"}),
    "附表清单": frozenset({"表"}),
    "插图和附表清单": frozenset({"图", "表"}),
}

# 有"标签+内容"两段式结构、且识别层（recognize.py）已经把内容段一并标好
# `meta["slot"]="content"` 的部件——跟 `recognize._CONTENT_OWNERS` 是同一份
# 名单，这里独立放一份是因为 doc_index 不 import recognize（层次相反）。
CONTENT_OWNER_PARTS = frozenset({
    "摘要", "关键词", "英文摘要", "英文关键词",
    "引言", "结论", "参考文献", "附录", "致谢",
})
# 引言/结论 2026-09-16 起"当完整的标题"，永不补齐（不补标签也不补内容）。
NO_FILL_PARTS = frozenset({"引言", "结论"})


def auto_ops(index: list[Entry], structure_required: dict[str, bool], *,
            list_kinds: dict[str, set[str] | frozenset[str]] | None = None,
            insert_toc: bool = True) -> list[Op]:
    """识别索引（`Entry.part` 已经填好）+ 结构开关 → 指令表。

    这是**没有人工调整过结构**（命令行/自动路径）时 `build_ops()` 的等价物：
    两者产出的都是 `Op` 列表，执行器（`rebuild_body`/`rebuild_from`）不关心
    指令是从哪条路径算出来的。存在的部件按识别到的顺序照抄；启用但缺失的
    部件在 `FRONT_ORDER`/`BACK_ORDER` 规定的位置插一条生成指令；标签存在但
    内容为空（`CONTENT_OWNER_PARTS` 里那几个部件一条内容条目都没有）时，
    在标签后面补一条"只生成内容"的指令。

    `list_kinds`：三个清单部件要插哪几个域（`{"插图清单": {"图"}, ...}`），
    调用方按 `alerts.check_list_parts()` 算好传进来，跟 `build_ops()` 同义。
    """
    list_kinds = list_kinds or {}
    ordinals = paragraph_ordinals(index)
    structure_required = {k: v for k, v in structure_required.items()
                          if k not in NO_FILL_PARTS}

    by_part: dict[str, list[int]] = {}
    for ordinal, entry in ordinals.items():
        if entry.part:
            by_part.setdefault(entry.part, []).append(ordinal)
    for positions in by_part.values():
        positions.sort()

    # body 末尾那个 sectPr **只有在它真的排在最后**（`index[-1]` 就是它）时才
    # 交给 `rebuild_body` 特殊处理（原位保留在最后一条，好让后面追加的生成
    # 条目不会被排到它后面）；不在最后（`平时作业.docx` 实测在第 0 位，这是
    # WPS 存出来的常见形态）时，**不能不发 copy 指令**——以前这里不分青红皂白
    # 直接跳过所有 `kind=="sectPr"` 的条目，`rebuild_body` 那边也只认
    # `kids[-1]` 是不是它，两头一凑，这份文件的分节定义（含纸张尺寸）会被
    # 整个丢掉、且不会被任何指令补回来（2026-09-22 审计揪出的真 bug，实测过：
    # 成品里连 `w:pgSz` 都没有了）。跟 `build_ops()`（结构页手动路径，那边一直
    # 是对的）用同一条判据：`defer_sectpr`。
    last_entry = index[-1] if index else None
    defer_sectpr = bool(last_entry is not None and T_BODY_SECTPR in last_entry.traits)

    ops: list[Op] = []
    emitted: set[int] = set()
    # ⚠️ **缺失部件的生成指令不能就地发出**（2026-09-18 实测发现的真 bug）：
    # 如果直接在扫到"这个部件缺失"的那一刻就 append 生成指令，它会排在
    # "还没被任何 flush_before 扫过的真实内容"前面——比如"结论"缺失、
    # "参考文献"也缺失，而正文一~三章还没被任何东西 flush 过，生成的
    # "参考文献"占位就会跑到正文章节前面去。改成**先排队，等下一次真的
    # flush_before 时再倒出来**：那一刻真实内容已经被搬空了，排队的占位
    # 指令插在"上一段已搬的内容"和"这一段要搬的内容"之间，位置才对。
    pending: list[Op] = []

    def flush_before(limit_ordinal: int | None) -> None:
        limit_top = ordinals[limit_ordinal].top_index if limit_ordinal is not None else None
        for entry in index:
            if defer_sectpr and T_BODY_SECTPR in entry.traits:
                continue          # 真的排在最后：交给 rebuild_body 自己管
            if limit_top is not None and entry.top_index >= limit_top:
                break
            if entry.top_index in emitted:
                continue
            ops.append(Op("copy", entry=entry))
            emitted.add(entry.top_index)

    def drain_pending() -> None:
        ops.extend(pending)
        pending.clear()

    def add_fields(part: str, *, immediate: bool) -> None:
        """域指令（目录域／清单域）跟在标签后面。

        ⚠️ `immediate` 必须跟调用方一致（2026-09-19 修，`民总.docx` 实测
        抓到的真事故：目录标签和目录域隔着一整段正文，规范 [202] 要求
        "【目录】后另起一行编排目录内容"）。`emit_existing()` 里标签
        **已经真实存在**、这一刻刚被 `ops.append` 复制过去，域就该**立刻**
        跟上——排进 `pending` 会等到下一次 `drain_pending()`（下一个部件
        `flush_before` 之后）才倒出来，中间隔着的所有"还没搬过的正文"全部
        插进了标签和域之间。`emit_missing()` 里标签本身还不存在（这一刻还
        没有"标签在哪"这件事），域跟着标签一起排队、等真正落脚点出现时再
        一起倒出来，这条路径的排队没问题，不用改。
        """
        # ⚠️ 2026-09-19 同一轮修：`by_part.get("目录")` 这个条件删掉了——它原来
        # 会把"目录标签本身也缺失（走 `emit_missing`，这一刻 `by_part` 里还
        # 没有"目录"）"这个场景整个堵死，域永远造不出来（`build_part_paragraphs`
        # 的"目录"分支只造标签那一段，域完全靠这里单独一条指令）。要不要插
        # 域只取决于 `insert_toc`（＝源文件里没有真域，调用方已经算过）。
        target = ops if immediate else pending
        if part == "目录" and insert_toc:
            target.append(Op("generate", part=part, kind="toc_field"))
        for kind in ("图", "表"):          # 规范 [91]：分开列时插图清单在前
            if kind in (list_kinds.get(part) or ()):
                target.append(Op("generate", part=part, kind="list_field", text=kind))

    def emit_existing(part: str) -> None:
        # ⚠️ "有没有内容"要按**识别索引本身**判断，不能看"这一轮有没有轮到
        # 它被 emit"（2026-09-18 实测发现的真 bug，触发场景是当时"附录允许
        # 出现多次"的旧假设：某个部件出现多次、中间夹着另一个部件时，后者
        # 的内容会先被前者第二次出现时的 flush_before 当"未分类内容"顺路
        # 搬走——搬走不等于"没有"，用 emitted 集合判断会把已存在的内容误判
        # 成缺失，导致重复生成占位语）。2026-09-19 起后置标签全篇只认一次
        # （见 recognize.py::_scan_backward），这个具体触发场景已不再出现，
        # 但独立判断仍然更稳妥，保留。
        # 关键词的同段写法（`meta["inline"]`）：标签和关键词在同一段里，算有内容。
        has_content = any(ordinals[o].meta.get("slot") == "content"
                          or ordinals[o].meta.get("inline")
                          for o in by_part.get(part, ()))
        for ordinal in by_part.get(part, ()):
            entry = ordinals[ordinal]
            if entry.top_index in emitted:
                continue
            flush_before(ordinal)
            drain_pending()          # 排队的占位指令，落在"它前面"这一刻
            ops.append(Op("copy", entry=entry))
            emitted.add(entry.top_index)
        add_fields(part, immediate=True)
        # 标签在、正文空——补一条"只生成内容"的指令（用户 2026-09-16：
        # "所有正文也要补，按表格。不能靠有没有标签判断补不补正文"）。
        if (part in CONTENT_OWNER_PARTS and not has_content
                and structure_required.get(part, False)):
            pending.append(Op("generate", part=part, slot_col="text"))

    def emit_missing(part: str) -> None:
        pending.append(Op("generate", part=part, slot_col="label"))
        add_fields(part, immediate=False)

    for part in FRONT_ORDER:
        if part in by_part:
            emit_existing(part)
        elif part in LIST_PART_KINDS:
            if list_kinds.get(part):
                emit_missing(part)
        elif structure_required.get(part, False):
            emit_missing(part)

    # 前置段收尾：如果排到"引言"都没有落脚（引言本身也缺失），把排队的
    # 生成指令安在"正文第一个已识别标题"之前——跟旧锚链模型的
    # `body_start_index` 是同一个锚点。
    body_start = min(
        (ordinal for ordinal, entry in ordinals.items()
         if entry.part in ("一级标题", "二级标题", "三级标题", "四级标题")),
        default=None,
    )
    if body_start is None:
        # ⚠️ 全篇一个编号标题都没有（"引言"这一档本身也没识别到）时，以前
        # 这里什么都不做，排队的生成指令就一路留到 `flush_before(None)` +
        # 收尾那次 `drain_pending()` 才倒出来——那时候真实正文早就整段
        # `flush_before` 搬完了，摘要/目录这些缺失部件的占位语全部被排到了
        # 文末（2026-09-22 审计揪出的真 bug，实测过：小稿直接导出，占位语
        # 堆在最后）。退回"第一个没有被任何部件认领的段落"当锚点——不管
        # 标题用哪种编号方案，这一段肯定是"真正没人管的正文"，排在它之前
        # 就对了；一整篇全是已识别部件（没有游离正文）的极端情况仍然留给
        # 收尾处理，不勉强找锚点。
        body_start = min(
            (ordinal for ordinal, entry in ordinals.items() if not entry.part),
            default=None,
        )
    if body_start is not None:
        flush_before(body_start)
        drain_pending()

    for part in BACK_ORDER:
        if part in by_part:
            emit_existing(part)
        elif structure_required.get(part, False):
            emit_missing(part)

    flush_before(None)   # 收尾：body 顶层还剩下的（正文标题、普通正文、表格……）
    drain_pending()       # 后置段排队到最后还没地方落脚的，安在文档末尾
    for k, op in enumerate(ops):
        op.out_index = k
    return ops


def dump_ops(ops: list[Op], *, limit: int | None = None) -> str:
    """人看的指令表（给用户过目／调试）。"""
    kinds: dict[str, int] = {}
    for op in ops:
        kinds[op.action] = kinds.get(op.action, 0) + 1
    lines = [f"指令 {len(ops)} 条：" + "、".join(f"{k} {v}" for k, v in kinds.items())]
    lines.append("-" * 96)
    for op in ops[:limit]:
        lines.append("  " + op.label())
    if limit is not None and len(ops) > limit:
        lines.append(f"  ...（还有 {len(ops) - limit} 条）")
    return "\n".join(lines)


def rebuild_body(document, ops: list[Op] | None = None, *, builder=None,
                 track: dict[int, int] | None = None,
                 new_field_elements: list | None = None) -> list[Op]:
    """**核心执行器**：在一份已经打开（且已经 `unwrap_section_break_sdt` 过）
    的 `document` 上，清空正文、按 `ops` 顺序逐项往后追加。**就地修改传进来的
    `document`**，不重新读盘——这样调用方可以先在这份文档上做完清理/套样式，
    再用它当"索引法"的原料，不用先落盘再重新打开一次。

    `rebuild_from(path, ...)` 是这个函数配上"从磁盘读一份干净副本"的外壳，
    两者共用这一份逻辑，行为保证一致。

    `track`：传一个空 dict 进来，会被填成 `{源 Entry.top_index: 重建后的段落
    序号}`（只登记"照抄"来的段落，`document.paragraphs` 意义上的下标）——
    重建之后原来的段落下标全部作废，调用方要定位"引言现在是第几段"这类
    问题时靠这个映射表，不用重新识别一遍。

    `new_field_elements`：传一个空 list 进来，会被填上这次**新生成**的目录/
    清单域相关元素（`op.kind in ("toc_field", "list_field")` 的 generate 指令
    展开出来的那几个元素）。调用方（`pipeline.run`）在"用户勾了重新为文档
    生成"时，要把源文件里的旧域删掉——旧域和刚插的新域按同一套"像不像目录
    域"的判据扫描，不认身份，光靠 `keep` 参数保护"这些是新插的、不能删"
    （2026-09-22 审计修：以前调用方永远传 `keep=[]`，新旧域一起被删干净，
    见 `cleanup.drop_regenerated_fields`）。
    """
    body = document.element.body
    kids = list(body)
    # body 末尾那个 sectPr 是**整篇最后一节的分节定义**：它在最后时要么留在最后
    # （后面还要追加生成条目，不能让它变成倒数第二），要么就在 ops 顺序里原位保留。
    # ⚠️ 实测（2026-09-17，28 份夹具）：27 份都在最后，**只有 `平时作业.docx` 在
    # 第 0 位**（不是 Word 正常存出来的形态）。索引法的第一原则是**照抄**，不去
    # "纠正"用户文件——它不在最后时，我们的做法是**原位保留**，既不丢也不挪。
    sect_pr = kids[-1] if kids and kids[-1].tag == f"{W}sectPr" else None
    by_index = {i: child for i, child in enumerate(kids)}     # 跟 build_index 同一套下标
    index = build_index(document)
    ops = copy_all_ops(index) if ops is None else ops

    # ① 清空正文（保留末尾的 sectPr：它不属于任何段落，但也绝不能丢）
    for child in kids:
        if child is not sect_pr:
            body.remove(child)

    # ② 按 ops 顺序逐项往后追加
    paragraph_ordinal = 0
    for op in ops:
        if op.action == "copy":
            if op.entry is None:
                continue
            source = by_index.get(op.entry.top_index)
            if source is not None:
                body.append(source)                  # 搬过去（元素本身没被改过）
                if op.entry.kind == "paragraph":
                    if track is not None:
                        track[op.entry.top_index] = paragraph_ordinal
                    paragraph_ordinal += 1
            continue
        if op.action == "generate":
            if builder is not None:
                built = list(builder(op))
                for element in built:
                    body.append(element)
                    if element.tag == f"{W}p":
                        paragraph_ordinal += 1
                if new_field_elements is not None and op.kind in ("toc_field", "list_field"):
                    new_field_elements.extend(built)
                continue
            paragraph = etree.SubElement(body, f"{W}p")
            paragraph_ordinal += 1
            if op.text:
                run = etree.SubElement(paragraph, f"{W}r")
                node = etree.SubElement(run, f"{W}t")
                node.text = op.text
    if sect_pr is not None:
        body.append(sect_pr)                         # 收回最后一位
    return ops


def rebuild_from(path, ops: list[Op] | None = None, *, out_path=None,
                 builder=None) -> tuple:
    """按 `ops` 重建一份成品（**绝不动原文件**）。返回 `(文档, 实际用到的 ops)`。

    `ops=None` ⇒ 照抄全部条目（零失真基线）。
    `builder`：把一条 `generate` 指令**展开**成具体几段的回调
    （`builder(op) -> list[元素]`）。不传就只写 `op.text` 一段文字
    ——那是给"照抄/加法"这类测试用的最简形式；真要生成一个部件
    （标签 + 内容 + 域），由 `apply_structure.build_generated_elements()` 展开。
    """
    from docx import Document

    from essay_tool.engine.cleanup import unwrap_section_break_sdt

    document = Document(str(path))
    unwrap_section_break_sdt(document, keep_content=True)     # 跟建索引同一口径
    ops = rebuild_body(document, ops, builder=builder)
    if out_path is not None:
        document.save(str(out_path))
    return document, ops

