"""清除段落/文字上的直接格式，为套用我们自己的样式做准备。

官方模板大量样式没有真正定义格式，而是每段自己加直接格式（字体、字号、
段落间距……）。要让"套用样式"真正生效，必须先清掉这些直接格式，只留下
跟"角色"无关、用户自己敲的强调标记。**留哪些不是写死的**：默认
加粗/斜体/下划线/上下标，实际由预设「清理规则」小节决定（填"无"就是全清）；
正文标题例外——标题"样式挂帅"，手打的加粗/斜体一律清掉、完全听样式。
"""

from __future__ import annotations

import re

from docx.document import Document as DocumentObject
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run


def drop_regenerated_fields(document: DocumentObject, parts: set[str], keep=()) -> int:
    """删掉指定部件里**已有的**目录/清单域（用户在结构页勾了"重新为文档生成"）。

    用户 2026-09-15 的规则：插图和附表清单（三种）跟目录一样——已有域就"不动"，
    但允许用户主动勾选覆盖；勾了之后就把旧的删干净、按"没有"重新生成，
    不然成品里会同时留着旧域和工具新生成的那份。

    * **目录**：整个**域区间**一起删（`fldChar begin` 那一段到 `end` 那一段，
      中间域生成的条目也都在区间里），另外兜底删掉 `toc N` 样式的条目段；
      如果这一片还裹在内容控件（`w:sdt`）里，整个控件删掉。
    * **清单**：清单域就是单独一段，删它即可。

    `keep`：**不能删的段落对象**（工具刚生成的那份新域及其标签）——它们在
    "参数与旧域一模一样"的情况下也会被上面那两条规则命中。

    ⚠️ **只能在管线最后调**：删段落会让所有下标平移，而识别结果、用户手动
    核对过的结构全是按下标存的；提前删会把后面对不上号（实测：会直接抛
    IndexError）。所以调用方是先把新域插好、所有按下标套样式的活都干完，
    再回头删旧的。
    """
    from essay_tool.engine.doc_index import (  # 局部导入：避免模块级循环引用
        has_toc_field,
        is_toc_paragraph,
        list_kinds_in_instr,
        toc_field_part,
        toc_field_span_indices,
    )

    protected = {id(p) for p in keep} | {id(getattr(p, "_p", None)) for p in keep}
    paragraphs = list(document.paragraphs)

    # 判据一：域区间（跟段落，含域渲染出来的条目段——条目可能一个样式都不挂）
    marked = toc_field_span_indices(paragraphs, parts)
    # 判据二：兜底——单段就是域（`fldSimple` 那种，或者跨段落域的起止段），
    #         以及样式像目录条目的段落（`toc N`，手打/粘贴出来的目录）
    for index, paragraph in enumerate(paragraphs):
        if index in marked:
            continue
        if has_toc_field(paragraph) and toc_field_part(paragraph) in parts:
            marked.add(index)
        elif "目录" in parts and is_toc_paragraph(paragraph):
            marked.add(index)

    removed = 0
    # 先只"记账"：要删的段落/控件先算齐，**动结构之前**才来得及抢救分节符
    doomed_ps: list = []
    for index in sorted(marked):
        paragraph = paragraphs[index]
        if id(paragraph) in protected or id(paragraph._p) in protected:
            continue
        doomed_ps.append(paragraph._p)

    body = document._element.body
    doomed_sdts: list = []
    for sdt in list(body.findall(f".//{qn('w:sdt')}")):
        instrs = [el.text or "" for el in sdt.iter(qn("w:instrText"))]
        if not any(text.strip().upper().startswith("TOC") for text in instrs):
            continue
        if id(sdt) in protected:
            continue
        # 控件里的域是哪个部件：跟 `detect` **共用一份判据**（`\c "图"` 与
        # `\t "图题,1"` 两种写法都认）。**不能只认 `\c`**：我们自己写的清单域
        # 用的就是 `\t "图题,1"`，只认 `\c` 会把它判成"目录"——用户勾
        # "重新为文档生成目录"时会把这张清单域整个删掉（2026-09-15 审计抓到的
        # 同类残留，`detect.py` 已两套都认，这里以前没跟上）。
        kinds: set[str] = set()
        for text in instrs:
            kinds |= list_kinds_in_instr(text)
        if "图" in kinds and "表" in kinds:
            part = "插图和附表清单"
        elif "图" in kinds:
            part = "插图清单"
        elif "表" in kinds:
            part = "附表清单"
        else:
            part = "目录"
        if part in parts:
            doomed_sdts.append(sdt)

    # 控件是整块删的，里面的段落跟着一起消失，也算"活不下来"
    doomed = set(doomed_ps)
    for sdt in doomed_sdts:
        doomed.update(sdt.iter(qn("w:p")))

    # ⚠️ **删之前先抢救分节符**（见下面 `_rescue_section_breaks`）：分节符挂在
    # 段落标记上，段落一删整节就没了。
    _rescue_section_breaks(document, doomed_ps, doomed)

    for element in doomed_ps:
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)
            removed += 1
    for sdt in doomed_sdts:
        sdt.getparent().remove(sdt)
        removed += 1
    return removed


def _sect_pr(element):
    """段落元素上的分节符（`w:pPr/w:sectPr`）；没有就返回 None。"""
    pPr = element.find(qn("w:pPr"))
    return pPr.find(qn("w:sectPr")) if pPr is not None else None


def _rescue_section_breaks(document: DocumentObject, doomed_ps, doomed) -> int:
    """把"马上就要被删掉的段落"上的分节符，挪到还活着的段落上去。

    **用户 2026-09-17 报的严重 bug**：勾上结构页的「重新为文档生成」目录之后，
    成品从 2 节变 1 节——前置部分的罗马页码和两套页眉静默失效，还不报任何告警。
    根因：**分节符是挂在段落标记上的**（`w:pPr/w:sectPr`），而
    `sections.split_body_section()` 借的那个"正文首标题前面的段落"，恰好就是
    旧目录的最后一段（民总：段34，空段），管线最后删旧目录时把分节符一起删了。
    ⇒ 两件事的次序本来就是对的（先分节、最后删目录，早删会让下标平移），
    坏的是**载体撞在了一起**。

    修法：删之前把分节符挪到**它前面最近的、活得下来的段落**上——这样它仍然是
    "前一节的最后一段"，节边界一个字符都没动。两种情况需要补一个空段落专门
    承载：① 那个锚点自己也有分节符（源文件本来就在这儿分节，两道符不能挤在
    同一个段落的 pPr 里）；② 前面一段活的都不剩（整个前置部分都被删了）。

    返回挪动过的分节符个数。
    """
    moved = 0
    body = document._element.body
    # **从后往前处理**：补空段时是"插在它后面那个还活着的元素之前"，从后往前做，
    # 先补的那道符会自然落在后补的前面，节边界的先后顺序不会被打乱
    # （从前往后做会把顺序倒过来，实测过）。
    for element in reversed(list(doomed_ps)):
        sectPr = _sect_pr(element)
        if sectPr is None:
            continue
        anchor = element.getprevious()
        while anchor is not None and (anchor.tag != qn("w:p") or anchor in doomed):
            anchor = anchor.getprevious()
        if anchor is not None and _sect_pr(anchor) is None:
            anchor.set_sectPr(sectPr)
            moved += 1
            continue
        carrier = OxmlElement("w:p")
        carrier.set_sectPr(sectPr)
        if anchor is not None:
            anchor.addnext(carrier)         # 紧跟锚点，节边界不变
        else:
            following = element.getnext()
            while following is not None and (following.tag != qn("w:p")
                                             or following in doomed):
                following = following.getnext()
            trailing = body.find(qn("w:sectPr"))
            if following is not None:
                following.addprevious(carrier)
            elif trailing is not None:
                trailing.addprevious(carrier)   # body 末尾那个 sectPr 必须留在最后
            else:
                body.append(carrier)
        moved += 1
    return moved


def iter_tables(tables):
    """列出所有表格，**包括嵌在单元格里的嵌套表格**。

    为什么必须递归：`document.tables` 只给 body 直属的表格。套样式那边
    （`apply_structure.apply_table_text_styles`）一直走的是递归版，而清直接
    格式这边以前只走顶层 ⇒ 嵌套表“**只套样式、不清直接格式**”，用户自己敲的
    字体字号原样留在里面（2026-09-16 审计报告 Bug#3）。

    实现**只此一份**：`apply_structure` 直接 import 这个名字（那边原来有自己的
    `_iter_tables`），两边必须看到同一批表格。
    """
    for table in tables:
        yield table
        for row in table.rows:
            for cell in row.cells:
                yield from iter_tables(cell.tables)


def _sdt_holds_toc_field(sdt_content) -> bool:
    """内容控件里是不是装着目录/清单域。

    判据跟 `detect.has_toc_field` 一致：域指令住在 `w:instrText` 与
    `w:fldSimple/@w:instr` 里（`paragraph.text` 看不到它们）。这里手上是元素、
    不是 Paragraph，所以按元素扫，不另立一套口径。
    """
    codes = [el.text or "" for el in sdt_content.iter(qn("w:instrText"))]
    for el in sdt_content.iter(qn("w:fldSimple")):
        instr = el.get(qn("w:instr"))
        if instr:
            codes.append(instr)
    return any(code.strip().upper().startswith("TOC") for code in codes)


# 目录域的"收哪些标题级别"参数：`\o "1-1"`（不带引号、空格任意也认）。
_TOC_RANGE_RE = re.compile(r'\\o\s*"?\s*\d+\s*-\s*\d+\s*"?')


def normalize_toc_field_range(document: DocumentObject, levels: str = "1-3") -> int:
    r"""把**已有的目录域**的标题级别范围校正成 `1-3`。

    用户 2026-09-16 裁决（背景：他自己那份测试文件是 `TOC \o "1-1" \h \u `，
    Word 刷新后只收得到一级标题，二级三级条目全没进去）；规范 [90]："目录
    一般精确到三级标题"，[202] 也专门给了三级条目的缩进规则。

    **只改范围这一个参数**：`\h`/`\z`/`\u` 原样保留；清单域（指令里带
    `\c` 或 `\t` 的）一律不碰——`\t "图题,1"` 那种是按**样式名**收题注的，
    给它加 `\o` 会把它变成"按大纲级别收"，清单条目全错。

    只认**同一个元素里**写全的 `\o "x-y"`（`w:instrText` 或
    `w:fldSimple/@w:instr`）：域指令被 Word 拆到多个 `w:instrText` 元素里时
    不动它——同一段里往往还有别的域的指令（目录第一条的 HYPERLINK/
    PAGEREF），拼起来改写会把别人的指令洗掉。

    返回改了几个域。
    """
    from essay_tool.engine.doc_index import _TOC_INSTR_RE  # 局部导入：避免循环引用

    def fixed(code: str) -> str | None:
        if not _TOC_INSTR_RE.match(code):
            return None
        if "\\c" in code or "\\t" in code:
            return None                     # 清单域，别碰
        new = _TOC_RANGE_RE.sub(lambda _m: '\\o "%s"' % levels, code, count=1)
        return new if new != code else None

    changed = 0
    body = document._element.body
    for el in body.iter(qn("w:instrText")):
        if not el.text:
            continue
        new = fixed(el.text)
        if new is not None:
            el.text = new
            changed += 1
    for el in body.iter(qn("w:fldSimple")):
        instr = el.get(qn("w:instr"))
        if not instr:
            continue
        new = fixed(instr)
        if new is not None:
            el.set(qn("w:instr"), new)
            changed += 1
    return changed


def unwrap_section_break_sdt(document: DocumentObject, keep_content: bool = False) -> None:
    """处理"里面藏着分节符或目录域"的内容控件（w:sdt）。

    Word 自动生成目录时，经常把目录内容整个包进一个内容控件，分节符（从
    目录到正文那一刀）也跟着被包在里面——`python-docx` 的段落/分节遍历都
    是只看 body 的直接子元素，看不进内容控件内部，会把这一节漏算，等我们
    自己再插一次分节符，就会跟这个"看不见"的分节符叠在一起，变成三节。
    所以这种内容控件必须处理掉，让藏在里面的分节符回到 body 这一层。

    **目录域同样会被包在控件里**（用户 2026-09-16 实测的坑）：那种情况比
    分节符更难发现，因为域是"看得见"的（`detect.body_toc_field_parts()` 扫的
    是整篇 XML）而**标签看不见**（标签识别只看 `document.paragraphs`）——于是
    程序一边判"已有目录、不插新域"，一边找不到"目录"这两个字、又补了一个标签
    插到正文前面去。成品里就变成两个目录：控件里的旧目录留在原位，新标签
    跟着目录标题样式的段前分页跑到下一页。所以这类控件也得拆包。

    两种处理方式：

    * **删掉**（默认，`keep_content=False`）：里面装的是 Word 渲染出来的旧目录
      条目文字（超链接+页码域），而我们马上要插自己的 TOC 域重新生成目录，
      留着这些条目它们会变成散落在正文里的段落，所以连内容一起扔。
    * **拆包保留**（`keep_content=True`）：源文件里已经有**真正的目录域**时
      用它——用户 2026-09-15："如果用户已有 toc，则不动"。把内容控件换成它的
      子节点、顺序不变：旧目录条目留在原地（它们挂 `toc N` 样式，不会被当成
      标题），分节符也跟着回到 body 层、被 python-docx 看见。

    ⚠️ **装着目录/清单域的控件一律"拆包保留"**，不看 `keep_content`：标签必须
    回到 body 层才能被认出来（否成就多插一个标签），而旧域要不要删由用户
    勾的"重新为文档生成"决定（管线最后 `drop_regenerated_fields` 按段落删，
    拆包之后那两样都在 `document.paragraphs` 里，照样删得掉）。

    两样都不含的内容控件（比如 Zotero/EndNote 插入的引注）原样不动，不碰。

    ⚠️ **正文级内容控件也要拆**（AAT 2026-09-17 实测的"回炉不收敛"）：把
    body 里所有段落包进一个 `w:sdt` 之后，`document.paragraphs` 一段都看不
    见 ⇒ 识别不到题目/副标题 ⇒ 补件逻辑每遍再补一遍（四遍 128→130→135），
    而且产物里那个控件还留着。那种控件装的是论文正文，不是引注那种行内
    控件（行内 sdt 挂在 `w:p` 里面，不是 body 的直接子元素，不受影响）。
    所以只看**内容里有块级 `w:p`/`w:tbl` 的控件** ⇒ 一律拆包保留内容。
    """
    body = document._element.body
    for sdt in list(body.findall(f".//{qn('w:sdt')}")):
        sdt_content = sdt.find(qn("w:sdtContent"))
        if sdt_content is None:
            continue
        has_break = sdt_content.find(f".//{qn('w:sectPr')}") is not None
        has_toc = _sdt_holds_toc_field(sdt_content)
        # 块级内容（段落/表格）＝ 论文正文的一部分，必须让它回到 body 层，
        # 否则识别、套样式、清格式全都看不见它
        has_block = any(child.tag in (qn("w:p"), qn("w:tbl"))
                        for child in sdt_content)
        if not has_break and not has_toc and not has_block:
            continue
        parent = sdt.getparent()
        if not keep_content and not has_toc and not has_block:
            parent.remove(sdt)          # 只剩旧目录条目的那种：整块扔掉
            continue
        position = list(parent).index(sdt)
        children = list(sdt_content)
        parent.remove(sdt)
        for offset, child in enumerate(children):
            parent.insert(position + offset, child)


def set_run_text(run: Run, text: str) -> None:
    """只改 run 里的**文字**，不动它身上挂着的其它东西。

    **千万别用 `run.text = …`**：python-docx 那个 setter 会把 run 里除 `rPr`
    之外的子元素全部删掉——图形（`w:drawing`）、旧式形状（`w:pict`）、公式
    （`m:oMath`）、换行（`w:br`）、域（`w:fldChar`/`w:instrText`）、脚注引用……
    一起没了。实测过：一个同时装着 `w:drawing` 和 `m:oMath` 的 run，被
    `run.text = ""` 洗一下就两个都不剩了。

    这里只动 `w:t`：第一个改成新文字、多余的删掉，其余子元素原样保留；
    run 里原本没有 `w:t`（纯图形 run / 纯域 run）时才补一个。
    """
    texts = run._element.findall(qn("w:t"))
    if texts:
        texts[0].text = text
        # 前后带空格时必须显式声明保留空白，不然 XML 解析会把空格吃掉
        if text != text.strip():
            texts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        for extra in texts[1:]:
            run._element.remove(extra)
        return

    if text:
        element = OxmlElement("w:t")
        element.text = text
        # w:t 属于 EG_RunInnerContent（choice*），位置不固定；append 会落在
        # rPr 之后，正合 schema 要求。
        run._element.append(element)


# 默认保留的字符格式（`w:rPr` 的子元素名）。这四种属于"内容语义"——作者
# 自己敲的强调——清掉会改变论文的意思；真正的"排版格式"（字体、字号、颜色、
# 行距）才是要按预设覆盖的。**这一组可以用预设「清理规则」小节改
# （JSON/Excel 都认）**，
# 中文写法在 `config/excel_io.py` 的 `_KEEPABLE_RUN_FORMATS` 里。
DEFAULT_KEEP_RUN_FORMATS: frozenset[str] = frozenset({"b", "i", "u", "vertAlign"})

# `w:rStyle`（这个 run 引用了哪个字符样式）**永远保留**，不交给用户配：
# 它是引用不是格式，而且正文里的脚注引用标记就靠它指向内置的「脚注引用」样式
# （`vertAlign=superscript` 就住在那里）——删了标记就不上标了（用户实测发现）。
_ALWAYS_KEEP = frozenset({"rStyle"})

# 链接文字（超链接、网址）额外保留的字符格式（用户 2026-09-20，参考文献里）：
# 颜色与划线是链接自己的外观——`4√正文9.1.docx` 里的域式超链接是"引用『超链接』字符
# 样式（蓝色）＋直接格式覆盖成黑色、无划线（color=000000 u=none）"，清直接格式时黑色
# 被清掉、样式里的蓝色就露出来，成品里的链接莫名变蓝。`rStyle` 本来就永远保留。
LINK_KEEP_RUN_FORMATS: frozenset[str] = frozenset({"color", "u"})

_URL_LIKE = re.compile(r"(?i)(https?://|ftp://|www\.|doi:\s*10\.)")


def _run_text(run_element) -> str:
    return "".join(t.text or "" for t in run_element.findall(qn("w:t")))


def link_run_elements(paragraph: Paragraph) -> set:
    """这一段里"算链接"的 run（元素对象本身）：

    * `w:hyperlink` 里面的 run；
    * 域式超链接（`HYPERLINK` 域）的**结果文字**（`separate` 与 `end` 之间，嵌套域也认）；
    * 文字看起来是网址的 run（http(s)://、ftp://、www.、doi:10.）——没做成超链接、只是
      手动染了色的网址也算"链接"。
    """
    p = paragraph._p
    found: set = set()
    for hyper in p.findall(qn("w:hyperlink")):
        found.update(hyper.findall(qn("w:r")))
    stack: list[dict] = []          # 正在经过的域：{instr: 指令文字, result: 是否已到结果部分}
    for r in p.findall(qn("w:r")):
        fld = r.find(qn("w:fldChar"))
        if fld is not None:
            kind = fld.get(qn("w:fldCharType"))
            if kind == "begin":
                stack.append({"instr": "", "result": False})
            elif kind == "separate" and stack:
                stack[-1]["result"] = True
            elif kind == "end" and stack:
                stack.pop()
            continue
        instr = r.find(qn("w:instrText"))
        if instr is not None and stack and not stack[-1]["result"]:
            stack[-1]["instr"] += instr.text or ""
            continue
        if any(f["result"] and f["instr"].lstrip().upper().startswith("HYPERLINK")
               for f in stack):
            found.add(r)
        elif _URL_LIKE.search(_run_text(r)):
            found.add(r)
    return found


def _rsid_attributes(element) -> list[tuple[str, str]]:
    """挑出元素上的 rsid 类属性（`w:rsidR`、`w:rsidP`、`w:rsidRPr`……）。

    用户的初衷是"加工后的 docx 看不出被程序处理过"（见 `loader.py`），
    rsid 就是 Word 留下的编辑痕迹。清直接格式会把 `w:rPr`/`w:pPr` 整块删掉，
    挂在它们上面的 rsid 会跟着没——实测注入 58 个 `w:rsidP` + 44 个
    `w:rsidRPr`，走完管线一个不剩。搬到新容器上之后，成品里会出现"只有
    rsid 属性的空 rPr/pPr"——**那正是 Word 自己也会产生的形态**，不影响渲染，
    也不会让人看出格式被程序重排过。
    """
    return [
        (name, value)
        for name, value in element.attrib.items()
        if name.split("}")[-1].startswith("rsid")
    ]


def strip_run_direct_formatting(
    run: Run,
    keep: frozenset[str] | None = None,
) -> None:
    """清掉 run 上的直接格式，只留下 `keep` 里点名的那些标签。

    **按标签白名单做，而不是逐个调 python-docx 属性**：`w:rPr` 里能挂几十种
    子元素（加粗、颜色、着重号 `w:em`、字符间距 `w:spacing`……），其中相当
    一部分 python-docx 根本没有对应属性（比如着重号），靠属性逐个搬运迟早漏。
    保留下来时**按原顺序**装回去，`w:rPr` 内部的 schema 顺序自动保持正确。
    """
    rPr = run._element.find(qn("w:rPr"))
    if rPr is None:
        return

    names = (DEFAULT_KEEP_RUN_FORMATS if keep is None else keep) | _ALWAYS_KEEP
    kept = [child for child in rPr if child.tag.split("}")[-1] in names]
    rsids = _rsid_attributes(rPr)
    run._element.remove(rPr)
    if not kept and not rsids:
        return
    new_rPr = run._element.get_or_add_rPr()
    # 属性先设、子元素后添：这样 rStyle 等仍然是子元素序列里的第一个
    for name, value in rsids:
        new_rPr.set(name, value)
    for child in kept:
        new_rPr.append(child)


def strip_paragraph_direct_formatting(
    paragraph: Paragraph,
    keep: frozenset[str] | None = None,
    keep_links: bool = False,
) -> None:
    """`keep_links=True`：链接文字（见 `link_run_elements`）在 `keep` 之外再留下
    颜色与划线（`LINK_KEEP_RUN_FORMATS`），字体字号等其余直接格式照常清掉。此时
    `w:hyperlink` 里的 run 也一并清理——`paragraph.runs` 不含它们，不然它们会带着
    原来的字号字体留在成品里、跟正文对不上。"""
    if keep_links:
        links = link_run_elements(paragraph)
        link_keep = (DEFAULT_KEEP_RUN_FORMATS if keep is None else keep) | LINK_KEEP_RUN_FORMATS
        for run in paragraph.runs:
            strip_run_direct_formatting(run, link_keep if run._element in links else keep)
        for hyper in paragraph._p.findall(qn("w:hyperlink")):
            for r in hyper.findall(qn("w:r")):
                strip_run_direct_formatting(Run(r, paragraph), link_keep)
    else:
        for run in paragraph.runs:
            strip_run_direct_formatting(run, keep)

    pPr = paragraph._p.find(qn("w:pPr"))
    if pPr is None:
        return

    # 自动编号引用（w:numPr）**不是直接格式**——它说明"这一段属于哪个编号
    # 列表"。整块 pPr 删掉的话，文档里所有用 Word 自动编号的列表（参考文献
    # 常常这么写）会在成品里集体消失。所以单独把它留下来，其余（字体、间距、
    # 缩进、大纲级别……）照旧清掉。
    # 顺序要紧：标题的自动编号已经在 normalize.py 里"固化"成真文字、并把
    # numPr 摘掉了，这里保住的都是**没固化**的那些（正文里的编号列表）。
    #
    # `w:pStyle`（这一段挂了哪个段落样式）同样**不是直接格式**，一样要留下
    # （跟 run 上的 `w:rStyle` 一个道理，见 `_ALWAYS_KEEP`）。整块删掉会连
    # 样式引用一起抹掉：目录/清单域生成的条目（`toc 1/2/3` 样式）会因此掉成
    # 不带样式的正文——用户 2026-09-15 要求"已有 toc 不动"，写作模板.docx
    # 实拍：14 行旧目录条目的样式引用全没了。正常流程里正文/标题的样式是
    # 后面重新套上去的，不受影响。
    numPr = pPr.find(qn("w:numPr"))
    pStyle = pPr.find(qn("w:pStyle"))
    # `w:sectPr`（分节符）同样**不是直接格式**——它是段落用来标记"这里断一节"
    # 的结构标记，跟这一段的字体/间距毫不相干。整块 pPr 删掉的话，挂在正文
    # 中段的分节符（比如一页横向表格用的那一节）会被静默删掉、也没有任何
    # 指令再把它补回来，其后的所有内容会被并进前一节，页面设置跟着错乱
    # （2026-09-22 审计揪出的真 bug，单独调用本函数实测复现：一清格式，2 节
    # 直接变 1 节）。工具自己在"引言"处重建的前置/正文分节点是后来才加的，
    # 跟这里要保护的"源文档原有的分节符"是两回事，互不影响。
    sectPr = pPr.find(qn("w:sectPr"))
    rsids = _rsid_attributes(pPr)
    paragraph._p.remove(pPr)
    if numPr is None and pStyle is None and sectPr is None and not rsids:
        return
    # 用 get_or_add_pPr() 让 python-docx 把新 pPr 插到 w:p 的正确位置
    # （pPr 必须是第一个子元素），别自己 insert(0, …)。子元素顺序按 schema：
    # pStyle 在最前、numPr 其次、sectPr 排在 pPr 的最后（CT_PPr 里 rPr 之后
    # 只剩它和 pPrChange），先 append pStyle、再 numPr、最后 sectPr 正好。
    kept = paragraph._p.get_or_add_pPr()
    for name, value in rsids:
        kept.set(name, value)
    if pStyle is not None:
        kept.append(pStyle)
    if numPr is not None:
        kept.append(numPr)
    if sectPr is not None:
        kept.append(sectPr)


# ---------------------------------------------------------------- 清除手打分页符 / 空行
# （用户 2026-09-20，预设「清理规则」：`clear_page_breaks` / `clear_blank_lines`。
#  **都不作用于封面范围**，调用方把封面的段落下标/表格/顶层位置传进来避开。）


def strip_page_breaks(
    document: DocumentObject,
    skip_indices: set[int] | None = None,
    skip_tables: "set | None" = None,
) -> int:
    """清掉正文里**所有**手打分页符（`w:br w:type="page"`，即 Ctrl+Enter），返回清了几个。

    只清"分页符"本身：分节符（`w:sectPr`，含"下一页"型）、换行（Shift+Enter）、
    列分隔符都不动；`w:lastRenderedPageBreak`（Word 排版时留的"这里曾经换页"标记，
    不是真分页符）顺手清掉。分页符所在的 run 清空后整个 run 一并删掉。段落里原来
    "段前分页"的**段落属性**本来就在清直接格式时清掉了，不归这里管。
    """
    skip = skip_indices or set()
    table_skip = skip_tables or set()
    removed = 0

    def clean(paragraph: Paragraph) -> int:
        n = 0
        for marker in list(paragraph._p.iter(qn("w:lastRenderedPageBreak"))):
            marker.getparent().remove(marker)
        for br in list(paragraph._p.iter(qn("w:br"))):
            if br.get(qn("w:type")) != "page":
                continue
            run = br.getparent()
            run.remove(br)
            n += 1
            if run.tag == qn("w:r") and not [c for c in run if c.tag != qn("w:rPr")]:
                run.getparent().remove(run)          # 只剩 rPr 的空 run 不留
        return n

    for index, paragraph in enumerate(document.paragraphs):
        if index not in skip:
            removed += clean(paragraph)
    for table in iter_tables(document.tables):
        if table._tbl in table_skip:
            continue
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    removed += clean(paragraph)
    return removed


# 出现在段落里就说明它"不是空行"的元素（图片/对象/域/公式/脚注引用/书签/修订……）
_NOT_BLANK_TAGS = frozenset(qn(t) for t in (
    "w:drawing", "w:pict", "w:object", "w:fldChar", "w:fldSimple", "w:instrText",
    "w:sdt", "w:footnoteReference", "w:endnoteReference", "w:commentReference",
    "w:sym", "w:ptab", "w:tab", "w:noBreakHyphen", "w:softHyphen",
    "w:bookmarkStart", "w:ins", "w:del", "w:moveFrom", "w:moveTo",
    "m:oMath", "m:oMathPara",
)) | {"{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent"}


def is_blank_paragraph(p, *, page_breaks_removed: bool = False) -> bool:
    """这个 `w:p` 是不是**确实为空**的空行。

    严格：段落里**一个字符都没有**（连空格、全角空格、制表符都算"有"），也没有图片/
    对象/域/公式/脚注引用/书签/修订，没有分节符（`pPr/sectPr`）、没有自动编号
    （`pPr/numPr`——空编号项显示成一个孤零零的序号，不是空行）。换行（Shift+Enter）
    也算内容，只有手打**分页符**在 `page_breaks_removed=True`（本次会清掉它）时不算。
    """
    if p.tag != qn("w:p"):
        return False
    pPr = p.find(qn("w:pPr"))
    if pPr is not None and (pPr.find(qn("w:sectPr")) is not None
                            or pPr.find(qn("w:numPr")) is not None):
        return False
    for el in p.iter():
        if el.tag == qn("w:t"):
            if el.text:
                return False
        elif el.tag == qn("w:br"):
            if not (page_breaks_removed and el.get(qn("w:type")) == "page"):
                return False
        elif el.tag in _NOT_BLANK_TAGS:
            return False
    return True


def cover_top_limit(document: DocumentObject, cover_end: int | None) -> int | None:
    """封面在 body 顶层里占到哪儿：顶层下标**小于**这个数的都是封面（含封面里的表格）。
    没设封面范围返回 None。"""
    if cover_end is None:
        return None
    paragraphs = document.paragraphs
    kids = list(document._element.body.iterchildren())
    if 0 <= cover_end + 1 < len(paragraphs):
        target = paragraphs[cover_end + 1]._p
        for i, child in enumerate(kids):
            if child is target:
                return i
    return len(kids)


def drop_blank_paragraph_ops(ops: list, document: DocumentObject, cover_limit: int | None,
                             *, page_breaks_removed: bool = False) -> tuple[list, int]:
    """从指令表里去掉"照抄一个空行"的指令（重建时那个元素就不会被搬过去），返回
    `(新指令表, 去掉几条)`。封面范围（`cover_limit`）里的不动；生成的占位语/域不动。

    两处要留一个空段，否则文档结构会变：**两张表格之间**（Word 会把紧挨着的两张表
    合成一张）、**文档末尾的表格之后**（正文必须以段落收尾）。
    """
    kids = list(document._element.body.iterchildren())
    by_index = dict(enumerate(kids))                 # 跟 doc_index.rebuild_body 同一套下标

    def element_of(op):
        if op.action != "copy" or op.entry is None:
            return None
        return by_index.get(op.entry.top_index)

    def removable(op) -> bool:
        el = element_of(op)
        if el is None or op.entry.kind != "paragraph":
            return False
        if cover_limit is not None and op.entry.top_index < cover_limit:
            return False
        return is_blank_paragraph(el, page_breaks_removed=page_breaks_removed)

    flags = [removable(op) for op in ops]

    def is_sect(op) -> bool:
        el = element_of(op)
        return el is not None and el.tag == qn("w:sectPr")

    def tag_after(o) -> str:
        """这条指令产出的第一个元素的标签。生成指令（占位语/标签/域）产出的都是段落。"""
        if o.action != "copy":
            return qn("w:p")
        el = element_of(o)
        return el.tag if el is not None else qn("w:p")

    kept: list = []
    dropped = 0
    for i, op in enumerate(ops):
        if not flags[i]:
            kept.append(op)
            continue
        prev = element_of(kept[-1]) if kept else None
        nxt_op = next((o for j, o in enumerate(ops[i + 1:], start=i + 1)
                       if not flags[j] and not is_sect(o)), None)
        nxt_tag = tag_after(nxt_op) if nxt_op is not None else None
        prev_is_table = prev is not None and prev.tag == qn("w:tbl")
        if prev_is_table and (nxt_tag is None or nxt_tag == qn("w:tbl")):
            kept.append(op)                           # 表格之间／文末表格之后：留一个
            continue
        dropped += 1
    return kept, dropped


def tables_in_prefix(document: DocumentObject, end_index: int) -> set:
    """body 顺序里排在 `document.paragraphs[end_index]` **之前**的表格（**元素对象本身**）。

    表格不在 `document.paragraphs` 里，所以"前 N 段"这类范围判断（目前只有
    "文档开头是封面，那一段以后别动"）对表格得另外算一次 body 顺序。
    `end_index` 超出段落总数＝"整篇所有的表格"（封面盖到文档末尾的情形）。

    **必须返回元素对象、不能返回 `id(元素)`**（2026-09-14 实拍修正）：lxml 的
    元素代理对象是用完就回收的，下一次 `.tbl_lst` 会造一个**新对象**、`id()` 就
    变了——实测连"刚算完立刻查"都对不上，于是所有 `skip_tables=` 判断
    （封面的表格不清理/不套字/不加标题行）和"删封面时连表格一起删"全部静默
    失效。把元素对象本身放进集合里、由集合**握着引用**，lxml 就会复用同一个
    代理，调用方的 `table._tbl in skip_tables` 才真的成立。
    """
    body = document._element.body
    paragraphs = document.paragraphs
    target = paragraphs[end_index]._p if 0 <= end_index < len(paragraphs) else None
    found: set = set()
    for child in body.iterchildren():
        if child is target:
            break
        if child.tag == qn("w:tbl"):
            found.add(child)
    return found


def strip_document_direct_formatting(
    document: DocumentObject,
    keep: frozenset[str] | None = None,
    style_only_indices: set[int] | None = None,
    skip_indices: set[int] | None = None,
    skip_tables: "set | None" = None,
    link_indices: set[int] | None = None,
) -> None:
    """整篇清一遍。`keep` 是要保留的字符格式标签，None 表示用默认值。

    `link_indices`：这些段落里的**链接文字**额外保留颜色与划线（参考文献，用户
    2026-09-20），见 `strip_paragraph_direct_formatting(keep_links=)`。

    `style_only_indices`：
    **样式挂帅**的段落（正文标题，由调用方传入它们的段落下标）——这些段落
    的格式完全由样式决定，**连加粗/斜体/下划线也不保留**：作者在标题上
    手打的加粗会以 run 级直接格式盖过样式（四级标题样式本来规定不加粗，
    `w:b` 一留它就变成加粗；2026-09-13 用户实拍"错误的也继承了标题上的
    加粗格式、原有标题的不当加粗不当保留"）。内容类段落（正文、摘要正文、
    参考文献条目……）不吃这个参数，照旧按白名单保留作者强调。

    `skip_indices` / `skip_tables`：**整段跳过不碰**的范围——用户指定的
    封面（2026-09-15："不作任何程度的封面自动识别，纯靠用户手敲"）。
    那一片是用户自己排的版，工具既不识别也不改它：清直接格式必须跳开，否则
    封面会被清成一堆光秃秃的宋体四号。`skip_tables` 装的是
    `tables_in_prefix()` 算出来的**表格元素对象本身**（不是 id，原因见那个
    函数）。
    """
    style_only = style_only_indices or set()
    links = link_indices or set()
    skip = skip_indices or set()
    table_skip = skip_tables or set()
    for index, paragraph in enumerate(document.paragraphs):
        if index in skip:
            continue
        if index in style_only:
            strip_paragraph_direct_formatting(paragraph, frozenset())
        else:
            strip_paragraph_direct_formatting(paragraph, keep, keep_links=index in links)
    for table in iter_tables(document.tables):
        if table._tbl in table_skip:
            continue
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    strip_paragraph_direct_formatting(paragraph, keep)
