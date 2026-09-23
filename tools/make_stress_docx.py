# -*- coding: utf-8 -*-
"""生成《压力测试》docx 夹具（只新建文件，不修改任何既有文件）。

设计原则（用户 2026-09-15 要求）：**夹具本身要人能读懂**——每段的测试目的
直接写在段落文字里，打开 Word 就能逐段看过去。

两种标记约定（两类样本，测的东西不一样）：
  * 【Zxx·目的：…】开头的段落 —— 它自己就是被测样本。测试目的是前缀，后面的
    文字就是样本内容；这类段落用来测「位置 / 是否被复制 / 是否崩」。
  * zmarkX 结尾的段落 —— 精确样本。这类段落的**开头字符或结尾标点本身就是
    被测对象**（比如「第一章」式非规范编号必须出现在段首），所以不能加前缀；
    改用一个唯一的英文标记 zmarkX，方便脚本在产物里定位它。这类段落前面一定
    有一行以 ◆ 开头的说明段，告诉读者它在测什么。

产物：
  reference/压力测试_Agent版.docx   ← 夹具本体（人可读）
  work/_stress_manifest.json        ← 脚本用：标记表、封面范围候选、段落数
  work/_压力测试_清单.txt            ← 人用：段号 + 目的 + 文字速览

用法：.venv\\Scripts\\python.exe tools\\make_stress_docx.py
"""

from __future__ import annotations

import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from docx import Document  # noqa: E402
from docx.enum.style import WD_STYLE_TYPE  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Cm  # noqa: E402

OUT_DOCX = ROOT / "reference" / "压力测试_Agent版.docx"
OUT_MANIFEST = ROOT / "work" / "_stress_manifest.json"
OUT_HUMAN = ROOT / "work" / "_压力测试_清单.txt"
IMAGE = ROOT / "essay_tool" / "gui" / "assets" / "cover_thesis.png"

doc = Document()
rows: list[dict] = []          # 给人看的清单
marks: dict[str, str] = {}     # 标记 -> 该段目的（脚本用）
cover_candidates: list[int] = []


def add(text: str = "", purpose: str = "", style=None, mark: str = "",
        cover: bool = False):
    """加一段。purpose 只进清单文件、不污染段落文字；mark 会写进段落末尾。"""
    if mark:
        text = f"{text} {mark}"
        marks[mark] = purpose
    paragraph = doc.add_paragraph(text, style=style) if style else doc.add_paragraph(text)
    index = len(doc.paragraphs) - 1
    rows.append({"index": index, "purpose": purpose, "text": text[:58]})
    if cover:
        cover_candidates.append(index)
    return paragraph


def add_field(paragraph, instr: str, cached: str) -> None:
    """往段落里塞一个域（fldSimple 最简形态）。"""
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), instr)
    run = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = cached
    run.append(t)
    fld.append(run)
    paragraph._p.append(fld)


def add_revision_paragraph() -> None:
    """未接受的修订：w:ins（插入）+ w:del（删除）——python-docx 看不见其中文字。"""
    paragraph = doc.add_paragraph()
    for tag, text, rid in (("w:ins", "这段文字是未接受的插入内容 zmarkR1", "1"),
                           ("w:del", "这段是被删除的内容 zmarkR2", "2")):
        el = OxmlElement(tag)
        el.set(qn("w:id"), rid)
        el.set(qn("w:author"), "Agent")
        el.set(qn("w:date"), "2026-09-15T00:00:00Z")
        run = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = text
        run.append(t)
        el.append(run)
        paragraph._p.append(el)
    rows.append({"index": len(doc.paragraphs) - 1,
                 "purpose": "未接受的修订（w:ins / w:del）",
                 "text": "（修订里的文字不在 paragraph.text 里）"})
    marks["zmarkR1"] = "未接受的插入（w:ins）里的标记"
    marks["zmarkR2"] = "被删除的内容（w:del）里的标记"


def add_table(rows_n: int, cols_n: int, cells: dict, purpose: str, text_summary: str = ""):
    table = doc.add_table(rows=rows_n, cols=cols_n)
    table.style = "Table Grid"
    for (r, c), value in cells.items():
        table.cell(r, c).text = value
    rows.append({"index": None, "purpose": purpose,
                 "text": text_summary or "表格"})
    return table


def ensure_style(name: str, fallback: str):
    try:
        return doc.styles[name]
    except KeyError:
        return doc.styles.add_style(fallback, WD_STYLE_TYPE.PARAGRAPH)


def _numbering_element(doc):
    """拿到（必要时手动新建）numbering.xml 的根元素。"""
    try:
        return doc.part.numbering_part.element
    except (KeyError, NotImplementedError):
        pass
    from docx.opc.constants import CONTENT_TYPE as CT
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.opc.packuri import PackURI
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    from docx.parts.numbering import NumberingPart
    element = parse_xml(f'<w:numbering {nsdecls("w")}/>')
    part = NumberingPart(PackURI("/word/numbering.xml"), CT.WML_NUMBERING, element,
                         doc.part.package)
    doc.part.relate_to(part, RT.NUMBERING)
    return element


def build_auto_numbering(doc) -> int:
    """往 numbering.xml 里加一个两级中文数字列表定义，返回 numId。

    这是为了真正测到「Word 自动编号」那条路：编号只存在于 numbering.xml，
    段落文字里什么都没有（要套 numPr 才会画出来）。
    """
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    numbering = _numbering_element(doc)
    used: set[int] = set()
    for tag, attr in ((qn("w:abstractNum"), qn("w:abstractNumId")),
                      (qn("w:num"), qn("w:numId"))):
        for el in numbering.findall(tag):
            value = el.get(attr)
            if value and value.isdigit():
                used.add(int(value))
    new_id = max(used, default=0) + 1
    abstract = parse_xml(
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{new_id}">'
        '<w:multiLevelType w:val="multilevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/>'
        '<w:numFmt w:val="chineseCounting"/><w:lvlText w:val="%1、"/>'
        '<w:lvlJc w:val="left"/></w:lvl>'
        '<w:lvl w:ilvl="1"><w:start w:val="1"/>'
        '<w:numFmt w:val="chineseCounting"/><w:lvlText w:val="（%2）"/>'
        '<w:lvlJc w:val="left"/></w:lvl>'
        '</w:abstractNum>')
    first_num = numbering.find(qn("w:num"))
    if first_num is None:
        numbering.append(abstract)
    else:
        first_num.addprevious(abstract)
    numbering.append(parse_xml(
        f'<w:num {nsdecls("w")} w:numId="{new_id}">'
        f'<w:abstractNumId w:val="{new_id}"/></w:num>'))
    return new_id


def set_numpr(paragraph, num_id: int, ilvl: int) -> None:
    """给段落挂上「属于哪个自动编号列表」——Word 就是靠这个画编号的。"""
    pPr = paragraph._p.get_or_add_pPr()
    numPr = OxmlElement("w:numPr")
    il = OxmlElement("w:ilvl")
    il.set(qn("w:val"), str(ilvl))
    ni = OxmlElement("w:numId")
    ni.set(qn("w:val"), str(num_id))
    numPr.append(il)
    numPr.append(ni)
    pPr.insert(0, numPr)


# ======================================================== 阅读约定 + 封面区样本
add("本文件是排版工具的压力测试样本，不是真实论文，请勿据此判断论文格式规范。",
    purpose="阅读约定①（句号结尾的普通正文）", cover=True)
add("◆ 阅读约定：带【Z…·目的】的段落，它自己就是被测样本（括号里是目的、后面是样本"
    "内容）；带 ◆ 的段落用来解释它下面那一段；带 zmark 加字母的是精确样本——它的"
    "开头字符或结尾标点本身就是被测对象，所以不能加前缀。",
    purpose="阅读约定②（说明段）", cover=True)
add("◆ 题目之前放了三种不该出现的内容：栏位行、下划线行、来源不明的一段正文。"
    "测：工具会不会把它们抢成论文题目。",
    purpose="说明", cover=True)
add("课程名称：知识产权法",
    purpose="封面区样本①：栏位行（无标点结尾，可能被抢成题目）", cover=True)
add("论文编号：____________",
    purpose="封面区样本②：下划线栏位行", cover=True)
add("【Z01·目的：题目之前的普通正文。测它会不会被吞掉、错位，或在结构页里看不见】"
    "这一段写在真题目之前，是再普通不过的正文。",
    purpose="封面区样本③：题目前的普通正文", cover=True)
print("— 封面区表格 —")
add_table(2, 2, {(0, 0): "封面表格样本", (0, 1): "zmarkT1",
                 (1, 0): "第二行", (1, 1): "zmarkT2"},
          purpose="封面区表格（表格不是段落，无段号）",
          text_summary="2×2 表格，含 zmarkT1 / zmarkT2")
marks["zmarkT1"] = "封面区表格里的标记（删除封面时它必须一起消失）"
marks["zmarkT2"] = "封面区表格里的标记 2"

# ======================================================== 题目 / 副标题
add("论网络平台侵权责任的边界",
    purpose="★真题目（没有编号、不以标点结尾）", cover=True)
add("◆ 下面一段测：紧跟题目之后、不以标点结尾的那段，会不会被当成副标题。",
    purpose="说明", cover=True)
add("张三　指导教师：李四　2026 年 9 月",
    purpose="副标题样本：作者信息行", mark="zmarkA", cover=True)

# ======================================================== 中文前置部件
add("【摘要】", purpose="摘要标签（精确样本，无前缀）")
add("【Z02·目的：摘要正文里出现图 1 字样，测它会不会被误判成图题】"
    "本文讨论平台侵权责任的边界，图 1 是本文的分析框架示意，全文共分三部分。",
    purpose="摘要正文 + 图题误判样本")
add("【关键词】", purpose="关键词标签（精确样本）")
add("关键词A；关键词B；关键词C", purpose="关键词内容", mark="zmarkB")

# ======================================================== 英文页
add("◆ 下面两段是英文题目 + 英文副标题（连着的两行），"
    "测会不会把副标题行当成英文题目。",
    purpose="说明")
add("RESEARCH ON THE BOUNDARY OF PLATFORM LIABILITY IN TORT LAW",
    purpose="英文题目（精确样本）", mark="zmarkC")
add("A COMPARATIVE STUDY OF CHINA AND THE EUROPEAN UNION",
    purpose="英文副标题（精确样本）", mark="zmarkD")
add("【ABSTRACT】", purpose="英文摘要标签（精确样本）")
add("【Z03·目的：英文摘要里出现 Table 1 字样，测它会不会被误判成表题】"
    "This paper studies Table 1 and the related rules in detail.",
    purpose="英文摘要 + 表题误判样本")
add("【KEYWORDS】", purpose="英文关键词标签（精确样本）")
add("platform liability; tort law; safe harbor",
    purpose="英文关键词内容", mark="zmarkE")

# ======================================================== 目录 / 清单页
add("【目录】", purpose="目录标签（精确样本）")
add("论网络平台侵权责任的边界……1", purpose="目录页样本①：手打假条目（不是真域）")
add("一、平台责任的一般理论……3", purpose="目录页样本②：手打假条目")
add("【Z04·目的：目录标签下面写了正文，测多余的正文告警】"
    "这段话粘在目录页里，本来不该出现在这里。",
    purpose="目录页样本③：标签下的正文（应告警）")
add("【Z05·目的：目录页第二段正文，测告警位置能否给出多个段号】"
    "这是目录页里第二段多余正文。",
    purpose="目录页样本④：多余正文第二段")
add("◆ 下面一段是真的目录域（TOC，脚本手写的域代码），"
    "测：已有真目录域时工具应当不动它、不插第二份。",
    purpose="说明")
toc_para = add("", purpose="真目录域（TOC \\o 1-3）")
add_field(toc_para, 'TOC \\o "1-3" \\h \\z \\u', "（真目录域）")
toc_style = ensure_style("TOC 1", "TOC 1")
add("论网络平台侵权责任的边界\t1",
    purpose="真目录域生成的条目①（toc 1 样式，测不被套成正文）", style=toc_style)
add("一、平台责任的一般理论\t3",
    purpose="真目录域生成的条目②", style=toc_style)

add("【插图清单】", purpose="插图清单标签（精确样本）")
add("◆ 下面一段是工具自己会写的那种清单域（TOC 带 \\t 按样式名收集），"
    "测：探测端认不认得出它（2026-09-15 修过的那条）。",
    purpose="说明")
list_para = add("", purpose="真清单域（TOC \\h \\z \\t 图题,1）")
add_field(list_para, 'TOC \\h \\z \\t "图题,1"', "（真清单域）")
add("【Z06·目的：清单标签下面写了正文，测多余的正文告警】"
    "清单页里也粘了一段正文。",
    purpose="清单页样本：标签下的正文（应告警）")
add("【附表清单】",
    purpose="附表清单标签（精确样本；本文档没有表题，测不存在的清单对象告警）")

# ======================================================== 正文区
add("引　言", purpose="锚点部件：引言（一级）")
add("【Z07·目的：引言标题下、第一个一级标题之前插入的游离正文①。"
    "测导出后它会不会被复制成多段、或被移出原位】游离正文①：这段不该被复制。",
    purpose="游离正文①（复制 / 错位重点样本）")
add("一、平台责任的一般理论", purpose="一级标题")
add("【Z08·目的：一级标题下面直接跟三级标题（跳级），测会不会误报】跳级是允许的。",
    purpose="正文（跳级场景）")
add("1. 归责原则", purpose="三级标题（跳级出现）")
add("◆ 下面一段是孤立列举样本：这一级下面没有第二个同级标题。", purpose="说明")
add("（1） 孤立列举样本", purpose="孤立列举样本（精确样本）", mark="zmarkF")
add("【Z09·目的：两条一级标题之间插入游离正文②，测复制 / 错位】游离正文②。",
    purpose="游离正文②")
add("二、平台责任的具体类型", purpose="一级标题")
add("（一） 内容审核责任", purpose="二级标题")
add("◆ 下面一段测未归零：新的二级标题下，第一条三级标题的编号不是 1。", purpose="说明")
add("3. 审核义务的边界", purpose="未归零样本（精确样本）", mark="zmarkG")
add("三、平台责任的例外", purpose="一级标题")
add("（二） 避风港规则", purpose="未归零 / 跳号样本（精确样本）", mark="zmarkH")
add("◆ 下面一段测跳号：前面到了三，这里直接写六。", purpose="说明")
add("六、跳号样本", purpose="跳号样本（精确样本）", mark="zmarkI")
add("◆ 下面三段用的是规范之外的编号写法，测会不会被报成疑似标题。", purpose="说明")
add("第一章 平台责任", purpose="非规范编号①：第一章式（精确样本）", mark="zmarkJ")
add("① 带圈数字样本", purpose="非规范编号②：带圈数字（精确样本）", mark="zmarkK")
add("1) 数字加右括号样本", purpose="非规范编号③：数字加右括号（精确样本）", mark="zmarkL")
add("◆ 下面一段测以标点结尾的标题：它很像标题，但句末有逗号，应当静默归正文。",
    purpose="说明")
add("一、这个标题以逗号结尾，应当是正文，", purpose="标点结尾的伪标题（精确样本）")

# ======================================================== 图表与题注
add("图 1 手打编号的中文图题",
    purpose="中文图题①（手打编号原样保留：挂图题样式 ⇒ 进插图清单；我们不动编号）",
    mark="zmarkM")
add("Fig 1 English caption of the same figure",
    purpose="英文对照图题（不挂样式 ⇒ 不进清单；文字与编号一字不动）",
    mark="zmarkN")
add("表 1 手打编号的中文表题",
    purpose="中文表题（放在表格上方；应挂表题样式 ⇒ 进附表清单）", mark="zmarkO")
add_table(2, 2, {(0, 0): "表头列1", (0, 1): "表头列2",
                 (1, 0): "单元格 zmarkP1", (1, 1): "单元格 zmarkP2"},
          purpose="带表题的表格（测单元格文字套样式 + 首行重复）",
          text_summary="2×2 表格，含 zmarkP1 / zmarkP2")
marks["zmarkP1"] = "表格单元格里的标记（测表内文字被套样式后是否还在）"
marks["zmarkP2"] = "表格单元格里的标记 2"
add("表 1（续）续表样本",
    purpose="续表行（同一张表：**不挂样式** ⇒ 不进附表清单；编号一字不动）",
    mark="zmarkQ")
add("图 2（a）分图名样本",
    purpose="分图行（不区分主分：挂图题样式 ⇒ 进插图清单；编号一字不动）",
    mark="zmarkR")
add("◆ 下面一段是一张没有题注的图片，测未编号的图片告警。", purpose="说明")
if IMAGE.exists():
    pic_p = doc.add_paragraph()
    pic_p.add_run().add_picture(str(IMAGE), width=Cm(3))
    rows.append({"index": len(doc.paragraphs) - 1,
                 "purpose": "无题注的图片（应报未编号的图片）", "text": "（图片）"})
else:
    add("【Z10·目的：本该在这里放一张无题注的图，但 assets 里的样张图不存在】",
        purpose="图片缺失占位")
add("◆ 下面是一个没有题注的空表格，测未编号的表格告警。", purpose="说明")
add_table(1, 2, {}, purpose="无题注的表格（应报未编号的表格）", text_summary="1×2 空表格")

# ======================================================== 极端内容
add("◆ 下面连续 5 个空段，测空段处理。", purpose="说明")
for _ in range(5):
    add("", purpose="空段")
add("   ", purpose="只有空格的段落（3 个空格）")
add("【Z11·目的：超长段落（300 字且无句末标点），测标题误判与折行】"
    + "平台责任的边界问题需要综合考虑多种因素" * 16,
    purpose="超长无标点段落")
add("【Z12·目的：全英文正文段，测英文题目 / 英文摘要误判】"
    "This paragraph is written entirely in English without any Chinese characters, "
    "and it is long enough to be mistaken for an English title if the rules are loose.",
    purpose="全英文正文段")
add("◆ 下面两段用的是**真正的 Word 自动编号**（多级列表定义写在 numbering.xml 里，"
    "段落文字里没有编号）。测：识别阶段能不能把编号反算出来、导出时能不能把编号"
    "固化成真文字（不固化的话，清直接格式会把编号一起弄丢）。",
    purpose="说明")
zmark_w1 = doc.add_paragraph("自动编号的一级标题样本 zmarkW1")
zmark_w2 = doc.add_paragraph("自动编号的二级标题样本 zmarkW2")
_auto_id = build_auto_numbering(doc)
set_numpr(zmark_w1, _auto_id, 0)
set_numpr(zmark_w2, _auto_id, 1)
for _p, _il, _mark in ((zmark_w1, 0, "zmarkW1"), (zmark_w2, 1, "zmarkW2")):
    rows.append({"index": len(doc.paragraphs) - (2 - _il),
                 "purpose": "Word 自动编号段落（ilvl=%d，编号不在文字里）" % _il,
                 "text": _p.text})
    marks[_mark] = "自动编号段落（ilvl=%d）的标记" % _il
add("◆ 下面一段是**用样式挂编号**的写法（List Number 样式），实测 python-docx 的默认"
    "模板里这个样式不带 numPr，所以它其实没有编号——留着只是提醒：识别只看段落"
    "自己的 numPr，样式带来的编号它看不见。",
    purpose="说明")
try:
    _list_num = doc.add_paragraph("List Number 样式段落（实际没有编号）", style="List Number")
    rows.append({"index": len(doc.paragraphs) - 1,
                 "purpose": "List Number 样式段落（样式不带 numPr，实际无编号）",
                 "text": _list_num.text})
except KeyError:
    add("【Z13·目的：本想再放一个 List Number 样式段落，但模板里没有这个样式】",
        purpose="List Number 样式段落（模板缺样式，已跳过）")
add("◆ 下面一段带未接受的修订（w:ins / w:del），测未接受的修订告警；"
    "注意修订里的文字不在段落文字里。",
    purpose="说明")
add_revision_paragraph()

# ======================================================== 后置部件
add("结　论", purpose="锚点部件：结论（一级）")
add("【Z14·目的：结论与参考文献之间插入游离正文③，测复制 / 错位】游离正文③。",
    purpose="游离正文③")
add("参考文献", purpose="锚点部件：参考文献（一级）")
add("[1] 张三：平台责任研究，载《法学研究》2020 年第 3 期",
    purpose="参考文献条目①（手打序号，应被剥掉改成自动编号）", mark="zmarkT")
add("鲁晓明：论网络平台的责任边界，载《中国法学》2021 年第 5 期",
    purpose="参考文献条目②（无序号）", mark="zmarkU")
add("附　录", purpose="锚点部件：附录（一级）")
add("附录A 附录序号题名样本", purpose="附录A 序号题名（S20 样本）", mark="zmarkV")
add("【Z15·目的：附录内容区的正文，测附录整区是否被套正文样式】附录里的一段正文。",
    purpose="附录内容段落")
add("致　谢", purpose="锚点部件：致谢（一级）")
add("【Z16·目的：致谢与文末之间插入游离正文④（全文最后一段），测复制 / 丢段】"
    "游离正文④：这是全文最后一段。",
    purpose="游离正文④（文末）")
add("【Z17·目的：文末再放一段普通正文，测它的位置与告警段号】"
    "如果有孤立标题告警，应该报到这一段附近。",
    purpose="文末游离正文②")

# ======================================================== 收尾
OUT_DOCX.parent.mkdir(parents=True, exist_ok=True)
doc.save(str(OUT_DOCX))

title_index = next((r["index"] for r in rows if r["purpose"].startswith("★真题目")), None)
cover_end = (title_index - 1) if title_index is not None else None

manifest = {
    "docx": str(OUT_DOCX),
    "marks": marks,
    "cover_end": cover_end,
    "title_index": title_index,
    "paragraph_count": len(doc.paragraphs),
    "table_count": len(doc.tables),
}
OUT_MANIFEST.parent.mkdir(parents=True, exist_ok=True)
OUT_MANIFEST.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

lines = [
    f"# 压力测试夹具段落清单：{OUT_DOCX.name}",
    f"# 共 {len(doc.paragraphs)} 段、{len(doc.tables)} 张表；"
    f"真题目在段 {title_index}；封面范围候选 cover_end={cover_end}",
    "",
    f"{'段号':>5}  {'测试目的':<44} 段落文字速览",
    "-" * 120,
]
for r in rows:
    idx = "表格" if r["index"] is None else str(r["index"])
    lines.append(f"{idx:>5}  {r['purpose']:<44} {r['text']}")
OUT_HUMAN.write_text("\n".join(lines), encoding="utf-8")

print("已生成：%s" % OUT_DOCX)
print("段落 %d 段，表格 %d 张，真题目段号=%s，cover_end=%s"
      % (len(doc.paragraphs), len(doc.tables), title_index, cover_end))
print("人用清单：%s" % OUT_HUMAN)
print("机器清单：%s" % OUT_MANIFEST)
