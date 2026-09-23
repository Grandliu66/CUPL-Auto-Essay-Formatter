r"""夹具生成器 v2（产物层穷举）：对 docx 包做"合法但怪"的搬运，跑整条管线看哪里炸。

分工同 v1：**元（搬运方式）由我构思，工具只负责逐条跑完**。

为什么这一层值得单独做：它**不需要判据**。下面几类都是客观错误，不用跟"应该
是什么"比，所以能撞到"我和表都没想到的形态"：
  ✘ 崩溃        —— 管线抛异常（附后几行 traceback）
  ✘ 坏产物      —— 产物没生成、或不是能打开的 docx
  ✘ 不收敛      —— 连跑四遍：P3 与 P4 还不一样（每跑一遍文档都在变）
  ⚠️ 两遍才稳定  —— P2≠P3 但 P3=P4（第一遍没跑到稳定点）
  ⚠️ 产物残留内容控件 —— 管线没把 `w:sdt` 拆开
  ✓ 正常

用法：
    .venv\Scripts\python.exe "Adversarial Automatic Test\gen_pipeline.py"
    .venv\Scripts\python.exe "Adversarial Automatic Test\gen_pipeline.py" --all --jobs 4
    .venv\Scripts\python.exe "Adversarial Automatic Test\gen_pipeline.py" --scan   # 看目录里多了什么

⚠️ `--all` 用的是**写死的夹具清单** `FIXTURES`，不是 `glob("*.docx")`——
`reference\` 是活目录，用户稿与 Word 临时文件随时进出，glob 出来的总量每次都不一样。
"""
from __future__ import annotations

import concurrent.futures
import copy
import os
import pathlib
import re
import shutil
import sys
import time
import traceback
import zipfile

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent
for entry in (str(ROOT), str(HERE)):
    if entry not in sys.path:
        sys.path.insert(0, entry)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import aat_diff                                                 # noqa: E402
from aat_docx import qn, read_docx                              # noqa: E402
from essay_tool.config.preset_io import load_preset              # noqa: E402
from essay_tool.engine import pipeline                          # noqa: E402
from essay_tool.engine.alerts import AlertCollector              # noqa: E402
from lxml import etree                                          # noqa: E402

PRESET = ROOT / "presets" / "中国政法大学-本科毕业论文.json"
WORK = HERE / "_out" / "gen_pipeline"
DEFAULT_JOBS = 4

# 夹具清单 —— **写死，不用 glob**（用户 2026-09-17：“全量测试本身不应该固定文件吗，
# 肯定会有新文件建立在那里面的啊”）。
#
# `reference\` 是**活的工作目录**：用户稿、Word 的临时锁文件（`~$xxx.docx`）、
# 别的 Agent 的产物随时进出。拿 `glob("*.docx")` 当清单的后果实测过：
# 同一份代码两次分别跑出 **372 条**和 **682 条**，而且还把两个根本不是 zip 的
# 文件当夹具（62 条“搬运失败”）⇒ 总量与结论每次都不一样，基线就没意义了。
#
# 要加夹具＝在这里加一行（显式动作）；想看看目录里多了什么，用 `--scan`。
# 找不到文件时**报“夹具缺失”**、不当成“跳过”（清单是固定的，缺了就该看得见）。
FIXTURES: list[str] = [
    # 自造夹具（覆盖各种结构/题注/压力场景）
    "PaperStructureTest.docx",
    "PaperTest.docx",
    "TitleIdentify.docx",
    "test2.docx",
    "空白模板.docx",
    "压力！.docx",
    "压力！-排好版.docx",
    "图表压力测试.docx",
    "图表脚注测试.docx",
    "压力测试_Agent版.docx",
    # 用户真实稿子（含“已排好版”的成品回炉——本次就是它撞出的 bug）
    "TOC.docx",
    "民总.docx",
    "民总-排好版.docx",
    "民总-排好1版.docx",
    "民总-排好1版-排好版.docx",
    "一般情景测试.docx",
    "一般情景测试-排好版.docx",
    # 官方文档（自带封面：AAT 这边一律不设封面范围——“识别封面”不是我们的判据）
    "中国政法大学本科毕业论文撰写格式规范.docx",
    "中国政法大学本科生毕业论文写作模板.docx",
    "中国政法大学本科生课程考核封面.docx",
]
# 用户稿可能被挪进这个子目录（2026-09-17 对面搬过一次）——两处都找。
SUBDIR = "用户测试稿"
DEFAULT_FIXTURES = ["民总.docx", "图表脚注测试.docx"]


def _fixture_path(name: str) -> pathlib.Path | None:
    r"""夹具在哪：`reference\` 根目录 → `reference\用户测试稿\`。"""
    for path in (ROOT / "reference" / name, ROOT / "reference" / SUBDIR / name):
        if path.exists():
            return path
    return None


def _scan() -> tuple[list[str], list[str]]:
    r"""扫描 `reference\*.docx`，返回 `(能当夹具的, 被跳过的及理由)`。

    只给 `--scan` 用（想看目录里多了什么）。默认全量走 `FIXTURES`。
    """
    usable: list[str] = []
    skipped: list[str] = []
    for path in sorted((ROOT / "reference").glob("*.docx")):
        if path.name.startswith("~$"):
            skipped.append(f"{path.name}（Word 临时锁文件）")
            continue
        try:
            with zipfile.ZipFile(path) as package:
                if DOC not in package.namelist():
                    skipped.append(f"{path.name}（包里没有 {DOC}）")
                    continue
        except zipfile.BadZipFile:
            skipped.append(f"{path.name}（不是 zip）")
            continue
        usable.append(path.name)
    return usable, skipped


def _output_paths(tag: str) -> tuple[pathlib.Path, pathlib.Path]:
    """现场目录与明细文件按 `--tag` 分家。

    因为同一条项目上可能有两个 Agent 同时跑这个工具：共用一个目录的话，
    一边收工时的 `rmtree` 会把另一边正在看的现场删掉。
    """
    suffix = f"-{tag}" if tag else ""
    return (HERE / "_out" / f"gen_pipeline{suffix}",
            HERE / "_out" / f"gen-pipeline{suffix}.txt")
DOC = "word/document.xml"


# ---------------------------------------------------------------- XML 小工具
def _body(root):
    return root.find(qn("w:body"))


def _paras(root):
    body = _body(root)
    return body, body.findall(qn("w:p"))


def _sectpr(root):
    return _body(root).find(qn("w:sectPr"))


def _ensure_pPr(para):
    pPr = para.find(qn("w:pPr"))
    if pPr is None:
        pPr = etree.Element(qn("w:pPr"))
        para.insert(0, pPr)
    return pPr


def _new(tag, **attrs):
    element = etree.Element(qn(tag))
    for key, value in attrs.items():
        element.set(qn(f"w:{key}"), value)
    return element


# ---------------------------------------------------------------- 搬运方式（元）
def m_wrap_sdt(root, *_):
    """把 body 里所有段落包进一个内容控件（目录裹在控件里那种形态，扩到全篇）。"""
    body, paras = _paras(root)
    if not paras:
        return
    sdt = _new("w:sdt")
    content = _new("w:sdtContent")
    sdt.append(content)
    sect = _sectpr(root)
    for para in paras:
        body.remove(para)
        content.append(para)
    body.insert(list(body).index(sect) if sect is not None else 0, sdt)


def m_sectpr_to_para(root, *_):
    """把 body 末尾的分节符另抄一份挂到首段的 pPr 上（分节符挂在段落标记上）。"""
    sect = _sectpr(root)
    _body_, paras = _paras(root)
    if sect is None or not paras:
        return
    _ensure_pPr(paras[0]).append(copy.deepcopy(sect))


def m_mid_section(root, *_):
    """在文档中间那一段再挂一个分节符（多一节、且前半节没有自己的页边距定义）。"""
    sect = _sectpr(root)
    _body_, paras = _paras(root)
    if sect is None or len(paras) < 3:
        return
    broken = copy.deepcopy(sect)
    for tag in ("w:pgSz", "w:pgMar", "w:headerReference", "w:footerReference"):
        for child in broken.findall(qn(tag)):
            broken.remove(child)
    _ensure_pPr(paras[len(paras) // 2]).append(broken)


def m_drop_pstyles(root, *_):
    """删掉所有段落的样式引用。"""
    for para in _paras(root)[1]:
        pPr = para.find(qn("w:pPr"))
        if pPr is None:
            continue
        for style in pPr.findall(qn("w:pStyle")):
            pPr.remove(style)


def m_dangling_pstyle(root, *_):
    """把所有段落的样式引用指向一个**不存在**的样式。"""
    for para in _paras(root)[1]:
        pPr = _ensure_pPr(para)
        style = pPr.find(qn("w:pStyle"))
        if style is None:
            style = _new("w:pStyle")
            pPr.insert(0, style)
        style.set(qn("w:val"), "__AAT无此样式__")


def _numid_values(root):
    _body_, paras = _paras(root)
    for para in paras:
        pPr = para.find(qn("w:pPr"))
        numPr = pPr.find(qn("w:numPr")) if pPr is not None else None
        if numPr is None:
            continue
        holder = numPr.find(qn("w:numId"))
        if holder is not None:
            yield holder


def m_dangling_numid(root, *_):
    """自动编号指向一个不存在的编号定义。"""
    for holder in _numid_values(root):
        holder.set(qn("w:val"), "9999")


def m_zero_numid(root, *_):
    """自动编号置 0（OOXML 里表示"取消编号"）。"""
    for holder in _numid_values(root):
        holder.set(qn("w:val"), "0")


def m_drop_numpr(root, *_):
    """删掉所有段落的自动编号引用。"""
    for para in _paras(root)[1]:
        pPr = para.find(qn("w:pPr"))
        if pPr is None:
            continue
        for numPr in pPr.findall(qn("w:numPr")):
            pPr.remove(numPr)


def m_change_numfmt(root, styles=None, numbering=None):
    """把所有编号定义的编号格式换成中文数字（十进制→中文，序号写法会变样）。"""
    if numbering is None:
        return
    for fmt in numbering.iter(qn("w:numFmt")):
        fmt.set(qn("w:val"), "chineseCounting")


def m_unclosed_field(root, *_):
    """删掉最后一处域的结束符 —— 域没闭合。"""
    ends = [node for node in root.iter(qn("w:fldChar"))
            if node.get(qn("w:fldCharType")) == "end"]
    if ends:
        ends[-1].getparent().remove(ends[-1])


def m_narrow_toc(root, *_):
    """把目录域的范围从 1-3 收窄到 1-1。"""
    for instr in root.iter(qn("w:instrText")):
        if instr.text and "\\o" in instr.text:
            instr.text = instr.text.replace('\\o "1-3"', '\\o "1-1"')
            instr.text = instr.text.replace("\\o \"1-4\"", '\\o "1-1"')


def m_table_first(root, *_):
    """在第一个段落之前插一张两格的表（表格当首段）。"""
    body, _paras_ = _paras(root)
    tbl = _new("w:tbl")
    for _row in range(1):
        tr = _new("w:tr")
        for text in ("AAT", "表格"):
            tc = _new("w:tc")
            para = _new("w:p")
            run = _new("w:r")
            holder = _new("w:t")
            holder.text = text
            run.append(holder)
            para.append(run)
            tc.append(para)
            tr.append(tc)
        tbl.append(tr)
    body.insert(0, tbl)


def m_duplicate_body(root, *_):
    """把 body 的内容整体再抄一份（段落翻倍、部件全部重复）。"""
    body, paras = _paras(root)
    sect = _sectpr(root)
    for para in paras:
        body.insert(list(body).index(sect) if sect is not None else len(body),
                    copy.deepcopy(para))


def m_drop_pgsz(root, *_):
    """删掉节里的纸张定义（页宽页高缺失）。"""
    for sect in root.iter(qn("w:sectPr")):
        for size in sect.findall(qn("w:pgSz")):
            sect.remove(size)


def m_wrap_in_ins(root, *_):
    """把首段的第一个 run 包进"未接受的插入修订"。"""
    _body_, paras = _paras(root)
    for para in paras:
        run = para.find(qn("w:r"))
        if run is None:
            continue
        ins = _new("w:ins", id="9001", author="AAT", date="2026-09-17T00:00:00Z")
        para.replace(run, ins)
        ins.append(run)
        return


def m_wrap_in_hyperlink(root, *_):
    """把首段的 run 包进超链接。"""
    _body_, paras = _paras(root)
    for para in paras:
        runs = para.findall(qn("w:r"))
        if not runs:
            continue
        link = _new("w:hyperlink", anchor="aat", history="1")
        para.replace(runs[0], link)
        link.append(runs[0])
        return


def m_wrap_all_hyperlinks(root, *_):
    """**全部**段落的 run 都包进超链接（不是只包首段）。"""
    _body_, paras = _paras(root)
    for para in paras:
        for run in list(para.findall(qn("w:r"))):
            link = _new("w:hyperlink", anchor="aat", history="1")
            para.replace(run, link)
            link.append(run)


def m_float_image(root, *_):
    """内嵌图片改浮动：第一个 `wp:inline` → `wp:anchor`（补齐锢定定位所需）。

    合法但怪：锢定图片要 `positionH/positionV` 与 `wrapNone` 等一套定位元素
    才说得通；这里按 schema 的顺序补上（插在 `extent` 前、`docPr` 前）。
    夹具里没图片时不动（空操作）。
    """
    WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    inline = root.find(".//" + WP + "inline")
    if inline is None:
        return
    inline.tag = WP + "anchor"
    for attr, value in (("distT", "0"), ("distB", "0"), ("distL", "0"),
                        ("distR", "0"), ("simplePos", "0"),
                        ("relativeHeight", "251658240"), ("behindDoc", "0"),
                        ("locked", "0"), ("layoutInCell", "1"),
                        ("allowOverlap", "1")):
        inline.set(attr, value)
    simple = etree.Element(WP + "simplePos")
    simple.set("x", "0")
    simple.set("y", "0")
    pos_h = etree.Element(WP + "positionH")
    pos_h.set("relativeFrom", "column")
    etree.SubElement(pos_h, WP + "posOffset").text = "0"
    pos_v = etree.Element(WP + "positionV")
    pos_v.set("relativeFrom", "paragraph")
    etree.SubElement(pos_v, WP + "posOffset").text = "0"
    for index, element in enumerate((simple, pos_h, pos_v)):
        inline.insert(index, element)
    wrap = etree.Element(WP + "wrapNone")
    doc_pr = inline.find(WP + "docPr")
    if doc_pr is not None:
        doc_pr.addprevious(wrap)
    else:
        inline.append(wrap)


def m_field_in_header(root, styles, numbering, entries):
    """页眉里塞一个**没闭合**的目录域（Word 打开会提示修复；管线不许崩）。

    这个搬运要改的是 `word/header?.xml`，不是 document.xml——所以它收第四个
    参数 `entries`（整包“名字→字节”），见 `_call`。夹具没有页眉部件时不动。
    """
    name = next((n for n in sorted(entries)
                 if n.startswith("word/header") and n.endswith(".xml")), None)
    if name is None:
        return
    header = etree.fromstring(entries[name])
    body = header.find(qn("w:hdr"))
    if body is None:
        return
    para = etree.SubElement(body, qn("w:p"))
    run = etree.SubElement(para, qn("w:r"))
    etree.SubElement(run, qn("w:fldChar")).set(qn("w:fldCharType"), "begin")
    run = etree.SubElement(para, qn("w:r"))
    etree.SubElement(run, qn("w:instrText")).text = ' TOC \\o "1-3" '
    run = etree.SubElement(para, qn("w:r"))
    etree.SubElement(run, qn("w:t")).text = "（页眉里的目录域）"
    entries[name] = etree.tostring(header, xml_declaration=True,
                                   encoding="UTF-8", standalone=True)


def m_caption_into_cell(root, *_):
    """把最后一个题注样子的段落搬进一张表的单元格里（表格单元格里的题注）。"""
    body, paras = _paras(root)
    target = None
    for para in paras:
        text = "".join(node.text or "" for node in para.iter(qn("w:t")))
        if re.match(r"^[图表]\s*\d", text):
            target = para
    if target is None:
        return
    index = list(body).index(target)
    tc = _new("w:tc")
    tc.append(target)                       # 从 body 摘走、移进单元格
    tr = _new("w:tr")
    tr.append(tc)
    tbl = _new("w:tbl")
    tbl.append(tr)
    body.insert(index, tbl)


def m_style_numbering(root, styles=None, *_):
    """把自动编号挂到 Normal 样式上（段落自己不带 numPr，全靠样式带）。"""
    if styles is None:
        return
    for style in styles.iter(qn("w:style")):
        if style.get(qn("w:styleId")) not in ("Normal", "a"):
            continue
        pPr = style.find(qn("w:pPr"))
        if pPr is None:
            pPr = etree.Element(qn("w:pPr"))
            style.append(pPr)
        if pPr.find(qn("w:numPr")) is None:
            numPr = _new("w:numPr")
            numPr.append(_new("w:ilvl", val="0"))
            numPr.append(_new("w:numId", val="1"))
            pPr.insert(0, numPr)


def m_ins_whole_para(root, *_):
    """把文档中间那一整段包进“未接受的插入修订”（块级 `w:ins`）。"""
    body, paras = _paras(root)
    if len(paras) < 3:
        return
    target = paras[len(paras) // 2]
    index = list(body).index(target)
    ins = _new("w:ins", id="9002", author="AAT", date="2026-09-17T00:00:00Z")
    body.remove(target)
    ins.append(target)
    body.insert(index, ins)


def m_seq_label_formula(root, *_):
    """把 SEQ 域的标签换成“公式”（表里说这类标签不属于图/表）。"""
    for instr in root.iter(qn("w:instrText")):
        if instr.text and "SEQ" in instr.text:
            instr.text = re.sub(r"(SEQ\s+)\S+", r"\g<1>公式", instr.text)


def m_toc_drop_switches(root, *_):
    """目录域指令去掉 `\\o` 参数（变成 Word 的默认收法）。"""
    for instr in root.iter(qn("w:instrText")):
        if instr.text and "TOC" in instr.text:
            instr.text = re.sub(r'\s*\\o\s*"[^"]*"', "", instr.text)


def m_list_field_bad_style(root, *_):
    """清单域的样式名改成不存在的（`\\t "图题,1"` 收不到任何条目）。"""
    for instr in root.iter(qn("w:instrText")):
        if instr.text and "\\t" in instr.text:
            instr.text = re.sub(r'"[^"]*",1', '"__AAT无此样式__,1', instr.text)


def m_outline_level(root, *_):
    """给每个段落都写上大纲级别（识别端说不采信，输出端会写）。"""
    for para in _paras(root)[1]:
        pPr = _ensure_pPr(para)
        if pPr.find(qn("w:outlineLvl")) is None:
            pPr.append(_new("w:outlineLvl", val="0"))


def m_huge_paragraph(root, *_):
    """把首段文字拉成 5000 字。"""
    for para in _paras(root)[1]:
        for node in para.iter(qn("w:t")):
            node.text = (node.text or "") + "字" * 5000
            return


def m_invisible_chars(root, *_):
    """首段文字里插全角空格与零宽字符。"""
    for para in _paras(root)[1]:
        for node in para.iter(qn("w:t")):
            node.text = "　\u200b" + (node.text or "") + "\u200b　"
            return


def m_extra_sectpr_at_end(root, *_):
    """body 末尾再放一个 sectPr（拼接文档常见的多余节定义）。"""
    sect = _sectpr(root)
    if sect is None:
        return
    _body(root).append(copy.deepcopy(sect))


def m_caption_style(root, *_):
    """给题注样子的段落挂上 caption 样式（别的工具生成的题注长这样）。"""
    for para in _paras(root)[1]:
        text = "".join(node.text or "" for node in para.iter(qn("w:t")))
        if re.match(r"^[图表]\s*\d", text):
            pPr = _ensure_pPr(para)
            if pPr.find(qn("w:pStyle")) is None:
                pPr.insert(0, _new("w:pStyle", val="caption"))


def m_vanish(root, *_):
    """把首段设成隐藏文字。"""
    for para in _paras(root)[1]:
        for run in para.findall(qn("w:r")):
            rPr = run.find(qn("w:rPr"))
            if rPr is None:
                rPr = etree.Element(qn("w:rPr"))
                run.insert(0, rPr)
            if rPr.find(qn("w:vanish")) is None:
                rPr.append(_new("w:vanish"))
        return


def m_seq_in_cell(root, *_):
    """在最后一张表的单元格里塞一个 SEQ 域（表格里的题注域）。"""
    tables = list(root.iter(qn("w:tbl")))
    if not tables:
        return
    tr = tables[-1].find(qn("w:tr"))
    tc = tr.find(qn("w:tc")) if tr is not None else None
    if tc is None:
        return
    para = _new("w:p")
    run = _new("w:r")
    run.append(_new("w:fldChar", fldCharType="begin"))
    instr = etree.Element(qn("w:instrText"))
    instr.text = " SEQ 图 \\* ARABIC "
    run.append(instr)
    run.append(_new("w:fldChar", fldCharType="end"))
    para.append(run)
    tc.append(para)


def m_drop_tbl_header(root, *_):
    """删掉所有“跨页重复标题行”标记。"""
    for node in list(root.iter(qn("w:tblHeader"))):
        node.getparent().remove(node)


def m_identity(root, *_):
    """基线：不施加任何搬运。

    必须留着这一条：否则分不清“不收敛”是搬运引起的、还是这份夹具本来就跑不到
    稳定点（比如它本身就是一个边缘文档）。
    """
    return


MUTATIONS: list[tuple[str, object]] = [
    ("内容控件包住全部段落", m_wrap_sdt),
    ("分节符抄到首段 pPr", m_sectpr_to_para),
    ("文档中间插一个分节符", m_mid_section),
    ("删掉全部 pStyle", m_drop_pstyles),
    ("pStyle 指向不存在的样式", m_dangling_pstyle),
    ("numId 指向不存在的定义", m_dangling_numid),
    ("numId 置 0", m_zero_numid),
    ("删掉全部 numPr", m_drop_numpr),
    ("编号格式全换中文数字", m_change_numfmt),
    ("域没闭合（删掉 end）", m_unclosed_field),
    ("目录域收窄到 1-1", m_narrow_toc),
    ("body 开头插一张表", m_table_first),
    ("body 内容整体复制", m_duplicate_body),
    ("删掉节的纸张定义", m_drop_pgsz),
    ("首段 run 包进未接受的修订", m_wrap_in_ins),
    ("首段 run 包进超链接", m_wrap_in_hyperlink),
    ("全部段落包进超链接", m_wrap_all_hyperlinks),
    ("图片改浮动", m_float_image),
    ("页眉里塞没闭合的目录域", m_field_in_header),
    # —— 2026-09-17 第二批：上面 16 种之后的都排在这里，顺序别动（job 编号靠它）
    ("题注搬进表格单元格", m_caption_into_cell),
    ("Normal 样式挂自动编号", m_style_numbering),
    ("整段包进未接受的修订", m_ins_whole_para),
    ("SEQ 标签换“公式”", m_seq_label_formula),
    ("目录域删掉 \\o 参数", m_toc_drop_switches),
    ("清单域样式名改错", m_list_field_bad_style),
    ("给段落写大纲级别", m_outline_level),
    ("首段拉到 5000 字", m_huge_paragraph),
    ("文字里插全角空格与零宽字符", m_invisible_chars),
    ("body 末尾再加一个 sectPr", m_extra_sectpr_at_end),
    ("题注挂 caption 样式", m_caption_style),
    ("首段设成隐藏文字", m_vanish),
    ("单元格里塞 SEQ 域", m_seq_in_cell),
    ("删掉表格重复标题行", m_drop_tbl_header),
    ("基线：不施加搬运", m_identity),
]


# ---------------------------------------------------------------- 搬运 + 跑
def _call(mutation, root, styles, numbering, entries) -> None:
    """叫搬运函数干活。

    2026-09-19 起多给第 4 个参数 `entries`（整包“名字→字节”）——页眉/页脚
    那类搬运要改的不是 document.xml。旧搬运只认前三个参数，用签名分辨，
    不用 try/except（函数内部抛的 TypeError 会被误当成“参数个数不对”）。
    """
    import inspect
    signature = inspect.signature(mutation)
    positional = [p for p in signature.parameters.values()
                  if p.kind in (p.POSITIONAL_ONLY, p.POSITIONAL_OR_KEYWORD)]
    variadic = any(p.kind is p.VAR_POSITIONAL
                   for p in signature.parameters.values())
    if variadic or len(positional) >= 4:
        mutation(root, styles, numbering, entries)
    else:
        mutation(root, styles, numbering)


def mutate(src: pathlib.Path, dst: pathlib.Path, mutation) -> None:
    """把 src 复制成 dst，并对包内 XML 施加一次搬运。"""
    with zipfile.ZipFile(src) as package:
        entries = {name: package.read(name) for name in package.namelist()}
    root = etree.fromstring(entries[DOC])
    styles = etree.fromstring(entries["word/styles.xml"]) if "word/styles.xml" in entries else None
    numbering = etree.fromstring(entries["word/numbering.xml"]) if "word/numbering.xml" in entries else None
    _call(mutation, root, styles, numbering, entries)
    entries[DOC] = etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                                  standalone=True)
    if styles is not None:
        entries["word/styles.xml"] = etree.tostring(
            styles, xml_declaration=True, encoding="UTF-8", standalone=True)
    if numbering is not None:
        entries["word/numbering.xml"] = etree.tostring(
            numbering, xml_declaration=True, encoding="UTF-8", standalone=True)
    dst.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as package:
        for name, data in entries.items():
            package.writestr(name, data)


def _run(source: pathlib.Path, target: pathlib.Path) -> tuple[pathlib.Path, list[str]]:
    """跑一遍管线，返回 `(产物路径, 告警文案)`。

    **告警也要收**：它是用户直接看到的东西，而“告警稳不稳定”不需要判据
    （同一份文档跑两遍，告警集合应当一样）；产物里少了一段、告警里往往能看出来。
    """
    target.mkdir(parents=True, exist_ok=True)
    collector = AlertCollector()
    product = pipeline.run(source, None, target, preset_config=load_preset(PRESET),
                           alerts=collector)
    collector.consolidate()
    return product, [alert.line for alert in collector.items]


def _texts(path: pathlib.Path) -> tuple[list[str], int]:
    """产物里**全部段落的文字**（按 XML 文档顺序，表格里的也算）＋ 内容控件个数。

    为什么不用 `aat_docx.read_docx` 比：它按“`w:body` 的直接子段落”数段号
    （与管线拆完控件的口径一致），产物里一旦残留 `w:sdt`，两侧段号就会错位、
    报出假差异（这条是它自己 warning 里写的）。比文字不受这个影响。
    """
    with zipfile.ZipFile(path) as package:
        root = etree.fromstring(package.read(DOC))
    body = root.find(qn("w:body"))
    if body is None:
        return [], 0
    texts = ["".join(node.text or "" for node in para.iter(qn("w:t")))
             for para in body.iter(qn("w:p"))]
    return texts, len(list(body.iter(qn("w:sdt"))))


def _job_name(fixture: str, mutation: str) -> str:
    r"""现场目录名＝“夹具-搬运”。

    ⚠️ 不要用序号：`--only` 过滤后序号会从头重排，同一个 `001` 在两次运行里
    可能是两条不同的搬运（交现场时就指错了）。用名字则不管怎么过滤都指得准。
    """
    stem = pathlib.Path(fixture).stem
    return re.sub(r'[<>:"/\\|?*]', "_", f"{stem}-{mutation}")


def one(payload: tuple[str, str, str, str]) -> tuple[str, str, str, str]:
    """跑一条搬运。返回 `(结论, 夹具, 搬运, 说明)`。"""
    fixture_name, mutation_name, work_dir, job = payload
    fixture = _fixture_path(fixture_name)
    if fixture is None:
        return ("✘夹具缺失（清单里写着、磁盘上没有）", fixture_name, mutation_name, "")
    case = pathlib.Path(work_dir) / job
    shutil.rmtree(case, ignore_errors=True)
    try:
        mutate(fixture, case / "mutated.docx", dict(MUTATIONS)[mutation_name])
    except Exception:
        return ("✘搬运失败（这个元本身写错了）", fixture_name, mutation_name,
                traceback.format_exc().strip().splitlines()[-1])
    try:
        first, alerts1 = _run(case / "mutated.docx", case / "p1")
        second, alerts2 = _run(first, case / "p2")
        third, alerts3 = _run(second, case / "p3")
        fourth, alerts4 = _run(third, case / "p4")
    except Exception:
        return ("✘崩溃", fixture_name, mutation_name,
                "\n".join(traceback.format_exc().strip().splitlines()[-4:]))
    for path in (first, second, third, fourth):
        if not pathlib.Path(path).exists():
            return ("✘坏产物", fixture_name, mutation_name, f"没生成 {path}")
    counts = (f"告警条数 P1={len(alerts1)} P2={len(alerts2)} "
              f"P3={len(alerts3)} P4={len(alerts4)}")
    try:
        texts2, sdts2 = _texts(second)
        texts3, sdts3 = _texts(third)
        texts4, _sdts4 = _texts(fourth)
    except Exception:
        return ("✘坏产物", fixture_name, mutation_name,
                "产物打不开：" + traceback.format_exc().strip().splitlines()[-1])
    if texts3 != texts4:
        verdict = "✘不收敛（每遍都变）"
    elif texts2 != texts3:
        verdict = "⚠️两遍才稳定"
    else:
        verdict = ""
    if verdict:
        lines = [counts, f"段数 P2={len(texts2)} P3={len(texts3)} P4={len(texts4)}",
                 f"（内容控件：P2 {sdts2} 个、P3 {sdts3} 个）"]
        for left, right, name_left, name_right in ((texts2, texts3, "P2", "P3"),
                                                  (texts3, texts4, "P3", "P4")):
            span = min(len(left), len(right))
            first_diff = next((i for i in range(span) if left[i] != right[i]), span)
            lines.append(f"  {name_left}/{name_right} 第 {first_diff} 段起：")
            for i in range(first_diff, min(first_diff + 3, max(len(left), len(right)))):
                one_left = left[i][:34] if i < len(left) else "（无）"
                one_right = right[i][:34] if i < len(right) else "（无）"
                lines.append(f"    {name_left}[{i}]={one_left!r}  {name_right}[{i}]={one_right!r}")
        return (verdict, fixture_name, mutation_name, "\n".join(lines))
    if sorted(alerts3) != sorted(alerts4):
        only3 = [line for line in alerts3 if line not in alerts4]
        only4 = [line for line in alerts4 if line not in alerts3]
        lines = [counts, "P3 独有：" + ("；".join(only3) or "（无）"),
                 "P4 独有：" + ("；".join(only4) or "（无）")]
        return ("⚠️告警不收敛", fixture_name, mutation_name, "\n".join(lines))
    if sdts2 or sdts3:
        detail = (f"{counts}；产物里残留内容控件 {sdts2} 个（P2） / {sdts3} 个（P3）"
                  f"、且文本与样式都已是稳定点")
        return ("⚠️产物残留内容控件", fixture_name, mutation_name, detail)
    try:
        diff = aat_diff.compare_shapes(read_docx(second), read_docx(third))
        report = aat_diff.render(diff, max_items=4)
    except Exception:
        return ("✘坏产物", fixture_name, mutation_name,
                "产物打不开：" + traceback.format_exc().strip().splitlines()[-1])
    if "没有差异" not in report:
        return ("⚠️格式不幂等", fixture_name, mutation_name,
                counts + "\n" + "\n".join(report.splitlines()[:14]))
    return ("✓正常", fixture_name, mutation_name, counts)


def main(argv: list[str]) -> int:
    only = argv[argv.index("--only") + 1] if "--only" in argv else None
    jobs = int(argv[argv.index("--jobs") + 1]) if "--jobs" in argv else DEFAULT_JOBS
    keep = "--keep" in argv
    tag = argv[argv.index("--tag") + 1] if "--tag" in argv else ""
    work, report = _output_paths(tag)
    given = [name for name in argv if not name.startswith("--")
             and name not in (only, str(jobs), tag)]
    if given:
        names = given
    elif "--scan" in argv:
        # **只看不跑**：拿写死的清单跟磁盘现状对账，看多了什么、少了什么。
        usable, skipped = _scan()
        print(f"reference\\ 下能当夹具的共 {len(usable)} 份（清单里有 {len(FIXTURES)} 份）：")
        for name in usable:
            mark = "" if name in FIXTURES else "   ← 清单里没有"
            print(f"  {name}{mark}")
        for name in FIXTURES:
            if name not in usable:
                print(f"  {name}   ← 清单里有、磁盘上找不到（或不可用）")
        if skipped:
            print("跳过的：")
            for line in skipped:
                print(f"  {line}")
        return 0
    elif "--all" in argv:
        names = FIXTURES          # ← 写死的清单，不 glob（见文件头说明）
    else:
        names = DEFAULT_FIXTURES

    plan = []
    for fixture in names:
        for mutation_name, _fn in MUTATIONS:
            if only and only not in mutation_name:
                continue
            plan.append((fixture, mutation_name, _job_name(fixture, mutation_name)))
    payloads = [(fixture, mutation, str(work), job)
                for fixture, mutation, job in plan]

    started = time.perf_counter()
    if jobs > 1:
        with concurrent.futures.ProcessPoolExecutor(max_workers=jobs) as pool:
            outcomes = list(pool.map(one, payloads))
    else:
        outcomes = [one(payload) for payload in payloads]
    elapsed = time.perf_counter() - started

    tally: dict[str, int] = {}
    for conclusion, _fixture, _mutation, _detail in outcomes:
        tally[conclusion] = tally.get(conclusion, 0) + 1
    print(f"跑完 {len(outcomes)} 条（{jobs} 进程）："
          + "、".join(f"{key} {value}" for key, value in sorted(tally.items())))
    print(f"耗时 {elapsed:.1f} 秒（{elapsed / max(len(outcomes), 1):.1f} 秒/条，每条 4 遍管线）")
    print()
    print("按夹具：")
    by_fixture: dict[str, dict[str, int]] = {}
    for conclusion, fixture, _mutation, _detail in outcomes:
        counter = by_fixture.setdefault(fixture, {})
        counter[conclusion] = counter.get(conclusion, 0) + 1
    for fixture in sorted(by_fixture):
        items = by_fixture[fixture]
        summary = "、".join(f"{key} {value}" for key, value in sorted(items.items()))
        print(f"  {fixture:<26} 共 {sum(items.values())}   {summary}")
    print()
    for conclusion, fixture, mutation, detail in outcomes:
        if conclusion.startswith("✓"):
            continue
        print(f"  {conclusion}  [{fixture}] {mutation}")
        for line in (detail or "").splitlines():
            print(f"        {line}")
    report.write_text("\n".join(
        f"{conclusion}\t{fixture}\t{mutation}\t{detail}"
        for conclusion, fixture, mutation, detail in outcomes), encoding="utf-8")
    print()
    print(f"明细已写出：{report.relative_to(ROOT)}")
    if keep:
        index = work / "索引.txt"
        index.write_text("\n".join(f"{job}\t{fixture}\t{mutation}"
                                   for fixture, mutation, job in plan), encoding="utf-8")
        print(f"现场已保留：{work.relative_to(ROOT)}（每个子目录：mutated.docx 与 p1~p4）")
        print(f"          索引：{(work / '索引.txt').relative_to(ROOT)}")
    else:
        shutil.rmtree(work, ignore_errors=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
