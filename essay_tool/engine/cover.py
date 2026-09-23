"""封面字段：让用户自己填的那部分内容 + 书写规则。

GUI（阶段 4）以后填的就是这里的 `CoverFields` 结构；生成封面时按下面的
规则逐条套用。三条**不许做**的红线（都是用户 2026-09-12 的明确要求）：

1. **学号不做任何格式校验/改写**——真的会有字母开头、甚至汉字开头的学号。
   一律原样进、原样出（本模块故意不提供任何"学号检查"函数，别加）。
2. 姓名只做"**两个字中间空一格**"这一条美化（"张三"→"张 三"，与官方模板的
   "王 五"一致）。**四个字、少数民族名字（含"·"）都是合法的**——不截断、
   不拆开、不按"两字姓＋两字名"那套去解析。
3. **指导教师姓名和职称是两个独立的量**：同时填了就"姓名＋空格＋职称"显示；
   课程考核封面**默认不带职称**（任课教师只写姓名）。

封面信息／书写规则（本模块上半部分）＋ 封面组装与插入（下半部分）。

组装：`build_thesis_cover_document` / `build_year_cover_document` /
`build_course_cover_document` / `build_tenstar_cover_document`——从
reference/ 的官方模板里抽封面页、清理、按字段填字；插入：
`add_cover(target, cover_type, fields)` 按类型键把封面拼到已生成文档的最前面
（`add_thesis_cover` 等是它背后各自的一键操作）。**论文/学年论文/十星封面
是纯文字，可以安全拼插；课程封面带校徽／书法图片和页眉部件，拼插时要把
"部件＋关系"一起搬**——这件事`prepend_cover_document` 已经做完（图片关系
迁移、页眉部件搬迁、样式去撞名）。
"""

from __future__ import annotations

import copy
import io
import math
import pathlib
from dataclasses import dataclass

from docx import Document
from docx.document import Document as DocumentObject
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Twips
from docx.text.paragraph import Paragraph

from essay_tool.config.integrity import verify_official
from essay_tool.config.paths import resource_dir

# ------------------------------------------------------------------ 字段


@dataclass
class CoverFields:
    """封面要靠用户填的内容。全部可留空——留空就是空位留白给手填。"""

    # —— 毕业论文封面 ——
    number: str = ""        # 论文编号（多是学校事后打印的位置，可留空）
    title: str = ""
    subtitle: str = ""
    student_name: str = ""
    student_id: str = ""    # 学号：原样使用，**不校验格式**（可能字母/汉字开头）
    college: str = ""
    grade: str = ""         # 如 "2022级"
    date: str = ""          # 完成日期；留空就是空，不补系统日期
    major: str = ""
    advisor_name: str = ""  # 指导教师姓名（与职称是两个独立的量）
    advisor_title: str = ""  # 指导教师职称；空则只显示姓名
    score: str = ""         # 论文成绩：默认留空给老师手填

    # —— 课程考核封面 ——
    course_name: str = ""
    course_code: str = ""
    course_seq: str = ""
    teacher_name: str = ""  # 任课教师：**默认只写姓名、不带职称**（用户拍板）；
                            # 以后要带职称再加开关，别把职称塞进这个字段

    # —— 学术十星封面 ——
    # 届数（如 "二十四"，不带"第""届"）：封面固定字样"第二十四届"和页眉里的 `{届数}`
    # 占位都用它（用户 2026-09-19："在 GUI 允许用户填写届数"）。空＝封面保持范例原文、
    # 页眉占位换成空。**页眉那一半不走 CoverFields**（封面关了也要换），走
    # `pipeline.run(session=…)`，见 `sections.expand_header_tokens`。
    session: str = ""


# ------------------------------------------------------------------ 规则


def format_person_name(name: str | None) -> str:
    """人名规范：**两个字**的名字中间空一格（"张三"→"张 三"）。

    其余一律原样：三个字不动；四个字、少数民族名字、名字里的"·"都合法，
    不截断不拆开。先去掉所有空白字符（防止"张  三"这类），再去判断长度——
    所以"张 三"再跑一遍还是"张 三"，重复调用结果稳定。
    """
    chars = "".join(ch for ch in (name or "") if not ch.isspace())
    if len(chars) == 2:
        return f"{chars[0]} {chars[1]}"
    return chars


def format_advisor(name: str | None, title: str | None) -> str:
    """指导教师显示：姓名（按 `format_person_name`）＋空一格＋职称。

    两个量各自可空：职称空 → 只写姓名；姓名空 → 只写职称；都空 → 空串
    （封面留白，等人手填）。
    """
    name_part = format_person_name(name)
    title_part = (title or "").strip()
    return " ".join(part for part in (name_part, title_part) if part)


def resolve_date(raw: str | None) -> str:
    """完成日期：原样收下（留空就是空，不补系统日期）。"""
    return (raw or "").strip()


# ================================================================ 组装
#
# 从 reference/ 的官方模板里“抽封面页 → 清理 → 按字段填字”。

ROOT = resource_dir()       # 只读资源根：开发时是仓库根，打包后是 _internal
THESIS_TEMPLATE = ROOT / "reference" / "中国政法大学本科生毕业论文写作模板.docx"
COURSE_TEMPLATE = ROOT / "reference" / "中国政法大学本科生课程考核封面.docx"

THESIS_BLANK_WIDTH = 16  # 下划线长度（半角位）。比模板的 18 少 2 位：两栏窄栏
# 里行宽贴边（模板原行 ≈196 磅 / 栏宽 ≈197 磅），换个渲染器就容易被挤到换行，
# 而“两端对齐”会把行内所有空格（含标签里“学    号”的间隔）拉得巨大
# （2026-09-12 用户截图实拍）。缩短留出安全余量，配合左对齐双保险。

# 行尾"零宽不换行"收尾字符。为什么必须有它：老式兼容模式下，Word 会把
# **行尾空格**裁掉不让它参与对齐——课程封面（模板是兼容模式文档）填上值后
# 每行各按自己值的长度居中、歪歪扭扭（2026-09-12 用户截图实拍）。末尾跟一个
# 非空字符，空格就不再是"行尾空格"，整条线的宽度都参与居中，而且零宽、不可见。
_TERMINATOR = "\uFEFF"

# 悬挂缩进的施加门槛（磅）：只有"提示词＋内容"超过它（栏宽 197 磅装不下）
# 的行才加悬挂。为什么不能全员加：老式兼容模式（compat 14，如民总.docx）
# 会把该行行宽算成"栏宽－左缩进"＝127 磅，于是标准行（70＋112＝182 磅）
# 也凭空折成两行；而未加缩进的行两种模式都按整栏 197 磅排，稳稳单行。
# 独立样本在 compat 15 下首行拿回悬挂量（197 磅），单行行观感完全一样，
# 所以"标准行不加、超长行才加"对两种模式都最优（2026-09-13 坐标尸检＋实验）。
_HANGING_MIN_PT = 193.0

# 封面各段的"标准行距"（磅）：以独立样张的实录为准，在建封面时逐段以
# atLeast 钉死。为什么必须钉死：模板封面各段没写行距，靠文档级兼容设置
# （compat 14/15、网格算法）算出来的自然行距在两种文档里能差一倍——插进
# 民总后 26 磅大标题的行高从 ~37 磅缩到 ~20 磅，第二行和副标题直接叠在
# 一起（2026-09-13 用户实拍"大标题换行行间距太小"）。绝对 atLeast 行距
# ＋关网格对齐 = 跨文档像素级稳定；独立样张本来就等于这些值，观感不变。
_COVER_PITCH_PT = {
    0: 31.5, 1: 25.5, 2: 31.5, 3: 30.0, 4: 15.8, 5: 15.8, 6: 15.8, 7: 21.8,
    8: 46.5, 9: 40.5, 10: 15.8, 11: 16.5, 12: 36.8, 13: 24.8,
    14: 15.8, 15: 15.8, 16: 15.8, 17: 15.8, 18: 15.8, 19: 15.8, 20: 15.8,
    21: 15.8, 23: 21.7,
}

# 上面这些段落是**空白垫段**（或只有分节符）：必须用 EXACTLY 精确行距——
# 民总里它们的自然行距是 18 磅，atLeast 15.8 根本压不住，一段多 2.2 磅、
# 十六段累积 30 磅，信息区整体下沉（2026-09-13 实测逐段漂移表）。
# 精确行距对空段无害（没有字形可裁），独立样张本来就是这些值、观感不变。
_EXACT_PITCH_IDS = frozenset((4, 5, 6, 7, 10, 11, 14, 15, 16, 17, 18, 19, 20, 21, 23))

# 标题与信息区之间的 8 个空白垫段（[14]～[21]）。标题超一行的部分从这里
# 扣回来：标题多一行（如 17 字长题折成两行），这 8 段就各缩 36.8/8 磅，
# 保证信息区在页面上的绝对位置不变（2026-09-13：不扣的话信息区被顶出
# 页面，多出一页空白）。
_TITLE_EXTRA_PT = 36.8  # 标题一行的高度（与 _COVER_PITCH_PT[12] 一致）
_TOP_BLANK_IDS = tuple(range(14, 22))
_TITLE_CHARS_PER_LINE = 37.7  # 版心 415.3 磅 ÷ 11 磅/半角位（标题 sz44＝22 磅字）


def _is_underlined(run) -> bool:
    rPr = run._element.find(qn("w:rPr"))
    if rPr is None:
        return False
    u = rPr.find(qn("w:u"))
    if u is None:
        return False
    return u.get(qn("w:val")) not in ("none", "false", "0")


def _width(text: str) -> int:
    """半角当 1、全角当 2，估个宽度（对齐下划线长度用，不用太精确）。

    例外一：间隔号"·"（U+00B7）虽然编码上在拉丁区，楷体里实际按**全角**渲染，
    按 2 算——不算的话"阿依古丽·买买提"这类名字会贴边换行（2026-09-12 实测）。
    例外二：零宽字符（FEFF 收尾符、零宽空格们）必须按 **0** 算——算宽了会
    把"标签+内容"虚增过悬挂阈值，给标准行误加悬挂缩进（2026-09-13 实测：
    民总这类带字符单位设置的文档把 hanging 换算成"+2 字符"，行首凭空
    缩进 28 磅还折行）。
    """
    total = 0
    for ch in text:
        if ch in ("\ufeff", "\u200b", "\u200c", "\u200d", "\u2060"):
            continue
        if ch in ("\u00b7", "\u30fb"):
            total += 2
        elif ord(ch) > 0x2E7F:
            total += 2
        else:
            total += 1
    return total


def _set_para_text(paragraph, text: str) -> None:
    """改段落文字：只动 w:t，别的一概不碰（项目铁律：绝不用 run.text=…）。"""
    runs = paragraph.runs
    if not runs:
        if text:
            paragraph.add_run(text)
        return
    first = True
    for run in runs:
        ts = run._element.findall(qn("w:t"))
        if not ts:
            continue
        ts[0].text = text if first else ""
        for t in ts[1:]:
            t.text = ""
        first = False


def _fill_field(
    paragraph,
    value: str,
    blank_width: int,
    drop_literal_underscores: bool = False,
    terminator: bool = True,
) -> None:
    """重排"标签＋下划线内容"一行：删旧下划线 run、保留标签、挂一条新下划线 run。

    value 为空 → blank_width 个下划线空格；否则 = 值+两侧垫空格、**居中**在
    这一宽度里。drop_literal_underscores=True 时，把 `___` 字面字符的 run 也
    当旧内容删掉（课程封面就是这种写法）。terminator 控制行尾要不要加零宽
    收尾字符（见上文注释）——居中的行（课程封面）需要；左对齐的行（论文
    信息区）不需要，去掉还省一丝宽度风险。
    """
    src_rPr = None
    for run in list(paragraph.runs):
        is_underline_fmt = _is_underlined(run)
        is_literal = False
        if drop_literal_underscores and "_" in run.text:
            cleaned = run.text.replace("_", "").rstrip()
            if cleaned:
                # 混合 run（如 "：____"）：只抠掉下划线字符，标签留着
                ts = run._element.findall(qn("w:t"))
                if ts:
                    ts[0].text = cleaned
                    for extra in ts[1:]:
                        extra.text = ""
                continue
            is_literal = True
        if is_underline_fmt or is_literal:
            if is_underline_fmt and src_rPr is None:
                src_rPr = run._element.find(qn("w:rPr"))
            others = [
                c for c in run._element if c.tag not in (qn("w:rPr"), qn("w:t"))
            ]
            if others:
                # 里面挂着分栏符之类的元素：只清文字、去掉下划线格式，元素留着
                for t in run._element.findall(qn("w:t")):
                    t.text = ""
                rPr = run._element.find(qn("w:rPr"))
                if rPr is not None:
                    u = rPr.find(qn("w:u"))
                    if u is not None:
                        rPr.remove(u)
            else:
                run._element.getparent().remove(run._element)

    if value:
        v = _width(value)
        left = max(0, (blank_width - v) // 2)
        right = max(1, blank_width - v - left)
        text = " " * left + value + " " * right
    else:
        text = " " * blank_width
    if terminator:
        text += _TERMINATOR

    new_run = paragraph.add_run(text)
    if src_rPr is not None:
        new_run._element.insert(0, copy.deepcopy(src_rPr))
    else:
        # 整行没有现成下划线 run（论文编号 / 课程封面的字面下划线行）：
        # 抄标签的字体字号，再用 python-docx 补下划线（它管 w:u 的 schema 位置）
        label_rPr = None
        for r in paragraph.runs:
            rp = r._element.find(qn("w:rPr"))
            if rp is not None:
                label_rPr = rp
                break
        if label_rPr is not None:
            new_run._element.insert(0, copy.deepcopy(label_rPr))
        new_run.font.underline = True


def _column_width_twips(doc: DocumentObject, section_index: int) -> int:
    """第 section_index 个节里**一栏**的宽度（缇）：版心宽扣掉栏间距后均分。"""
    sect = doc.element.body.findall(".//" + qn("w:sectPr"))[section_index]
    pg_w = int(sect.find(qn("w:pgSz")).get(qn("w:w")))
    mar = sect.find(qn("w:pgMar"))
    text_w = pg_w - int(mar.get(qn("w:left"))) - int(mar.get(qn("w:right")))
    cols = sect.find(qn("w:cols"))
    num = int(cols.get(qn("w:num"), "1")) if cols is not None else 1
    space = int(cols.get(qn("w:space"), "720")) if cols is not None else 0
    return (text_w - space * (num - 1)) // num


def _flush_underline_right(paragraph, col_width_twips: int, *,
                           protect_label_spacing: bool = False) -> None:
    """超长值折行时，让每一行的下划线右端都齐到栏的右边缘。

    做法：整段两端对齐（非末行被撑满到栏宽，下划线跟着到边），末行的行尾垫空
    改成"下划线制表符"跳到栏右边缘（一个右对齐制表位，位置＝栏宽），
    这样末行的线右端与上面各行等齐（用户 2026-09-20：折出的行下划线右端应与
    上一行等长）。一行装得下时制表符也只是把线补到栏边缘，不会折行。

    ⚠️ `protect_label_spacing=True` 时（课程封面用）：标签内部本来就有空格的
    行（"学　　号："这种，靠内部空格把两个字撑开对齐）不用"两端对齐"，改用
    左对齐——两端对齐会把这一行里**所有**空格（含标签内部那几个装饰用的）
    一起拉伸/压缩，标签会被拉变形、位置跑偏（2026-09-22 审计在课程封面上
    验出、实拍复现过）。论文封面的信息行**不受影响**（默认 False，维持
    2026-09-20 定下的行为原样）——那边的行本来就先被强制改成左对齐（见
    `build_thesis_cover_document` 里"信息行改左对齐"那段），`_flush_
    underline_right` 对它们一律用两端对齐是刻意的、已经用真实文档验证过，
    没有这个毛病；课程封面的行没有那道"先改左对齐"的预处理，模板原样是
    居中对齐，两端对齐撞上标签内部空格才会出问题，只在这条路径上收敛处理，
    别不分青红皂白影响论文封面那边已经验证过的效果。
    """
    runs = [r for r in paragraph.runs if _is_underlined(r)]
    if not runs:
        return
    last = runs[-1]
    ts = last._element.findall(qn("w:t"))
    if not ts:
        return
    core = "".join(t.text or "" for t in ts).replace(_TERMINATOR, "").rstrip(" ")
    ts[0].text = core
    ts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for t in ts[1:]:
        t.text = ""
    # 在文字之后追加：制表符（同一 run，沿用下划线格式）＋零宽收尾字符
    tab = OxmlElement("w:tab")
    ts[-1].addnext(tab)
    tail = OxmlElement("w:t")
    tail.text = _TERMINATOR
    tab.addnext(tail)
    if protect_label_spacing:
        label_text = "".join(r.text for r in paragraph.runs if not _is_underlined(r))
        has_space = " " in label_text or "　" in label_text
    else:
        has_space = False
    paragraph.alignment = (WD_ALIGN_PARAGRAPH.LEFT if has_space
                           else WD_ALIGN_PARAGRAPH.JUSTIFY)
    paragraph.paragraph_format.tab_stops.add_tab_stop(
        Twips(col_width_twips), WD_TAB_ALIGNMENT.RIGHT)


def _run_size_pt(paragraph, default: float = 14.0) -> float:
    """该行第一个带字号的 run 的字号（磅）；估宽用（半角位＝字号的一半）。"""
    for r in paragraph.runs:
        sz = r._element.find(qn("w:rPr") + "/" + qn("w:sz"))
        if sz is not None:
            return int(sz.get(qn("w:val"))) / 2.0
    return default


def _label_width_pt(paragraph) -> float:
    """该行提示词（标签）占多宽（磅）——悬挂缩进量就取它。

    信息行是 14 磅字体，所以 1 个半角位 = 7 磅。各行的提示词宽度不一样
    （"姓    名："=70 磅，"专    业:"因为半角冒号只有 63 磅），逐行算。
    """
    text = "".join(r.text for r in paragraph.runs if not _is_underlined(r))
    return _width(text) * 7.0


def _underscore_width(paragraph) -> int:
    """量出这行里字面下划线最多的一串有几个 `_` 字符。"""
    counts = [run.text.count("_") for run in paragraph.runs]
    return max(counts) if max(counts) else 31


def _cut_after(document: DocumentObject, keep_index: int) -> None:
    """只留前 keep_index+1 段；末段若背着节属性，请到 body 级（文档末节的规矩）。"""
    body = document.element.body
    last_el = document.paragraphs[keep_index]._p
    pPr = last_el.find(qn("w:pPr"))
    moved = None
    if pPr is not None:
        moved = pPr.find(qn("w:sectPr"))
        if moved is not None:
            pPr.remove(moved)
    children = list(body)
    for el in children[children.index(last_el) + 1 :]:
        body.remove(el)
    if moved is not None:
        body.append(moved)


def _strip_cover_refs(document: DocumentObject) -> None:
    """封面节清掉模板残留的页眉/页脚引用和"首页不同"标志。

    模板的封面节挂着两个页脚引用（页码“1”和偶数页占位）——它们是给后面的
    前置部分当“链表头”用的，单独抽出来的封面会把这些遗留一起带走：实测
    会把“页眉样式的下边框”牵到封面页上画出一条全宽横线（2026-09-12 用
    PDF 解剖 + 变体实验定位；去掉引用/标志后，线消失）。封面上本来就不
    该有页眉页脚，全清。
    """
    for sec in document.sections:
        s = sec._sectPr
        for el in list(s):
            if el.tag.endswith("}headerReference") or el.tag.endswith("}footerReference"):
                s.remove(el)
        tp = s.find(qn("w:titlePg"))
        if tp is not None:
            s.remove(tp)


def _debug_special_runs(document: DocumentObject, label: str, limit: int) -> None:
    """体检：哪些 run 还挂着文字以外的东西（分栏符/图片…），帮我们确认没误伤。"""
    for i in range(min(limit + 1, len(document.paragraphs))):
        for ri, run in enumerate(document.paragraphs[i].runs):
            extra = [
                c.tag.split("}")[-1]
                for c in run._element
                if c.tag not in (qn("w:rPr"), qn("w:t"))
            ]
            if extra:
                print(f"  [{label}{i:03}] run{ri} 挂有 {extra} 文字={run.text[:16]!r}")


def _widen_label_gap(paragraph) -> None:
    """把提示词里的 4 个半角空格换成 2 个全角空格（宽度都是 28 磅）。

    为什么：半角空格在“新式”模式（compat 15）里渲染 7 磅/个，4 个＝ 28 磅，
    标签合计 70 磅；但老式兼容模式（compat 14，民总.docx 就是）会把它按
    全角算 14 磅/个，标签凭空长 28 磅——整条下划线右移、填写区看着“不居中”
    （2026-09-13 用户实拍 + PDF 坐标对照：线起点 160→188、宽 112→99）。
    全角空格（U+3000）两种模式都渲染 14 磅，2 个＝ 28 磅正好等价，两边一致。
    """
    for run in paragraph.runs:
        if _is_underlined(run) or "    " not in run.text:
            continue
        for t in run._element.findall(qn("w:t")):
            if t.text and "    " in t.text:
                t.text = t.text.replace("    ", "\u3000\u3000")


def _clear_char_unit_indent(paragraph) -> None:
    """补齐段落缺失的缩进属性（全写 0），**不动已存在的直接值**。

    为什么：民总的默认段落样式“正文”自带 **2 字符首行缩进**（字符单位）。
    没有直接缩进属性的封面段落会继承它——行首凭空缩进 28 磅、可用宽度
    变窄、标准行折行（2026-09-13 实测：Word 报告 FirstLineIndent=+28 /
    CharUnitFirstLine=2.0）。

    策略：**只给没写的属性补 0**——模板自己设计的缩进（比如课程封面
    “总分/评分人”两行的首行缩进 1920 twips = 96 磅）必须原样保留，
    清零函数曾把它误杀成"标签贴页边距"（2026-09-13 用户实拍
    “底下那两个歪了”）。成对互斥：firstLine 系列与 hanging 系列只补
    缺的那一对，避免两个同时存在造成渲染歧义。
    """
    pPr = paragraph._p.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        # schema 顺序：spacing 之后、jc 之前
        spacing = pPr.find(qn("w:spacing"))
        jc = pPr.find(qn("w:jc"))
        if spacing is not None:
            spacing.addnext(ind)
        elif jc is not None:
            jc.addprevious(ind)
        else:
            pPr.append(ind)
    has_first_line = (
        ind.get(qn("w:firstLine")) is not None
        or ind.get(qn("w:firstLineChars")) is not None
    )
    has_hanging = (
        ind.get(qn("w:hanging")) is not None
        or ind.get(qn("w:hangingChars")) is not None
    )
    for attr in ("w:left", "w:leftChars"):
        if ind.get(qn(attr)) is None:
            ind.set(qn(attr), "0")
    if not has_hanging:
        for attr in ("w:firstLine", "w:firstLineChars"):
            if ind.get(qn(attr)) is None:
                ind.set(qn(attr), "0")
    if not has_first_line:
        for attr in ("w:hanging", "w:hangingChars"):
            if ind.get(qn(attr)) is None:
                ind.set(qn(attr), "0")


def _pin_zero_spacing(paragraph) -> None:
    """把封面段落的"段前/段后"显式钉成 0（段落自己写了值的不动）。

    为什么必须显式写：封面段落是从官方模板整段搬进**目标文档**的，它们自己
    没有段距时就吃**目标文档的文档默认样式**。官方两个封面模板的文档默认段距
    都是空的（`tools/_probe_cover_template_spacing.py` 实测）⇒ 实际就是 0；
    而空白模板的"源文件"是 python-docx 自带模板，它的 Normal 带 **段后 10 磅**
    ⇒ 封面 24 段各多 10 磅≈240 磅，封面被顶到第二页（用户 2026-09-16 实拍：
    "生成考核空白模板时，插入的封面错误的分到了两页。本科毕业论文也是"）。
    真文档（民总等）的默认样式没这个值，所以只有空白模板会犯。
    课程封面 [10] 自己写了 `beforeLines=50`，那种不再动。
    """
    pPr = paragraph._p.get_or_add_pPr()
    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        # schema 顺序：spacing 在 ind 前面
        ind = pPr.find(qn("w:ind"))
        if ind is not None:
            ind.addprevious(spacing)
        else:
            pPr.append(spacing)
    for attr in ("w:before", "w:after"):
        if spacing.get(qn(attr)) is None:
            spacing.set(qn(attr), "0")


def _pin_line_pitch(
    paragraph, pt: float, rule: WD_LINE_SPACING = WD_LINE_SPACING.AT_LEAST
) -> None:
    """钉死段落行距（绝对磅值）并关掉网格对齐。

    绝对 atLeast 不受文档级兼容设置影响；关 snapToGrid 防止被章节的
    docGrid 二次拉伸（网格行高 15.6 磅，小值会被吸上去）。两招合用，
    自然行距再短也不缩水、再长也不会被网格拉松。
    空白垫段用 EXACTLY：只有"精确"才能把行高压得比自然值小（标题多出的
    行要从空白里扣回来，保证信息区位置固定）。
    """
    pf = paragraph.paragraph_format
    pf.line_spacing_rule = rule
    pf.line_spacing = Pt(pt)
    pPr = paragraph._p.get_or_add_pPr()
    snap = pPr.find(qn("w:snapToGrid"))
    if snap is None:
        snap = OxmlElement("w:snapToGrid")
        # schema 顺序：snapToGrid 必须紧跟在 spacing 前面
        spacing = pPr.find(qn("w:spacing"))
        ind = pPr.find(qn("w:ind"))
        anchor = spacing if spacing is not None else ind
        if anchor is not None:
            anchor.addprevious(snap)
        else:
            pPr.append(snap)
    snap.set(qn("w:val"), "0")


# ------------------------------------------------------------------ 两种封面


def build_thesis_cover_document(
    fields: CoverFields, hanging_indent: bool = True, verbose: bool = False
) -> DocumentObject:
    """论文封面文档（独立 docx：文字区 + 两栏信息区，共 2 个节）。

    hanging_indent=True（默认，用户 2026-09-12 拍板）时给信息行加**悬挂
    缩进**：缩进量=该行提示词宽度——首行完全不变，超长值换行后的续行对齐到
    提示词右侧（值起点正下方）。接入流水线后由预设「封面」小节的"超长值
    悬挂缩进"开关控制（默认要）。verbose=True 打印非文字元素体检信息。
    """
    doc = Document(str(verify_official(THESIS_TEMPLATE)))
    if verbose:
        _debug_special_runs(doc, "论文", 33)
    paras = doc.paragraphs

    # 论文编号：不加下划线（原件如此——这行是学校事后打印编号的位置；
    # 实测右对齐行尾的"空格下划线"会被 Word 当行尾空格吞掉，加了也不显示）
    # 2026-09-19 用户拍板：没填编号时，冒号后按四个汉字宽度直接留白（不加下划线）；
    # 垫零宽收尾字符保宽度，否则行尾空白会被 Word 裁掉、留白不生效。
    number_tail = fields.number or ("\u3000" * 4 + _TERMINATOR)
    _set_para_text(paras[0], "论文编号：" + number_tail)
    _set_para_text(paras[12], fields.title)  # 题目（楷体_GB2312 二号）
    _set_para_text(paras[13], fields.subtitle)  # 副标题（楷体四号）

    # 每行的字面值先过一遍书写规则（本模块上半部分）：
    # 学号原样；姓名两个字空格；指导教师"姓名 职称"；日期原样（留空就是空）
    values = {
        24: fields.student_id,
        25: fields.college,
        26: fields.grade,
        27: resolve_date(fields.date),
        28: format_person_name(fields.student_name),
        29: fields.major,
        30: format_advisor(fields.advisor_name, fields.advisor_title),
        31: fields.score,
    }
    for idx in values:
        _widen_label_gap(paras[idx])
    for idx, value in values.items():
        # terminator=True：行尾垫的全角/半角空格后面跟一个零宽字符，否则
        # compat 14 会把行尾空格连同下划线一起裁短，值和线左右不对称
        # （2026-09-13 实拍：右侧线只剩 8 磅，应该是 21 磅）。
        _fill_field(paras[idx], value, THESIS_BLANK_WIDTH, terminator=True)

    # 信息行改**左对齐**：模板是两端对齐，两栏窄栏里一旦行贴近栏宽，有的
    # 渲染器会拉伸行内所有空格（标签里"学    号"的间隔被撑大——用户 2026-09-12
    # 截图实拍：有内空格的标签都被拉长，"完成日期"这种没内空格的却正常）。
    # 单行左对齐观感完全一样，但任何渲染器都不可能再拉伸空格。
    for idx in range(24, 32):
        paras[idx].alignment = WD_ALIGN_PARAGRAPH.LEFT

    # 缩进清零：封死民总默认样式“正文”的 2 字符首行缩进继承（[13] 副标题
    # 的 700 字符悬挂是模板刻意设计，跳过）
    for idx in range(0, 24):
        if idx != 13:
            _clear_char_unit_indent(paras[idx])
    for idx in range(24, 32):
        _clear_char_unit_indent(paras[idx])

    if hanging_indent:
        # 悬挂缩进：left=提示词宽、first_line 负同值 → 首行不动，续行对齐到
        # 值起点（"悬挂到提示词以右"效果）。但**只给真装不下的行加**，标准行
        # 加了反而在 compat 14 文档（民总.docx）里被压窄折行（见 _HANGING_MIN_PT）
        for idx in range(24, 32):
            label_pt = _label_width_pt(paras[idx])
            fill_pt = sum(
                _width(r.text) for r in paras[idx].runs if _is_underlined(r)
            ) * 7.0
            if label_pt + fill_pt > _HANGING_MIN_PT:
                pf = paras[idx].paragraph_format
                pf.left_indent = Pt(label_pt)
                pf.first_line_indent = Pt(-label_pt)
                _flush_underline_right(paras[idx], _column_width_twips(doc, 1))

    for idx, pitch in _COVER_PITCH_PT.items():
        rule = (
            WD_LINE_SPACING.EXACTLY
            if idx in _EXACT_PITCH_IDS
            else WD_LINE_SPACING.AT_LEAST
        )
        _pin_line_pitch(paras[idx], pitch, rule=rule)

    # 段前/段后钉 0：见 `_pin_zero_spacing` 的注释（空白模板里那段 10 磅段后
    # 会把封面顶到第二页）。
    for idx in range(0, 33):
        _pin_zero_spacing(paras[idx])

    # 标题多行补偿：多出的行高从标题下方的空白垫段里均摊扣回，信息区
    # 的绝对位置保持不动（长题的两种情况：17 字折两行、20+ 字折三行）。
    title_units = _width(fields.title)
    extra_lines = max(0, math.ceil(title_units / _TITLE_CHARS_PER_LINE - 1e-9) - 1)
    if extra_lines:
        base_pt = _COVER_PITCH_PT[14]
        per_blank = max(4.0, base_pt - extra_lines * _TITLE_EXTRA_PT / len(_TOP_BLANK_IDS))
        for idx in _TOP_BLANK_IDS:
            _pin_line_pitch(paras[idx], per_blank, rule=WD_LINE_SPACING.EXACTLY)

    _cut_after(doc, 33)
    _strip_cover_refs(doc)

    # [032][033] 是模板留下的两行空段，纯占位。删掉它们：值超长换行时
    # （如"国际法学院/涉外法治学院"）信息区往下长，这两行空白会先被挤出页，
    # 多出一整张空白第二页（2026-09-12 5 号压力样例实拍）。删掉后正常值的
    # 观感不变，超长值多两行容错。
    for idx in (33, 32):
        p = doc.paragraphs[idx]
        p._element.getparent().remove(p._element)

    return doc


# 课程封面顶部空行垫 [00..09] 的标准行距（磅，样张实录）。为什么单独钉死：
# 顶部是一串空行垫（两张校徽/书法图是浮动锚，不占行高），目标文档的行高
# 算法不同时每段多出 ~4.5 磅，十段累积 70 磅把信息区整体推下去，末尾空段
# 被挤出页（ 2026-09-13 实测：民总里成品 17 页 → 钉死后 16 页）。
_COURSE_PITCH_PT = {
    0: 15.7, 1: 15.8, 2: 15.0, 3: 15.8, 4: 15.7, 5: 15.7, 6: 15.8, 7: 15.0,
    8: 15.8, 9: 21.7,
}


def build_course_cover_document(
    fields: CoverFields, verbose: bool = False
) -> DocumentObject:
    """课程考核封面文档（独立 docx，1 个节）。

    它带校徽/书法图片和页眉（顶部标题在页眉里）——拼进别的文档由
    `prepend_cover_document` 负责把图片和页眉部件跨包迁移。
    """
    doc = Document(str(verify_official(COURSE_TEMPLATE)))
    if verbose:
        _debug_special_runs(doc, "课程", 34)
    paras = doc.paragraphs

    # 任课教师**默认只写姓名、不带职称**（用户 2026-09-12 拍板）——要带头衔
    # 得以后加开关，别把"李四 副教授"整串塞进 teacher_name
    values = {
        10: fields.course_name,
        11: fields.course_code,
        12: fields.course_seq,
        13: format_person_name(fields.teacher_name),
        15: format_person_name(fields.student_name),
        16: fields.student_id,
    }
    for idx, value in values.items():
        _fill_field(
            paras[idx], value, _underscore_width(paras[idx]), drop_literal_underscores=True
        )

    # 值长到一行装不下的：折出的行下划线右端与上一行齐（同论文封面，见 _flush_underline_right）
    col_pt = _column_width_twips(doc, 0) / 20.0
    for idx, value in values.items():
        if not value:
            continue
        size = _run_size_pt(paras[idx])
        label = "".join(r.text for r in paras[idx].runs if not _is_underlined(r))
        if (_width(label) + _width(value) + 1) * size / 2 > col_pt:
            _flush_underline_right(paras[idx], _column_width_twips(doc, 0),
                                   protect_label_spacing=True)

    for idx in (22, 23):  # 总分 / 评分人：阅卷老师填，两种样例都留空
        _fill_field(
            paras[idx], "", _underscore_width(paras[idx]), drop_literal_underscores=True
        )

    # 缩进清零：同论文封面——封死目标文档默认样式的字符缩进继承
    # （民总的「正文」样式自带 2 字符首行缩进，会让居中行整体偏移）
    for idx in range(0, 24):
        _clear_char_unit_indent(paras[idx])

    # 顶部空行垫钉死（见 _COURSE_PITCH_PT 注释）——EXACTLY，两种模式都精确
    for idx, pitch in _COURSE_PITCH_PT.items():
        _pin_line_pitch(paras[idx], pitch, rule=WD_LINE_SPACING.EXACTLY)

    # 段前/段后钉 0：同论文封面（见 `_pin_zero_spacing`）
    for idx in range(0, 24):
        _pin_zero_spacing(paras[idx])

    _cut_after(doc, 24)

    # “评分区域（由阅卷老师红笔填写）：”这行按用户要求删掉（2026-09-12）
    p = doc.paragraphs[19]
    p._element.getparent().remove(p._element)

    return doc


def build_year_cover_document(
    fields: CoverFields, hanging_indent: bool = True, verbose: bool = False
) -> DocumentObject:
    """学年论文封面：与毕业论文封面**唯一**的区别是标题行里的"毕 业"换成
    "学 年"（用户 2026-09-19："学年模板和毕业论文模板的唯一区别是，封面上的
    '毕业论文'换成'学年论文'"）。所以直接拿论文封面改字，字段、版式全同。"""
    doc = build_thesis_cover_document(fields, hanging_indent=hanging_indent, verbose=verbose)
    for p in doc.paragraphs[:20]:
        if "毕" in p.text and "业" in p.text:
            # 只动 w:t（项目铁律），"毕 业"中间的空格、分散在各 run 的情况都照样换
            for t in p._p.iter(qn("w:t")):
                if t.text:
                    t.text = t.text.replace("毕", "学").replace("业", "年")
            break
    return doc


# 十星封面：官方"封面范例"文件里的段落（不含表格内的段落）：
#   [00]      浮动表格"论文编号"的锚点段
#   [05][06]  两行固定字样（"中国政法大学“中伦杯”""第 N 届“学术十星”论文大赛参赛论文"）
#   [12]      论文题目（楷体 二号，居中）
#   [13]      副标题（居右；四号楷体）
# 其余 [07]~[11] 与 [01]~[04] 是撑位置的空段，[14] 起到页末全是空段。
_TENSTAR_EDITION_IDX = 6           # 第二行固定字样（含届数）
_TENSTAR_TITLE_IDX = 12
_TENSTAR_SUBTITLE_IDX = 13
_TENSTAR_SUBTITLE_PT = 14.0        # 四号
_TENSTAR_SUBTITLE_PITCH_PT = 20.0

# 各段的"标准行距"（磅）：官方范例在 Word 里的实测位置（[00] 15.8，[01]~[12]
# 每段 31.3，即网格两行）。跟论文封面同一个理由必须钉死：范例的空段行高靠
# 文档自己的默认样式和网格算出来，拼进别的文档就会变，题目会跟着漂。
_TENSTAR_PITCH_PT = {0: 15.8, **{i: 31.3 for i in range(1, 13)}}


def _tenstar_template() -> pathlib.Path:
    """reference/ 里的十星封面范例。文件名带届数（"第二十四届"），每年会换，
    所以按关键词找；跳过 Word 打开原文件时留下的 `~$` 锁文件。"""
    for p in sorted((ROOT / "reference").glob("*学术十星*封面*.docx")):
        if not p.name.startswith("~$"):
            return p
    raise FileNotFoundError("reference/ 里找不到学术十星封面范例（文件名应含「学术十星」「封面」）")


def _set_run_font(run, size_pt: float, font: str) -> None:
    """给新建的 run 写字体：中西文都用 `font`，字号 `size_pt`（只写直接属性）。"""
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia"):
        rFonts.set(qn(attr), font)
    rFonts.set(qn("w:hint"), "eastAsia")
    run.font.size = Pt(size_pt)


def build_tenstar_cover_document(
    fields: CoverFields, verbose: bool = False
) -> DocumentObject:
    """学术十星封面文档（独立 docx，1 个节）：官方"封面范例"只填题目和副标题。

    - 封面上**不放个人信息**（规范：封面及全文不能出现作者姓名、院系等），
      "论文编号"格保持空白；固定字样（含届数）照范例原样保留；
    - 页眉里的文字水印随封面节的页眉引用一并摘掉（`_strip_cover_refs`），
      不带进成品；
    - 范例末尾一大片空段裁掉——它们只是把页面撑满，拼进别的文档反而会撑出
      多余的一页。
    """
    doc = Document(str(verify_official(_tenstar_template())))
    if verbose:
        _debug_special_runs(doc, "十星", 14)
    paras = doc.paragraphs

    if fields.session.strip():
        # 固定字样第二行"第二十四届“学术十星”论文大赛参赛论文"里的届数换成用户填的；
        # 该段的几个 run 字体字号都一样，整段文字写进第一个 run 即可
        _set_para_text(paras[_TENSTAR_EDITION_IDX],
                       f"第{fields.session.strip()}届“学术十星”论文大赛参赛论文")
    _set_para_text(paras[_TENSTAR_TITLE_IDX], fields.title)
    if fields.subtitle:
        run = paras[_TENSTAR_SUBTITLE_IDX].add_run(fields.subtitle)
        _set_run_font(run, _TENSTAR_SUBTITLE_PT, "楷体")

    # 裁掉副标题之后的空段（保留 body 级的节属性）
    body = doc.element.body
    last = paras[_TENSTAR_SUBTITLE_IDX]._p
    children = list(body)
    for el in children[children.index(last) + 1:]:
        if el.tag != qn("w:sectPr"):
            body.remove(el)

    for idx in range(0, _TENSTAR_SUBTITLE_IDX + 1):
        _clear_char_unit_indent(paras[idx])
        _pin_zero_spacing(paras[idx])
    for idx, pitch in _TENSTAR_PITCH_PT.items():
        _pin_line_pitch(paras[idx], pitch, rule=WD_LINE_SPACING.EXACTLY)
    _pin_line_pitch(paras[_TENSTAR_SUBTITLE_IDX], _TENSTAR_SUBTITLE_PITCH_PT,
                    rule=WD_LINE_SPACING.EXACTLY)

    # "论文编号"浮动表格：边框都是直接写在表上的，摘掉它对表样式的引用（范例里
    # 的样式 id 在别的文档里可能是完全不同的样式）；格内文字没写字号，会继承
    # 目标文档的正文样式——钉成五号宋体，缩进/段距清零。
    for tbl in doc.element.body.iter(qn("w:tbl")):
        style = tbl.find(qn("w:tblPr") + "/" + qn("w:tblStyle"))
        if style is not None:
            style.getparent().remove(style)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _clear_char_unit_indent(p)
                    _pin_zero_spacing(p)
                    _pin_line_pitch(p, 15.8)
                    for run in p.runs:
                        _set_run_font(run, 10.5, "宋体")

    _strip_cover_refs(doc)
    return doc


# ------------------------------------------------------------------ 插入流水线


# 图片引用的两个命名空间（DrawingML 的 a:blip / VML 的 v:imagedata）
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_V_NS = "urn:schemas-microsoft-com:vml"


def _migrate_image_refs(target: DocumentObject, owner_part, element) -> None:
    """把 element（已搬进 target）里的图片引用从 owner_part 搬到 target 包。

    封面带图（如课程封面的校徽、书法题名）时，图内嵌在封面的包里：直接挪
    元素后 r:embed 指向的关系在 target 里悬空/错位。做法 = 把图片本体递进
    target（get_or_add_image_part 自带去重），再把元素里的 rId 改成新的。
    """
    for blip in element.iter(f"{{{_A_NS}}}blip"):
        rid = blip.get(qn("r:embed"))
        if not rid or rid not in owner_part.related_parts:
            continue
        src = owner_part.related_parts[rid]
        new_rid, _ = target.part.get_or_add_image(io.BytesIO(src.blob))
        blip.set(qn("r:embed"), new_rid)
    for img in element.iter(f"{{{_V_NS}}}imagedata"):
        rid = img.get(qn("r:id"))
        if not rid or rid not in owner_part.related_parts:
            continue
        src = owner_part.related_parts[rid]
        new_rid, _ = target.part.get_or_add_image(io.BytesIO(src.blob))
        img.set(qn("r:id"), new_rid)


def _copy_header_into(target: DocumentObject, header_part, src_part) -> None:
    """把源页眉/页脚部件的内容整块搬进 target 的新部件（图片关系重建）。

    还要做**样式去撞名**：源部件的段落 pStyle 指向源文档的样式 Id（课程封面
    页眉是 "3"），而目标文档的同名 Id 可能是完全不同的样式——民总里 "3" =
    「二级标题」（段前段后各 18 磅），搬过去后页眉变胖、把正文顶下去 26 磅
    （2026-09-13 实测）。pStyle 一律摘掉（源里居中是直接格式、字号也是，不会
    丢观感），段落缩进清零、段前段后归零，行高就回到源位置。
    """
    header_part._element = copy.deepcopy(src_part._element)
    for blip in header_part._element.iter(f"{{{_A_NS}}}blip"):
        rid = blip.get(qn("r:embed"))
        if not rid or rid not in src_part.related_parts:
            continue
        src_img = src_part.related_parts[rid]
        image_part = target.part.package.get_or_add_image_part(io.BytesIO(src_img.blob))
        new_rid = header_part.relate_to(image_part, RT.IMAGE)
        blip.set(qn("r:embed"), new_rid)

    for p in header_part._element.iter(qn("w:p")):
        pPr = p.find(qn("w:pPr"))
        if pPr is None:
            continue
        pstyle = pPr.find(qn("w:pStyle"))
        if pstyle is not None:
            pPr.remove(pstyle)
        spacing = pPr.find(qn("w:spacing"))
        if spacing is None:
            spacing = OxmlElement("w:spacing")
            anchor = pPr.find(qn("w:ind")) if pPr.find(qn("w:ind")) is not None else pPr.find(qn("w:jc"))
            if anchor is not None:
                anchor.addprevious(spacing)
            else:
                pPr.insert(0, spacing)
        spacing.set(qn("w:before"), "0")
        spacing.set(qn("w:after"), "0")
        _clear_char_unit_indent(Paragraph(p, None))


def _force_next_page(sectPr) -> None:
    """把节的开始方式强制为"另起一页"（封面拼进来后，原第一节必须换页）。"""
    t = sectPr.find(qn("w:type"))
    if t is None:
        t = OxmlElement("w:type")
        # type 必须排在页眉页脚引用之后、pgSz 之前，不能无脑 append
        for tag in ("w:pgSz", "w:pgMar", "w:pgNumType", "w:cols", "w:titlePg", "w:docGrid"):
            anchor = sectPr.find(qn(tag))
            if anchor is not None:
                anchor.addprevious(t)
                break
        else:
            sectPr.append(t)
    t.set(qn("w:val"), "nextPage")


def prepend_cover_document(target: DocumentObject, cover: DocumentObject) -> None:
    """把封面文档的全部内容插到 target 文档的最前面（论文/课程通用）。

    封面文档是独立文档：它"末节"的 sectPr 挂在 body 级。拼进 target 后它
    变成**中间节**——中间节的 sectPr 必须挂在段落身上，所以先把它请到一个
    新空段落里，再把封面所有段落按顺序挪到 target 第一个元素前面。
    另外把 target 原来的第一节（现在排在封面后面了）强制改成"另起一页"。

    跨包迁移（课程封面必须，论文封面也要跑一遍空检查）：
    - 正文里的图片（校徽/书法）→ 把图本体递进 target 包，rId 改新的；
    - 封面节的页眉引用（课程封面标题在页眉里）→ 复制部件内容到 target 新
      页眉部件（含页眉里的图片）；**页脚引用直接移除**（封面不要页码）。
    - 没有任何页眉引用的节（论文封面）→ 显式挂空页眉，段落样式用正文
      （否则 Word 自造的空页眉样式是"页眉"，会被页眉线边框画出一条线）。
    """
    cover_body = cover.element.body
    target_body = target.element.body
    # 在拼接**之前**数好封面节数：拼接会把元素全部搬走，cover.sections 会变空
    cover_section_count = len(cover.sections)

    # 0) 拼接前记下封面各节的眉脚引用计划（rId 要对照封面包的 rels 解析）
    ref_plans = []
    for sec in cover.sections:
        plan = []
        for el in sec._sectPr:
            if el.tag.endswith("}headerReference") or el.tag.endswith("}footerReference"):
                plan.append(
                    (
                        el.get(qn("w:type")) or "default",
                        el.tag.split("}")[-1],
                        el.get(qn("r:id")),
                    )
                )
        ref_plans.append(plan)

    # 1) target 原来的第一节：找到它挂 sectPr 的段落（都没有就是 body 级），
    #    强制 nextPage —— 封面之后必须另起一页
    target_sect = None
    for el in target_body:
        if el.tag == qn("w:p"):
            pPr = el.find(qn("w:pPr"))
            if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
                target_sect = pPr.find(qn("w:sectPr"))
                break
    if target_sect is None:
        target_sect = target_body.find(qn("w:sectPr"))
    if target_sect is not None:
        _force_next_page(target_sect)

    # 2) 封面末节的 sectPr（body 级）请到新空段落里
    sect = cover_body.find(qn("w:sectPr"))
    if sect is not None:
        cover_body.remove(sect)
    elements = list(cover_body)
    if sect is not None:
        anchor = OxmlElement("w:p")
        pPr = OxmlElement("w:pPr")
        pPr.append(sect)
        anchor.append(pPr)
        elements.append(anchor)

    # 3) 全部挪到 target 第一个元素前面
    first = target_body[0]
    for el in elements:
        first.addprevious(el)

    # 3.5) 搬进 target 的元素里如果有图片引用，把图片一并搬过来
    for el in elements:
        _migrate_image_refs(target, cover.part, el)

    # 4) 封面各节的页眉/页脚处理
    for idx, sec in enumerate(target.sections[:cover_section_count]):
        sectPr = sec._sectPr
        raw = ref_plans[idx] if idx < len(ref_plans) else []
        # 先摘光旧引用：那些 rId 指向封面包，在 target 里是悬空的错误引用
        for el in list(sectPr):
            if el.tag.endswith("}headerReference") or el.tag.endswith("}footerReference"):
                sectPr.remove(el)

        header_infos = [(t, rid) for (t, tag, rid) in raw if tag == "headerReference"]
        if not header_infos:
            # 无页眉引用（论文封面）：显式空页眉 + 正文样式
            header = sec.header
            header.is_linked_to_previous = False
            for paragraph in header.paragraphs:
                paragraph.style = target.styles["Normal"]
            continue

        for type_attr, rid in header_infos:
            src_part = cover.part.related_parts.get(rid)
            if src_part is None:
                continue
            api = {
                "default": "header",
                "first": "first_page_header",
                "even": "even_page_header",
            }.get(type_attr, "header")
            header = getattr(sec, api)
            header.is_linked_to_previous = False
            _copy_header_into(target, header._definition, src_part)


def add_thesis_cover(
    target: DocumentObject,
    fields: CoverFields,
    hanging_indent: bool = True,
    verbose: bool = False,
) -> None:
    """流水线一键操作：造论文封面 → 插到 target 最前面。"""
    cover = build_thesis_cover_document(
        fields, hanging_indent=hanging_indent, verbose=verbose
    )
    prepend_cover_document(target, cover)


def add_course_cover(
    target: DocumentObject,
    fields: CoverFields,
    verbose: bool = False,
) -> None:
    """流水线一键操作：造课程考核封面 → 插到 target 最前面。

    与论文封面的区别：保留封面自己的页眉（顶部标题在页眉里），校徽/书法
    图片随内容一起跨包迁移（见 prepend_cover_document）。
    """
    cover = build_course_cover_document(fields, verbose=verbose)
    prepend_cover_document(target, cover)


def add_year_cover(
    target: DocumentObject,
    fields: CoverFields,
    hanging_indent: bool = True,
    verbose: bool = False,
) -> None:
    """流水线一键操作：造学年论文封面 → 插到 target 最前面。"""
    cover = build_year_cover_document(
        fields, hanging_indent=hanging_indent, verbose=verbose
    )
    prepend_cover_document(target, cover)


def add_tenstar_cover(
    target: DocumentObject,
    fields: CoverFields,
    verbose: bool = False,
) -> None:
    """流水线一键操作：造学术十星封面 → 插到 target 最前面。"""
    cover = build_tenstar_cover_document(fields, verbose=verbose)
    prepend_cover_document(target, cover)


def add_cover(
    target: DocumentObject,
    cover_type: str,
    fields: CoverFields,
    hanging_indent: bool = True,
) -> None:
    """按封面类型键（`config.schema.COVER_TYPE_BY_FORM` 的值）拼封面。

    认不出的类型报错，**不**悄悄退回论文封面——静默换一种封面比报错更难查。
    `hanging_indent` 只有论文/学年论文封面有这个选项。"""
    if cover_type == "thesis":
        add_thesis_cover(target, fields, hanging_indent=hanging_indent)
    elif cover_type == "year":
        add_year_cover(target, fields, hanging_indent=hanging_indent)
    elif cover_type == "course":
        add_course_cover(target, fields)
    elif cover_type == "tenstar":
        add_tenstar_cover(target, fields)
    else:
        raise ValueError(f'不认识的封面类型："{cover_type}"（应为 thesis/year/course/tenstar）')
