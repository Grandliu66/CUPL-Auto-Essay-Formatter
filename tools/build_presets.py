"""生成 presets/ 下的预设 Excel 配置表。

用法：
    .venv\\Scripts\\python.exe tools\\build_presets.py

标题识别规则（哪个正则/锚点对应哪个部件）不纳入此表，仍是工具内部
判定机制；但"某部件是否启用"已作为用户可调参数纳入"结构部件"小节，
默认全部设为"否"（可选）——是否启用是排版偏好，不该由工具替用户判断。

"备注"列只是留给人自己写笔记用的，工具读取时不看这一列，默认是空的；
"值"这类数据格只放纯数字/固定写法，单位已经写进"项目"名字或表头里了。
"""

from __future__ import annotations

import pathlib

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# 「封面形态」的合法值。故意不 import `essay_tool.config.schema`——这个脚本
# 是独立跑的历史产物（`python tools\build_presets.py`，脚本目录不在 sys.path
# 里、导不进本包），名单跟 `schema.COVER_FORM_NAMES` 保持一致就行。
COVER_FORM_NAMES = ("毕业论文封面", "学年论文封面", "课程考核封面", "学术十星封面")

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT_DIR = ROOT / "presets"

MAX_COLS = 19

TITLE_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(vertical="center", wrap_text=True)
CENTER = Alignment(vertical="center", horizontal="center", wrap_text=True)

YN = ["是", "否"]
ALIGN = ["两端对齐", "居中", "左对齐", "右对齐"]
LINE_RULE = ["多倍行距（倍）", "固定值（磅）"]
FONT_SIZES = [
    "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三",
    "四号", "小四", "五号", "小五", "六号", "小六", "七号", "八号",
]
INDENT_UNITS = ["字符", "厘米", "-"]
SPACING_UNITS = ["磅"]
PAPER_SIZES = ["A4", "A3", "B5", "自定义"]
DIRECTIONS = ["纵向", "横向"]
PAGE_NUMBER_FORMATS = ["1,2,3", "I,II,III", "i,ii,iii", "A,B,C", "a,b,c"]

COL_WIDTHS = [10, 16, 12, 16, 8, 6, 10, 10, 8, 10, 8, 8, 8, 8, 8, 12, 8, 10, 22]

NOTE_TEXT = (
    "说明\n"
    "1. 该Excel是可供工具直接读取的配置文件来源，所有的填写都必须按格式进行，否则将无法为工具正确读取。\n"
    '2. "备注"栏仅为分享和自我提醒之用，无法为工具理解。'
)


def intro_note(ws, row: int, text: str) -> int:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=MAX_COLS)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(size=10, color="595959")
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 50
    return row + 1


def section_title(ws, row: int, text: str) -> int:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=MAX_COLS)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(bold=True, size=13, color="FFFFFF")
    cell.fill = TITLE_FILL
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 24
    return row + 1


def header_row(ws, row: int, headers: list[str]) -> int:
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = BORDER
    return row + 1


def data_rows(ws, row: int, rows: list[list], center_cols: set[int]) -> int:
    for offset, data in enumerate(rows):
        r = row + offset
        for c, val in enumerate(data, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.border = BORDER
            cell.alignment = CENTER if c in center_cols else WRAP
    return row + len(rows)


def add_block(ws, row: int, title: str, headers: list[str], rows: list[list],
              center_cols: set[int] = frozenset()) -> tuple[int, int]:
    """写入一个小节，返回 (数据起始行, 数据结束行+1)。"""
    row = section_title(ws, row, title)
    row = header_row(ws, row, headers)
    data_start = row
    row = data_rows(ws, row, rows, center_cols)
    return data_start, row


def _add_validation(ws, col: int, row_start: int, row_end: int, formula1: str,
                    allow_blank: bool = False):
    if row_start >= row_end:
        return
    dv = DataValidation(type="list", formula1=formula1, allow_blank=allow_blank)
    ws.add_data_validation(dv)
    letter = get_column_letter(col)
    dv.add(f"{letter}{row_start}:{letter}{row_end - 1}")


def add_dropdown(ws, col: int, row_start: int, row_end: int, options: list[str],
                allow_blank: bool = False):
    """选项内联写进公式；选项本身不能含逗号，否则会被 Excel 当成分隔符拆散。"""
    _add_validation(ws, col, row_start, row_end, '"' + ",".join(options) + '"',
                    allow_blank=allow_blank)


def add_dropdown_range(ws, col: int, row_start: int, row_end: int, ref: str):
    """选项本身带逗号（比如"1,2,3"）时用这个：选项另外写在隐藏表里，下拉引用那个区域。"""
    _add_validation(ws, col, row_start, row_end, ref)


def build(ws, page_number_format_range: str):
    row = 1
    row = intro_note(ws, row, NOTE_TEXT)
    row += 1

    # ------------------------------------------------------------ 页面设置
    headers = ["项目", "值", "备注"]
    rows = [
        ["纸张", "A4", ""],
        ["自定义宽（厘米）", "", ""],
        ["自定义高（厘米）", "", ""],
        ["纸张方向", "纵向", ""],
        ["上边距（厘米）", "2.5", ""],
        ["下边距（厘米）", "2.5", ""],
        ["左边距（厘米）", "2.0", ""],
        ["右边距（厘米）", "2.0", ""],
        ["装订线（厘米）", "0", ""],
        ["页眉距边（厘米）", "1.5", ""],
        ["页脚距边（厘米）", "1.75", ""],
    ]
    data_start, data_end = add_block(ws, row, "页面设置", headers, rows, center_cols={2, 3})
    add_dropdown(ws, 2, data_start, data_start + 1, PAPER_SIZES)  # 纸张
    add_dropdown(ws, 2, data_start + 3, data_start + 4, DIRECTIONS)  # 纸张方向
    row += 1 + 1 + len(rows) + 1  # 标题行+表头行+数据行+空行

    # ----------------------------------------------------------------- 样式
    headers = [
        "编号", "名称", "中文字体", "西文字体", "字号", "加粗", "对齐",
        "首行缩进单位", "首行缩进值", "左缩进单位", "左缩进值",
        "段前单位", "段前值", "段后单位", "段后值",
        "行距类型", "行距值", "段前分页", "备注",
    ]
    rows = [
        ["S01", "正文", "宋体", "Times New Roman", "四号", "否", "两端对齐", "字符", "2", "-", "0", "磅", "0", "磅", "0", "固定值（磅）", "20", "否", ""],
        ["S02", "一级标题", "宋体", "宋体", "小三", "是", "居中", "-", "0", "-", "0", "磅", "24", "磅", "24", "多倍行距（倍）", "1", "是", ""],
        ["S03", "二级标题", "宋体", "宋体", "四号", "是", "左对齐", "-", "0", "-", "0", "磅", "18", "磅", "18", "固定值（磅）", "20", "否", ""],
        ["S04", "三级标题", "宋体", "宋体", "四号", "否", "左对齐", "-", "0", "厘米", "0.85", "磅", "12", "磅", "6", "固定值（磅）", "20", "否", ""],
        ["S05", "四级标题", "楷体", "楷体", "四号", "否", "左对齐", "-", "0", "厘米", "0.85", "磅", "6", "磅", "6", "固定值（磅）", "20", "否", ""],
        ["S06", "中文题目", "黑体", "黑体", "三号", "否", "居中", "-", "0", "-", "0", "磅", "24", "磅", "24", "多倍行距（倍）", "1", "否", ""],
        ["S07", "摘要标签", "黑体", "黑体", "四号", "是", "左对齐", "-", "0", "-", "0", "磅", "18", "磅", "18", "多倍行距（倍）", "1", "否", ""],
        ["S08", "摘要正文", "宋体", "宋体", "小四", "否", "两端对齐", "字符", "2", "-", "0", "磅", "0", "磅", "0", "固定值（磅）", "20", "否", ""],
        ["S09", "关键词标签", "黑体", "黑体", "四号", "是", "左对齐", "-", "0", "-", "0", "磅", "18", "磅", "18", "多倍行距（倍）", "1", "否", ""],
        ["S10", "关键词内容", "宋体", "宋体", "小四", "否", "左对齐", "-", "0", "-", "0", "磅", "18", "磅", "18", "多倍行距（倍）", "1", "否", ""],
        ["S11", "英文题目", "Times New Roman", "Times New Roman", "三号", "是", "居中", "-", "0", "-", "0", "磅", "24", "磅", "24", "多倍行距（倍）", "1", "是", ""],
        ["S12", "英文摘要标签", "Times New Roman", "Times New Roman", "四号", "是", "左对齐", "-", "0", "-", "0", "磅", "18", "磅", "18", "多倍行距（倍）", "1", "否", ""],
        ["S13", "英文摘要正文", "Times New Roman", "Times New Roman", "小四", "否", "两端对齐", "字符", "2", "-", "0", "磅", "0", "磅", "0", "固定值（磅）", "20", "否", ""],
        ["S14", "目录标题", "黑体", "黑体", "三号", "否", "居中", "-", "0", "-", "0", "磅", "24", "磅", "24", "多倍行距（倍）", "1", "是", ""],
        ["S15", "目录一级", "宋体", "宋体", "小四", "否", "左对齐", "-", "0", "-", "0", "磅", "0", "磅", "0", "固定值（磅）", "20", "否", ""],
        ["S16", "目录二级", "宋体", "宋体", "小四", "否", "左对齐", "-", "0", "字符", "1", "磅", "0", "磅", "0", "固定值（磅）", "20", "否", ""],
        ["S17", "目录三级", "宋体", "宋体", "小四", "否", "左对齐", "-", "0", "字符", "2", "磅", "0", "磅", "0", "固定值（磅）", "20", "否", ""],
        ["S18", "长引文", "楷体", "楷体", "四号", "否", "两端对齐", "字符", "2", "字符", "2", "磅", "0", "磅", "0", "固定值（磅）", "20", "否", ""],
        ["S19", "附录标题", "黑体", "黑体", "小四", "是", "居中", "-", "0", "-", "0", "磅", "24", "磅", "24", "多倍行距（倍）", "1", "是", ""],
        ["S20", "附录序号题名", "宋体", "宋体", "五号", "否", "左对齐", "字符", "2", "-", "0", "磅", "18", "磅", "18", "多倍行距（倍）", "1", "否", ""],
        ["S21", "参考文献条目", "宋体", "Times New Roman", "五号", "否", "两端对齐", "-", "0", "-", "0", "磅", "3", "磅", "0", "固定值（磅）", "16", "否", ""],
        ["S22", "致谢正文", "宋体", "Times New Roman", "五号", "否", "两端对齐", "字符", "2", "-", "0", "磅", "3", "磅", "0", "固定值（磅）", "16", "否", ""],
        ["S23", "脚注文本", "宋体", "宋体", "小五", "否", "两端对齐", "-", "0", "-", "0", "磅", "0", "磅", "0", "多倍行距（倍）", "1", "否", ""],
        ["S24", "图题", "宋体", "宋体", "五号", "否", "居中", "-", "0", "-", "0", "磅", "6", "磅", "12", "多倍行距（倍）", "1", "否", ""],
        ["S25", "表题", "宋体", "宋体", "五号", "否", "居中", "-", "0", "-", "0", "磅", "6", "磅", "12", "多倍行距（倍）", "1", "否", ""],
        ["S26", "表格文字", "宋体", "宋体", "五号", "否", "左对齐", "-", "0", "-", "0", "磅", "3", "磅", "3", "多倍行距（倍）", "1", "否", ""],
        ["S27", "页眉", "宋体", "宋体", "小四", "否", "居中", "-", "0", "-", "0", "磅", "0", "磅", "0", "多倍行距（倍）", "1", "否", ""],
        ["S28", "页脚页码", "Times New Roman", "Times New Roman", "五号", "否", "居中", "-", "0", "-", "0", "磅", "0", "磅", "0", "多倍行距（倍）", "1", "否", ""],
        # 副标题（用户 2026-09-16：“中文副标题…至少应该靠右对齐，和主标题间距
        # 也不能那么大”）：原来是居中，改成右对齐（官方封面上的副标题也是偏右
        # 的，正文前置区跟着对齐才一致）。段后 24→0，句间距靠“摘要标签”自己的
        # 段前 18 磅拉开。
        ["S29", "副标题", "黑体", "黑体", "小四", "否", "右对齐", "-", "0", "-", "0", "磅", "0", "磅", "0", "多倍行距（倍）", "1", "否", ""],
        ["S30", "英文关键词标签", "Times New Roman", "Times New Roman", "四号", "是", "左对齐", "-", "0", "-", "0", "磅", "18", "磅", "18", "多倍行距（倍）", "1", "否", ""],
        ["S31", "英文关键词内容", "Times New Roman", "Times New Roman", "小四", "否", "左对齐", "-", "0", "-", "0", "磅", "18", "磅", "18", "多倍行距（倍）", "1", "否", ""],
        # 英文副标题（用户 2026-09-15 定的对称部件）：字体、字号、对齐全部参照
        # 中文副标题 S29，只有西文字体不同——**不能直接复用 S29**，S29 的
        # font_ascii 是黑体，英文拿它排出来是黑体的英文字。
        ["S32", "英文副标题", "Times New Roman", "Times New Roman", "小四", "否", "右对齐", "-", "0", "-", "0", "磅", "0", "磅", "0", "多倍行距（倍）", "1", "否", ""],
    ]
    data_start, data_end = add_block(
        ws, row, "样式", headers, rows,
        center_cols={1, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18},
    )
    add_dropdown(ws, 5, data_start, data_end, FONT_SIZES)
    add_dropdown(ws, 6, data_start, data_end, YN)
    add_dropdown(ws, 7, data_start, data_end, ALIGN)
    add_dropdown(ws, 8, data_start, data_end, INDENT_UNITS)
    add_dropdown(ws, 10, data_start, data_end, INDENT_UNITS)
    add_dropdown(ws, 12, data_start, data_end, SPACING_UNITS)
    add_dropdown(ws, 14, data_start, data_end, SPACING_UNITS)
    add_dropdown(ws, 16, data_start, data_end, LINE_RULE)
    add_dropdown(ws, 18, data_start, data_end, YN)
    row += 1 + 1 + len(rows) + 1

    # --------------------------------------------------------------- 页眉页脚
    headers = ["节", "页眉文字", "启用页码", "页码编码", "前缀", "后缀", "备注"]
    rows = [
        ["前置部分(摘要起至目录)", "中 国 政 法 大 学 本 科 生 毕 业 论 文", "是", "I,II,III", "", "", ""],
        ["正文部分(引言起)", "中 国 政 法 大 学 本 科 生 毕 业 论 文", "是", "1,2,3", "", "", ""],
    ]
    data_start, data_end = add_block(ws, row, "页眉页脚", headers, rows, center_cols={3, 4, 5, 6})
    add_dropdown(ws, 3, data_start, data_end, YN)
    add_dropdown_range(ws, 4, data_start, data_end, page_number_format_range)
    row += 1 + 1 + len(rows) + 1

    # ---------------------------------------------------------------- 页眉线
    # 官方模板的页眉文字下面有一条横线——那是画在“页眉”样式上的下边框，不是
    # 逐页加的横线形状。这里是要不要的总开关；填“不要”会把样式上的下边框摘掉
    # （源文档自带的线也算）。封面永远没有这条线：封面节不放页眉，跟本开关无关。
    headers = ["项目", "值", "备注"]
    rows = [
        ["页眉文字下画横线", "要",
         "官方模板有这条线；填“不要”则页眉文字下面不画线（封面本来就没有）"],
    ]
    data_start, data_end = add_block(ws, row, "页眉线", headers, rows, center_cols={2, 3})
    add_dropdown(ws, 2, data_start, data_start + 1, ["要", "不要"])
    row += 1 + 1 + len(rows) + 1

    # -------------------------------------------------------------------- 引注
    headers = ["项目", "值", "备注"]
    rows = [
        ["脚注序号样式", "①②③", ""],
        ["脚注每页重新编号", "是", ""],
        ["脚注序号字体", "正文(非上标)", ""],
        ["脚注序号后空白", "半角空格", ""],
        ["参考文献编号样式", "[1][2][3]", ""],
    ]
    data_start, data_end = add_block(ws, row, "引注", headers, rows, center_cols={2, 3})
    add_dropdown(ws, 2, data_start, data_start + 1, ["①②③", "1)2)3)", "[1][2][3]"])  # 脚注序号样式
    add_dropdown(ws, 2, data_start + 1, data_start + 2, YN)  # 脚注每页重新编号
    add_dropdown(ws, 2, data_start + 2, data_start + 3, ["正文(非上标)", "上标"])  # 脚注序号字体
    add_dropdown(ws, 2, data_start + 3, data_start + 4, ["半角空格", "全角空格", "不空格"])  # 脚注序号后空白
    add_dropdown(ws, 2, data_start + 4, data_start + 5, ["[1][2][3]", "1.2.3.", "(1)(2)(3)"])  # 参考文献编号样式
    row += 1 + 1 + len(rows) + 1

    # --------------------------------------------------------------- 清理规则
    # 套版前会把源文档的直接格式（字体、字号、行距……）清掉，好让预设样式真正
    # 生效。但**加粗/斜体/下划线/上下标这类是"内容语义"**——作者自己敲的强调，
    # 清掉会改变论文的意思——所以默认保留。这一格就是控制保留哪些的开关；
    # 填"无"则全部清掉（输出完全按预设来，作者自己的强调也不留）。
    headers = ["项目", "值", "备注"]
    rows = [
        ["保留的字符格式", "加粗、斜体、下划线、上标下标",
         "源文档正文里的这些强调格式原样保留；填“无”则一律清掉"],
    ]
    data_start, data_end = add_block(ws, row, "清理规则", headers, rows, center_cols={2, 3})
    add_dropdown(
        ws, 2, data_start, data_start + 1,
        ["加粗、斜体、下划线、上标下标", "加粗、斜体、下划线", "加粗、斜体", "无"],
    )
    row += 1 + 1 + len(rows) + 1

    # ---------------------------------------------------------------- 封面
    # 封面信息本身从 GUI 填（不进 Excel，用户 2026-09-12 拍板）；这里只放
    # 封面**行为**的开关。后两行 2026-09-15 跟着 `excel_writer` 补上：
    # `cover_enabled`/`cover_form` 以前只在 JSON 预设里存得住，导出成 Excel
    # 这两项就没了（留空＝套用这个预设时不改界面上当前的选择，旧预设就是
    # 这种“没写”的状态）。
    headers = ["项目", "值", "备注"]
    rows = [
        ["超长值悬挂缩进", "要",
         "值太长换行时，续行对齐到提示词右侧（首行不变）；填“不要”则续行回到最左边"],
        ["是否插入封面", "",
         "留空＝套用这个预设时不改封面开关；也可填“要”/“不要”"],
        ["封面形态", "", "留空＝套用时保持界面上当前选的形态（可填" + "/".join(COVER_FORM_NAMES) + "）"],
    ]
    data_start, data_end = add_block(ws, row, "封面", headers, rows, center_cols={2, 3})
    add_dropdown(ws, 2, data_start, data_start + 1, ["要", "不要"])
    add_dropdown(ws, 2, data_start + 1, data_start + 2, ["要", "不要"], allow_blank=True)
    add_dropdown(ws, 2, data_start + 2, data_start + 3, list(COVER_FORM_NAMES), allow_blank=True)
    row += 1 + 1 + len(rows) + 1

    # ------------------------------------------------------------ 结构部件
    # "插入文字"：启用了但源文件没有时，自动补上的那一段该写成什么字面文字
    # （带不带【】、字间要不要空格）。题目/副标题/英文题目没有"标准答案"可填
    # ——题目本身就是内容，没法自动生成——这几行留空，工具遇到空值会插入占位
    # 提示，提示语的措辞是工具行为，不是排版规范，留在代码里，不进这张表。
    #
    # "是否启用"这一列＝**这份预设**的部件开关（2026-09-14 用户逐个点过：
    # 题目/副标题/摘要/关键词/英文题目/英文摘要/英文关键词/目录/参考文献/
    # 致谢 全开，三张清单和附录关）。英文题目和英文关键词 2026-09-14
    # 起是**独立部件**（以前英文关键词并"英文摘要"那一格里、用换行分两行写），
    # 各自有自己的"插入文字"格。
    #
    # 引言/结论**不在本表里**（用户 2026-09-16："把这两个从预设里删掉，当完整
    # 的标题"）：它们是正文里的一级标题，位置即身份，不需要开关，也没有"没
    # 识别到就补一段占位"这回事。
    #
    # "插图清单/附表清单/插图和附表清单"（规范 [91]-[93]）是三个**可选的**
    # 部件，规范原文是"图表较多时，可以……"，所以默认都是"否"。它们的"内容"
    # 不是一个词、而是工具插的 Word 域（跟目录域同机制，Word 自己收集题注、
    # 自己算页码），"插入文字"只决定标题那几个字怎么写；三行怎么组合决定形态：
    # 启用独立两行 = 两张分开的清单（[91]），启用"插图和附表清单" = 一张合并
    # 清单（[92] 的"插图在前、附表在后"）——细节见 apply_structure.LIST_PART_KINDS。
    headers = ["部件", "是否启用", "插入文字", "备注"]
    rows = [
        ["题目", "是", "", ""],
        ["副标题", "是", "", ""],
        ["摘要", "是", "【摘要】", ""],
        ["关键词", "是", "【关键词】", ""],
        ["英文题目", "是", "", "2026-09-14 起是独立部件（英文摘要页的第一句标题）"],
        ["英文副标题", "是", "", "2026-09-15 起是独立部件（英文题目下面那一行副标题）"],
        ["英文摘要", "是", "【ABSTRACT】", ""],
        ["英文关键词", "是", "【KEYWORDS】", "2026-09-14 起是独立部件（原属「英文摘要」）"],
        ["目录", "是", "【目录】", ""],
        ["插图清单", "否", "插图清单", "规范[91]：图片较多时才用；与「插图和附表清单」二选一"],
        ["附表清单", "否", "附表清单", "规范[91]：表格较多时才用"],
        ["插图和附表清单", "否", "插图和附表清单", "规范[92]：图表较少时的合并形态；不要与上面两行同时启用"],
        ["参考文献", "是", "参考文献", ""],
        ["附录", "否", "附  录", ""],
        ["致谢", "是", "致  谢", ""],
    ]
    data_start, data_end = add_block(ws, row, "结构部件", headers, rows, center_cols={2})
    add_dropdown(ws, 2, data_start, data_end, YN)
    row += 1 + 1 + len(rows) + 1

    for c, w in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w


def main():
    OUT_DIR.mkdir(exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "格式设置"

    # "页码编码"下拉的选项自己带逗号（"1,2,3"），没法直接内联进 Excel 的下拉公式
    # （逗号是公式里的分隔符，会被拆散），所以另放一张隐藏表，下拉引用这张表的区域。
    options_ws = wb.create_sheet("选项")
    for i, v in enumerate(PAGE_NUMBER_FORMATS, start=1):
        options_ws.cell(row=i, column=1, value=v)
    options_ws.sheet_state = "hidden"
    page_number_format_range = f"选项!$A$1:$A${len(PAGE_NUMBER_FORMATS)}"

    build(ws, page_number_format_range)
    out_path = OUT_DIR / "中国政法大学-本科毕业论文.xlsx"
    wb.save(out_path)
    print(f"已生成: {out_path}")


if __name__ == "__main__":
    main()
