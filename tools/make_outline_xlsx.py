# -*- coding: utf-8 -*-
"""生成《工具操作逻辑》写作表：四列（章 / 节 / 正文 / 备注）。

约定的口径（用户 2026-09-16 定）：
  * 只列标题层级，不写"这一节该写什么"——正文由用户自己在表里写；
  * 第 2 章已定稿的正文**按节拆回表里**，直接放在对应行的「正文」列；
  * 列只有四列：章、节、正文、备注（"此外每个字都多余"）；
  * 章用合并单元格跨行表示归属。

第 2 章的正文不在这里重复维护：脚本从 `essay_tool/gui/main_window.py` 的
`OPERATION_LOGIC_TEXT` 里原样抠出来，按 "2.x" 标题切分后填进各行的正文列。
**要改第 2 章的正文，改程序里那一处，再重跑本脚本**；在表里改了正文，
下次重跑会被覆盖（所以本表适合往下写第 3 章起的内容）。

⚠️ **2026-09-16 起方向已经反过来**：现在以 `reference/工具操作逻辑写作表.xlsx`
那份表为准，窗口里显示的正文改由 `tools/make_operation_logic_text.py` 从表生成
（见 `essay_tool/gui/operation_logic_text.py`）。本脚本只用于当初建表，写的是
`docs/` 下的旧镜像，**不要再拿它去覆盖 reference/ 那份表**。

用法：.venv\\Scripts\\python.exe tools\\make_outline_xlsx.py
"""
from __future__ import annotations

import pathlib
import re
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = pathlib.Path(__file__).resolve().parents[1]
OUT = ROOT / "docs" / "工具操作逻辑写作表.xlsx"
# 正文的出处：2026-09-16 起从 main_window 挪到了这个生成模块
TEXT_MODULE = ROOT / "essay_tool" / "gui" / "operation_logic_text.py"

from openpyxl import Workbook                                        # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402

# ---------------------------------------------------------------- 标题层级
# (章, [(节, [小节, ...]), ...])
TITLES: list[tuple[str, list[tuple[str, list[str]]]]] = [
    ("1. 软件概况与界面", [
        ("1.1 用途与适用范围", []),
        ("1.2 运行环境与启动方式", []),
        ("1.3 界面分区", []),
        ("1.4 六个页签一览", []),
    ]),
    ("2. 导入和自动备份", [
        ("2.1 导入", []),
        ("2.2 自动备份的触发时机与目标路径确定方式", []),
        ("2.3 同名文件的处理规则", []),
        ("2.4 备份路径的自定义设置", []),
        ("2.5 自动清理策略及保留方式", []),
        ("2.6 备份清单", []),
        ("2.7 备份机制与核心处理流程的独立性", []),
    ]),
    ("3. 格式设置（预设）", [
        ("3.1 预设的作用与存放位置", []),
        ("3.2 应用预设：何时生效", []),
        ("3.3 页面设置", []),
        ("3.4 样式表（S01–S32 各角色）", []),
        ("3.5 页眉、页脚与页码", []),
        ("3.6 清理规则：保留的字符格式", []),
        ("3.7 结构部件与封面的启用开关", []),
        ("3.8 预设的导入与导出", []),
        ("3.9 不套用任何预设时的默认值", []),
    ]),
    ("4. 识别", [
        ("4.1 识别时机", []),
        ("4.2 识别产出：每个段落的身份", []),
        ("4.3 标题的判定", [
            "4.3.1 一级标题",
            "4.3.2 二级标题",
            "4.3.3 三、四级标题与长度口径",
            "4.3.4 不算标题的情形与「疑似标题」提醒",
        ]),
        ("4.4 题目与副标题", []),
        ("4.5 英文题目与英文副标题", []),
        ("4.6 前置标签", []),
        ("4.7 目录与三张清单", []),
        ("4.8 锚点部件与内容范围", []),
        ("4.9 题注：图题与表题", [
            "4.9.1 判定路径",
            "4.9.2 续行、英文对照行、分图",
            "4.9.3 不认识的标签",
        ]),
        ("4.10 封面范围：为什么不做自动识别", []),
        ("4.11 识别不准时的处置", []),
        ("4.12 告警", []),
    ]),
    ("5. 清理与规范化", [
        ("5.1 为什么要先清一遍", []),
        ("5.2 清除范围与保留范围", []),
        ("5.3 标题的例外", []),
        ("5.4 Word 自动编号的固化", []),
        ("5.5 段落层面的规范化", []),
        ("5.6 图片段落的规范化", []),
        ("5.7 内容控件（Word 目录被裹在其中时）", []),
        ("5.8 封面区间内不做任何加工", []),
    ]),
    ("6. 样式应用与套版", [
        ("6.1 页面设置与分节", []),
        ("6.2 真样式与直接格式", []),
        ("6.3 各级标题与正文", []),
        ("6.4 前置、后置部件", []),
        ("6.5 题注、表格与表格文字", []),
        ("6.6 长引文", []),
        ("6.7 脚注", []),
        ("6.8 缺部件的补齐", []),
        ("6.9 目录与三张清单", [
            "6.9.1 已有真域：一律不动",
            "6.9.2 没有真域时插入",
            "6.9.3 「重新为文档生成」如何生效",
        ]),
        ("6.10 加工顺序总表", []),
        ("6.11 重复导出的稳定性", []),
    ]),
    ("7. 结构页（人工核对与调整）", [
        ("7.1 表格怎么读", []),
        ("7.2 拖动改身份", []),
        ("7.3 勾选标记", []),
        ("7.4 封面范围与删除封面", []),
        ("7.5 清单部件的启用与重新生成", []),
        ("7.6 参考文献的自动重新编号", []),
        ("7.7 调整结果如何影响导出", []),
        ("7.8 显示规则", []),
        ("7.9 重新识别、重置与确认", []),
    ]),
    ("8. 封面", [
        ("8.1 两种封面形态", []),
        ("8.2 字段与书写规则", []),
        ("8.3 超长值的悬挂缩进", []),
        ("8.4 插入位置与分节", []),
        ("8.5 原有封面的保留与删除", []),
    ]),
    ("9. 导出", [
        ("9.1 导出流程", []),
        ("9.2 输出位置、默认命名与覆盖确认", []),
        ("9.3 导出空白模板", []),
        ("9.4 导出之后要做什么", []),
        ("9.5 导出失败与提示", []),
    ]),
    ("10. 日志、告警与异常处置", [
        ("10.1 日志区", []),
        ("10.2 告警的两种查看途径", []),
        ("10.3 常见异常与自处办法", []),
    ]),
    ("11. 关于页", [
        ("11.1 使用协议与第三方许可", []),
        ("11.2 查阅《规范》原文", []),
        ("11.3 另存官方版、联系开发者", []),
    ]),
]

HEADERS = ["章", "节", "正文", "备注"]
WIDTHS = {"A": 20, "B": 34, "C": 96, "D": 28}

F_HEAD = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
F_CH = Font(name="微软雅黑", size=10, bold=True, color="1A3E6E")
F_SEC = Font(name="微软雅黑", size=10, bold=True)
F_SUB = Font(name="微软雅黑", size=10)
F_BODY = Font(name="宋体", size=10)
FILL_HEAD = PatternFill("solid", fgColor="4472C4")
FILL_CH = PatternFill("solid", fgColor="D9E2F3")
FILL_SEC = PatternFill("solid", fgColor="EDF3FA")

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP_TOP = Alignment(vertical="top", wrap_text=True)
CENTER = Alignment(vertical="center", horizontal="center", wrap_text=True)
CENTER_LEFT = Alignment(vertical="center", horizontal="left", wrap_text=True)


def read_chapter_bodies() -> dict[str, str]:
    """从 main_window.py 抠出 OPERATION_LOGIC_TEXT，再按 "2.x" 标题切成
    {节标题: 该节正文}（正文按原样分段，段与段之间用一个换行）。"""
    source = TEXT_MODULE.read_text(encoding="utf-8")
    m = re.search(r'OPERATION_LOGIC_TEXT = r"""(.*?)"""', source, re.S)
    if m is None:
        raise SystemExit(f"没在 {TEXT_MODULE.name} 里找到 OPERATION_LOGIC_TEXT")

    bodies: dict[str, str] = {}
    current: str | None = None
    for raw in m.group(1).strip().splitlines():
        line = raw.strip()
        head = re.match(r"^(\d+(?:\.\d+)*)\s+(\S.*)$", line)
        if head:
            current = f"{head.group(1)} {head.group(2)}"
            bodies[current] = ""
            continue
        if not line or current is None:
            continue
        bodies[current] = (bodies[current] + "\n" + line) if bodies[current] else line
    return bodies


def main() -> None:
    bodies = read_chapter_bodies()

    wb = Workbook()
    ws = wb.active
    ws.title = "工具操作逻辑"

    for col, head in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col, value=head)
        cell.font = F_HEAD
        cell.fill = FILL_HEAD
        cell.alignment = CENTER
        cell.border = BOX
    ws.row_dimensions[1].height = 22

    row = 2
    filled: list[str] = []
    for chapter, sections in TITLES:
        span = sum(max(1, len(subs)) for _title, subs in sections)
        ch_cell = ws.cell(row=row, column=1, value=chapter)
        ch_cell.font = F_CH
        ch_cell.fill = FILL_CH
        ch_cell.alignment = CENTER
        if span > 1:
            ws.merge_cells(start_row=row, start_column=1,
                           end_row=row + span - 1, end_column=1)

        for title, subs in sections:
            if not subs:
                title_cell = ws.cell(row=row, column=2, value=title)
                title_cell.font = F_SEC
                title_cell.fill = FILL_SEC
                title_cell.alignment = CENTER_LEFT
                body = bodies.get(title)
                if body:
                    body_cell = ws.cell(row=row, column=3, value=body)
                    body_cell.font = F_BODY
                    body_cell.alignment = WRAP_TOP
                    filled.append(title)
                row += 1
                continue

            # 有下级标题的节：节标题单独一行，小节各占一行（靠编号体现层级）
            title_cell = ws.cell(row=row, column=2, value=title)
            title_cell.font = F_SEC
            title_cell.fill = FILL_SEC
            title_cell.alignment = CENTER_LEFT
            row += 1
            for sub in subs:
                sub_cell = ws.cell(row=row, column=2, value=sub)
                sub_cell.font = F_SUB
                sub_cell.alignment = CENTER_LEFT
                body = bodies.get(sub)
                if body:
                    body_cell = ws.cell(row=row, column=3, value=body)
                    body_cell.font = F_BODY
                    body_cell.alignment = WRAP_TOP
                    filled.append(sub)
                row += 1

    for r in range(2, row):
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = BOX
    for col, width in WIDTHS.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "C2"

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"已生成：{OUT}")
    print(f"  章 {len(TITLES)} 个、节 "
          f"{sum(len(s) for _t, s in TITLES)} 个、表行 {row - 2} 行")
    print(f"  已从程序里填入正文的节：{filled}")
    missing = [t for _ch, secs in TITLES for t, _s in secs
               if t.startswith("2.") and t not in bodies]
    if missing:
        print(f"  ★ 第 2 章里有节没匹配到正文：{missing}")


if __name__ == "__main__":
    main()
