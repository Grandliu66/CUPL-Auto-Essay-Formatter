# -*- coding: utf-8 -*-
r"""把《工具操作逻辑》窗口正文导出成一份**可编辑的 Word 文档**。

正文源头＝`essay_tool/gui/operation_logic_text.py` 的 `OPERATION_LOGIC_TEXT`
（2026-09-16 起手工维护，就是源头）。导出规则**与查阅窗口同一套判据**
（直接 import `logic_viewer` 的解析函数，不抄第二份）：

  * 标题：`2.` / `2.1` / `2.1.1`（数字与文字之间有空格）——按点号个数分级，
    落成 Word 的「标题 1/2/3…」样式（导航窗格能用）；
  * 表格：Markdown 管道表（第二行必须是 `|---|` 分隔行）→ **真 Word 表格**
    （网格线、表头加粗、按分隔行的冒号保留右对齐/居中）；
  * 其余非空行 → 正文段落；格子里写的字面量 `\n` 落成真换行。

用法：
    .venv\Scripts\python.exe tools\export_operation_logic_docx.py
    .venv\Scripts\python.exe tools\export_operation_logic_docx.py 输出.docx
    .venv\Scripts\python.exe tools\export_operation_logic_docx.py --force
                                                # 目标已存在（比如你改过的稿）时覆盖

⚠️ 目标文件已存在时**默认拒绝执行**——防止把用户改过的稿覆盖掉。
导出后可跑 `--check`（在同一脚本里默认做）：重新读一遍产物、逐块对回原文。
"""
from __future__ import annotations

import datetime
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from docx import Document                                       # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH                   # noqa: E402
from docx.oxml.ns import qn                                     # noqa: E402
from docx.shared import Cm, Pt, RGBColor                        # noqa: E402
from essay_tool.gui.logic_viewer import (_HEADING_RE,           # noqa: E402
                                         _column_aligns,
                                         _split_table_row,
                                         parse_table)
from essay_tool.gui.operation_logic_text import OPERATION_LOGIC_TEXT  # noqa: E402


def blocks_of(text: str) -> list[tuple]:
    """把正文切成块：("h", 级别, 原文) / ("p", 原文) / ("t", 表头, 数据行)。"""
    lines = text.splitlines()
    blocks: list[tuple] = []
    i = 0
    while i < len(lines):
        stripped = lines[i].strip()
        if not stripped:
            i += 1
            continue
        table = parse_table(lines, i, len(lines))
        if table is not None:
            header, rows, stop = table
            sep = _split_table_row(lines[i + 1])
            blocks.append(("t", header, rows, _column_aligns(sep)))
            i = stop
            continue
        match = _HEADING_RE.match(stripped)
        if match:
            blocks.append(("h", match.group(1).count(".") + 1, stripped))
        else:
            blocks.append(("p", stripped))
        i += 1
    return blocks


def _style_fonts(style, ascii_font: str, east_font: str) -> None:
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:ascii"), ascii_font)
    rfonts.set(qn("w:hAnsi"), ascii_font)
    rfonts.set(qn("w:eastAsia"), east_font)


def _run_fonts(run, ascii_font: str, east_font: str) -> None:
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:ascii"), ascii_font)
    rfonts.set(qn("w:hAnsi"), ascii_font)
    rfonts.set(qn("w:eastAsia"), east_font)


def build(blocks: list[tuple], target: pathlib.Path) -> dict:
    document = Document()
    section = document.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)

    normal = document.styles["Normal"]
    normal.font.size = Pt(10.5)
    _style_fonts(normal, "Times New Roman", "宋体")
    normal.paragraph_format.space_after = Pt(4)

    heading_sizes = {1: 15, 2: 13, 3: 11.5}
    for level in (1, 2, 3, 4):
        style = document.styles[f"Heading {level}"]
        style.font.size = Pt(heading_sizes.get(level, 10.5))
        style.font.bold = True
        style.font.color.rgb = RGBColor(0, 0, 0)
        _style_fonts(style, "Arial", "黑体")
        style.paragraph_format.space_before = Pt(10 if level == 1 else 8)
        style.paragraph_format.space_after = Pt(4)

    title = document.add_paragraph("工具操作逻辑")
    title.style = document.styles["Title"]
    title.style.font.size = Pt(18)
    _style_fonts(title.style, "Arial", "黑体")

    stats = {"标题": 0, "段落": 0, "表格": 0, "表格行": 0, "表格列": 0}
    for block in blocks:
        if block[0] == "h":
            document.add_paragraph(block[2], style=f"Heading {min(block[1], 4)}")
            stats["标题"] += 1
        elif block[0] == "p":
            document.add_paragraph(block[1])
            stats["段落"] += 1
        else:
            _, header, rows, aligns = block
            table = document.add_table(rows=1 + len(rows), cols=len(header))
            table.style = document.styles["Table Grid"]
            table.autofit = True
            stats["表格"] += 1
            stats["表格行"] += len(rows)
            stats["表格列"] = max(stats["表格列"], len(header))
            align_map = {"w": WD_ALIGN_PARAGRAPH.LEFT,
                         "center": WD_ALIGN_PARAGRAPH.CENTER,
                         "e": WD_ALIGN_PARAGRAPH.RIGHT}
            for col, text in enumerate(header):
                _fill_cell(table.cell(0, col), text, align_map.get(aligns[col]) if
                           col < len(aligns) else None, bold=True)
            for row_index, row in enumerate(rows, start=1):
                for col, text in enumerate(row):
                    _fill_cell(table.cell(row_index, col), text,
                               align_map.get(aligns[col]) if col < len(aligns) else None)
    document.save(str(target))
    return stats


def _fill_cell(cell, text: str, align, *, bold: bool = False) -> None:
    """一格的文字——字面量 `\\n` 已经是真换行（`_split_table_row` 换的），
    每行落成格子里一个段落（可读、回读也稳）。"""
    lines = str(text).split("\n") or [""]
    cell.paragraphs[0].text = ""
    for index, line in enumerate(lines):
        paragraph = cell.paragraphs[0] if index == 0 else cell.add_paragraph()
        paragraph.paragraph_format.space_after = Pt(0)
        if align is not None:
            paragraph.alignment = align
        run = paragraph.add_run(line)
        run.font.size = Pt(9)
        run.font.bold = bold
        _run_fonts(run, "Times New Roman", "宋体")


def check(blocks: list[tuple], target: pathlib.Path) -> list[str]:
    """把产物读回来逐块对原文——对不上就列出差异（导出保真度的机械验证）。"""
    document = Document(str(target))
    again: list[tuple] = []
    body = document.element.body
    for child in body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            from docx.text.paragraph import Paragraph               # noqa: PLC0415
            paragraph = Paragraph(child, document)
            name = str(paragraph.style.name)
            if name == "Title" or not paragraph.text.strip():
                continue
            if name.startswith("Heading "):
                again.append(("h", int(name.split(" ")[1]), paragraph.text.strip()))
            else:
                again.append(("p", paragraph.text.strip()))
        elif tag == "tbl":
            from docx.table import Table                            # noqa: PLC0415
            table = Table(child, document)
            rows = [[cell.text.strip() for cell in row.cells]
                    for row in table.rows]
            again.append(("t", rows[0], rows[1:]))
    problems: list[str] = []
    if len(again) != len(blocks):
        problems.append(f"块数不同：原文 {len(blocks)}，产物 {len(again)}")
    for source, produced in zip(blocks, again):
        if source[0] != produced[0]:
            problems.append(f"块类型不同：{source!r:.80} vs {produced!r:.80}")
            continue
        if source[0] == "t":
            head_ok = [c.replace("\\n", "\n").strip() for c in source[1]] == produced[1]
            rows_ok = ([[c.replace('\\n', '\n').strip() for c in row]
                        for row in source[2]] == produced[2])
            if not (head_ok and rows_ok):
                problems.append(f"表格不一致：表头 {source[1]} vs {produced[1]}")
        elif source[0] == "h":
            if (source[1], source[2]) != (produced[1], produced[2]):
                problems.append(f"标题不一致：{source!r} vs {produced!r}")
        else:
            if source[1] != produced[1]:
                problems.append(f"段落不一致：{source[1][:40]!r} vs {produced[1][:40]!r}")
    return problems


def main(argv: list[str]) -> int:
    force = "--force" in argv
    args = [a for a in argv[1:] if not a.startswith("--")]
    target = (pathlib.Path(args[0]) if args else
              ROOT / "reference" /
              f"工具操作逻辑-编辑稿-{datetime.date.today():%Y%m%d}.docx")
    if target.exists() and not force:
        print(f"目标已存在（可能你改过）：{target}\n"
              "要覆盖就加 --force，或直接传入另一个输出路径。")
        return 2

    blocks = blocks_of(OPERATION_LOGIC_TEXT)
    stats = build(blocks, target)
    print(f"已导出：{target}")
    print(f"  标题 {stats['标题']} 个、正文段 {stats['段落']} 段、"
          f"表格 {stats['表格']} 张（共 {stats['表格行']} 数据行，"
          f"最宽 {stats['表格列']} 列）")

    problems = check(blocks, target)
    if problems:
        print(f"\n✘ 回读校验失败（{len(problems)} 处）：")
        for item in problems[:20]:
            print("  " + item)
        return 1
    print("  回读校验通过：产物逐块与正文一致。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
