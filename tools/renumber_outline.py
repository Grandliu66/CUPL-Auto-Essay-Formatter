"""按当前行序重新编号《工具操作逻辑写作表.xlsx》第一个 sheet 的章号与节号。

用户删掉了若干章、节（2026-09-16 晚），号出现空档。本脚本只做"把号重新连起来"这一件事：

- A 列章标题按出现顺序依次编 1、2、3……
- B 列节标题在新章内从 1 顺序编，章号跟随该章
- 原编号与标题之间的空格（有的节没有空格）照抄，标题文字一个字不改
- C 列里指向章节号的两处交叉引用（参见4.4识别范围、参见4.7自动告警）跟着改
- 改过的单元格一律填黄底；写盘前备份到 work/

幂等：已经是连续编号时什么都不改。
跑法：.venv\\Scripts\\python.exe tools\\renumber_outline.py
"""

from __future__ import annotations

import pathlib
import re
import shutil

import openpyxl
from openpyxl.styles import PatternFill

ROOT = pathlib.Path(__file__).resolve().parents[1]
TARGET = ROOT / "reference" / "工具操作逻辑写作表.xlsx"
MIRROR = ROOT / "docs" / "工具操作逻辑写作表.xlsx"
BACKUP_DIR = ROOT / "work"
SHEET = "工具操作逻辑"
YELLOW = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")

# 「2. 导入和自动备份」→「章号 + 分隔符 + 空格 + 标题」
CHAPTER_RE = re.compile(r"^(\d+)([.．、])(\s*)(\S.*)$")
# 「2.1 导入」「4.4识别范围」→「章号 + 分隔符 + 节号 + 空格 + 标题」
SECTION_RE = re.compile(r"^(\d+)([.．、])(\d+)(\s*)(\S.*)$")
# 正文里引用的节号，如「参见4.4识别范围」
REF_RE = re.compile(r"(?<![\d.])(\d+[.．、]\d+)(?![\d.])")


def next_backup() -> pathlib.Path:
    for i in range(1, 100):
        candidate = BACKUP_DIR / f"_logic_sheet_backup_before_renumber-{i}.xlsx"
        if not candidate.exists():
            return candidate
    raise SystemExit("备份文件名用尽")


def text_of(cell) -> str:
    return "" if cell.value is None else str(cell.value)


def main() -> None:
    book = openpyxl.load_workbook(TARGET)
    ws = book[SHEET]

    chapter = 0
    section = 0
    label_map: dict[str, str] = {}
    changes: list[tuple[str, str, str]] = []
    troubles: list[str] = []

    for row in range(2, ws.max_row + 1):
        a_cell, b_cell = ws.cell(row=row, column=1), ws.cell(row=row, column=2)
        a_val, b_val = text_of(a_cell).strip(), text_of(b_cell).strip()

        if a_val:
            match = CHAPTER_RE.match(a_val)
            if match is None:
                troubles.append(f"A{row} 认不出章号：{a_val!r}")
                continue
            chapter += 1
            section = 0
            new_a = f"{chapter}{match.group(2)}{match.group(3)}{match.group(4)}"
            if new_a != a_val:
                changes.append((f"A{row}", a_val, new_a))
                a_cell.value = new_a
                a_cell.fill = YELLOW

        if b_val:
            match = SECTION_RE.match(b_val)
            if match is None:
                troubles.append(f"B{row} 认不出节号：{b_val!r}")
                continue
            section += 1
            old_label = f"{match.group(1)}{match.group(2)}{match.group(3)}"
            new_label = f"{chapter}{match.group(2)}{section}"
            label_map[old_label] = new_label
            new_b = f"{new_label}{match.group(4)}{match.group(5)}"
            if new_b != b_val:
                changes.append((f"B{row}", b_val, new_b))
                b_cell.value = new_b
                b_cell.fill = YELLOW

    # 正文里的交叉引用跟着改（一次扫描、不回头，避免 4.4→3.4 又被 3.4→2.4 二次替换）
    for row in range(2, ws.max_row + 1):
        c_cell = ws.cell(row=row, column=3)
        text = text_of(c_cell)
        if not text:
            continue
        new_text = REF_RE.sub(lambda m: label_map.get(m.group(1), m.group(1)), text)
        if new_text != text:
            changes.append((f"C{row}", text, new_text))
            c_cell.value = new_text
            c_cell.fill = YELLOW

    if troubles:
        print("⚠ 有认不出的编号，未动那些格子：")
        for item in troubles:
            print(f"   {item}")

    if not changes:
        print("编号已经是连续的，没有改动。")
        return

    backup = next_backup()
    shutil.copy2(TARGET, backup)
    print(f"改前版本已备份：{backup.relative_to(ROOT)}")

    print(f"\n共改 {len(changes)} 个单元格：")
    for coord, old, new in changes:
        shown_old = old.replace("\n", "⏎")
        shown_new = new.replace("\n", "⏎")
        print(f"  {SHEET}!{coord:<5} {shown_old}  →  {shown_new}")

    book.save(TARGET)
    print(f"\n已写回 {TARGET.relative_to(ROOT)}")

    if MIRROR.exists():
        print(
            "⚠ docs/ 那份镜像与 reference/ 不一致（被单独改过？），这次只动了 reference/"
            if MIRROR.read_bytes() != TARGET.read_bytes()
            else "docs/ 那份镜像与 reference/ 一致。"
        )


if __name__ == "__main__":
    main()
