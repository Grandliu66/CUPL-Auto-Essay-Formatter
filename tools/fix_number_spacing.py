"""把《工具操作逻辑写作表.xlsx》第一个 sheet 里"编号后缺空格"的写法补上空格。

用户 2026-09-16：3.4识别范围 / 3.5自动排版 / 3.6识别和自动排版逻辑 这类写法与
其余各节不一致，要统一成「3.4 识别范围」。正文里指向节号的交叉引用同理
（参见3.4识别范围 → 参见 3.4 识别范围）。

只动这两处，不改任何标题文字、不动其他单元格。幂等：已经是带空格的就跳过。
跑法：.venv\\Scripts\\python.exe tools\\fix_number_spacing.py
"""

from __future__ import annotations

import pathlib
import re
import shutil

import openpyxl
from openpyxl.styles import PatternFill

ROOT = pathlib.Path(__file__).resolve().parents[1]
TARGET = ROOT / "reference" / "工具操作逻辑写作表.xlsx"
BACKUP_DIR = ROOT / "work"
SHEET = "工具操作逻辑"
YELLOW = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")

# 「3.4识别范围」：节号后面直接跟文字
TIGHT_SECTION_RE = re.compile(r"^(\d+[.．、]\d+)(?=\S)")
# 「参见3.4识别范围」：引用语后面直接跟节号
TIGHT_REF_RE = re.compile(r"参见(\d+[.．、]\d+)(\S*)")


def next_backup() -> pathlib.Path:
    for i in range(1, 100):
        candidate = BACKUP_DIR / f"_logic_sheet_backup_before_spacing-{i}.xlsx"
        if not candidate.exists():
            return candidate
    raise SystemExit("备份文件名用尽")


def main() -> None:
    book = openpyxl.load_workbook(TARGET)
    ws = book[SHEET]
    changes: list[tuple[str, str, str]] = []

    for row in range(2, ws.max_row + 1):
        # 节标题（B 列）
        b_cell = ws.cell(row=row, column=2)
        b_val = "" if b_cell.value is None else str(b_cell.value).strip()
        if b_val:
            new_b = TIGHT_SECTION_RE.sub(r"\1 ", b_val, count=1)
            if new_b != b_val:
                changes.append((f"B{row}", b_val, new_b))
                b_cell.value = new_b
                b_cell.fill = YELLOW

        # 正文里的交叉引用（C 列）：只在"参见+数字"这个组合上补空格
        c_cell = ws.cell(row=row, column=3)
        c_val = "" if c_cell.value is None else str(c_cell.value)
        if c_val:
            new_c = TIGHT_REF_RE.sub(lambda m: f"参见 {m.group(1)} {m.group(2)}".rstrip(),
                                     c_val, count=1)
            if new_c != c_val:
                changes.append((f"C{row}", c_val, new_c))
                c_cell.value = new_c
                c_cell.fill = YELLOW

    if not changes:
        print("没有缺空格的编号，未改动。")
        return

    backup = next_backup()
    shutil.copy2(TARGET, backup)
    print(f"改前版本已备份：{backup.relative_to(ROOT)}")
    print(f"\n共改 {len(changes)} 个单元格：")
    for coord, old, new in changes:
        print(f"  {SHEET}!{coord:<4} {old.replace(chr(10), '⏎')}  →  {new.replace(chr(10), '⏎')}")

    book.save(TARGET)
    print(f"\n已写回 {TARGET.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
