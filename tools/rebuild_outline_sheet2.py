# -*- coding: utf-8 -*-
"""按用户的左侧合并单元格格式重排 Sheet1，并自动行高（交给 Excel/WPS 自适应）。

规则（用户 2026-09-16 说明）：
  * 一个识别对象由多个子对象组成时，B 列写对象名并纵向合并，C 列逐个列出子对象
    （“摘要”先写“摘要”，再拆成“摘要标签”和“摘要正文”）；
  * 只有一个子对象时，B:C 合并成一格；
  * 行高设为自动（不加锁），由表格程序按内容自适应；
  * 用户已写好的文字一字不改，只在需要拆分的地方按同一句式补出子行。

产物：docs\\工具操作逻辑写作表.xlsx（同时另存一份到 reference\\）
用法：.venv\\Scripts\\python.exe tools\\rebuild_outline_sheet2.py
"""
from __future__ import annotations

import copy
import pathlib
import shutil
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = pathlib.Path(__file__).resolve().parents[1]
BOOK = ROOT / "docs" / "工具操作逻辑写作表.xlsx"
MIRROR = ROOT / "reference" / "工具操作逻辑写作表.xlsx"
BACKUP = ROOT / "work" / "_写作表备份-重排前.xlsx"

import openpyxl                                                     # noqa: E402

SHEET = "Sheet1"

# 需要拆成多个子对象的（(对象名, [子对象名, …])）；其余对象保持单行（B:C 合并）
SPLIT: dict[str, list[str]] = {
    "英文摘要": ["“Abstract”标签", "英文摘要正文"],
    "英文关键词": ["“Keywords”标签", "英文关键词正文"],
    "引言": ["“引言”标签", "引言正文"],
    "结论": ["“结论”标签", "结论正文"],
    "附录": ["“附录”标签", "附录序号题名", "附录正文"],
    "致谢": ["“致谢”标签", "致谢正文"],
}

# 拆分后各子行的内容（列未写出的，表示该格留空）
SUB_TEXT: dict[tuple[str, str], dict[str, str]] = {
    ("英文摘要", "“Abstract”标签"): {
        "D": "满足\n（包括且只包括连续的“Abstract”文本）\n的段落，将会被识别为英文摘要标签。\n"
             "如果在下一个识别到的部件前不存在满足条件的段落，则判定不存在英文摘要。",
        "E": "忽略大小写；\n允许“Abstract”外存在任意的“【”“】”“[”“]”，不要求成对。",
        "F": "如果识别到本部件，则将直接在段落上按照预设格式调整格式，并强制将内容改为“【ABSTRACT】”；\n"
             "未识别到英文摘要标签且该部件启用，则自动插入占位文本。",
    },
    ("英文摘要", "英文摘要正文"): {
        "D": "满足\n（紧贴英文摘要标签之后）且（位于下一个部件标签之前）\n的段落，将会被判定为英文摘要正文。",
        "F": "如果识别到本部件，则将直接在段落上按照预设格式调整格式；\n"
             "未识别到英文摘要正文且该部件启用，则自动插入占位文本。",
    },
    ("英文关键词", "“Keywords”标签"): {
        "D": "满足\n（包括且只包括连续的“Keywords”文本）\n的段落，将会被识别为英文关键词标签。\n"
             "如果在下一个识别到的部件前不存在满足条件的段落，则判定不存在英文关键词。",
        "E": "忽略大小写；\n允许“Keywords”外存在任意的“【”“】”“[”“]”，不要求成对。",
        "F": "如果识别到本部件，则将直接在段落上按照预设格式调整格式，并强制将内容改为“【KEYWORDS】”；\n"
             "未识别到英文关键词标签且该部件启用，则自动插入占位文本。",
    },
    ("英文关键词", "英文关键词正文"): {
        "D": "满足\n（紧贴英文关键词标签之后）且（位于下一个部件标签之前）\n的段落，将会被判定为英文关键词正文。",
        "F": "如果识别到本部件，则将直接在段落上按照预设格式调整格式；\n"
             "未识别到英文关键词正文且该部件启用，则自动插入占位文本。",
    },
    ("引言", "“引言”标签"): {
        "D": "满足\n（包括且只包括连续的“引言”文本）\n的段落，将会被识别为引言标签。",
        "E": "允许字间与词尾存在半角或全角空格。",
        "F": "如果识别到本部件，则套用预设样式；\n"
             "未识别到引言标签且该部件启用，则自动插入占位文本。",
    },
    ("引言", "引言正文"): {
        "D": "满足\n（紧贴引言标签之后）且（位于下一个部件标签之前）\n的段落，将会被判定为引言正文。",
        "F": "如果识别到本部件，则将按照正文格式处理。",
    },
    ("结论", "“结论”标签"): {
        "D": "满足\n（包括且只包括连续的“结论”文本）\n的段落，将会被识别为结论标签。",
        "E": "允许字间与词尾存在半角或全角空格。",
        "F": "如果识别到本部件，则套用预设样式；\n"
             "未识别到结论标签且该部件启用，则自动插入占位文本。",
    },
    ("结论", "结论正文"): {
        "D": "满足\n（紧贴结论标签之后）且（位于下一个部件标签之前）\n的段落，将会被判定为结论正文。",
        "F": "如果识别到本部件，则将按照正文格式处理。",
    },
    ("附录", "“附录”标签"): {
        "D": "满足\n（包括且只包括连续的“附录”文本）\n的段落，将会被识别为附录标签。",
        "F": "如果识别到本部件，则套用预设样式；\n"
             "未识别到附录标签且该部件启用，则自动插入占位文本。",
    },
    ("附录", "附录序号题名"): {
        "D": "满足\n（紧贴附录标签之后）且（以“附”字、“录”字和一个拉丁字母开头）且（不以标点结尾）"
             "且（长度不超过40个字符）\n的段落，将会被识别为附录序号题名。",
        "E": "允许“附”字与“录”字之间存在空格；\n拉丁字母不区分大小写。",
        "F": "如果识别到本部件，则不视为标题，仅在段落上按照预设格式调整格式。",
    },
    ("附录", "附录正文"): {
        "D": "满足\n（紧贴附录序号题名之后）且（位于下一个部件标签之前）\n的段落，将会被判定为附录正文。",
        "F": "如果识别到本部件，则其中的第一段按照附录序号题名处理，其余段落按照正文格式处理。",
    },
    ("致谢", "“致谢”标签"): {
        "D": "满足\n（包括且只包括连续的“致谢”文本）\n的段落，将会被识别为致谢标签。",
        "E": "允许字间与词尾存在半角或全角空格。",
        "F": "如果识别到本部件，则套用预设样式；\n"
             "未识别到致谢标签且该部件启用，则自动插入占位文本。",
    },
    ("致谢", "致谢正文"): {
        "D": "满足\n（紧贴致谢标签之后）且（位于下一个部件标签之前）\n的段落，将会被判定为致谢正文。",
        "F": "如果识别到本部件，则将按照预设格式调整格式。",
    },
}


# 用户表里原本就拆成两行的对象：“摘要”“关键词”。这两组的内容写死在这里，
# 不依赖被改坏的文件；行顺序与文字都按他的原文。
PRESERVE: dict[str, list[dict]] = {
    "摘要": [
        {"sub": "“摘要”标签",
         "D": "满足\n（包括且只包括连续的“摘要”文本）\n的段落，将会被识别为“摘要”标签。\n"
              "如果在下一个识别到的部件前不存在满足条件的段落，则判定不存在摘要。",
         "E": "允许“摘要”二字外有任意的“【”“】”“[”“]”，不要求成对",
         "F": "如果识别到本部件，则将直接在段落上按照预设格式调整格式，并强制将内容改为“【摘要】”；\n"
              "未识别到题目且该部件启用，则自动插入占位文本。"},
        {"sub": "摘要正文",
         "D": "满足\n（紧贴“摘要”标签之后）且（位于下一个部件标签之前）\n的段落，将会被判定为摘要正文。",
         "E": None,
         "F": "如果识别到本部件，则将直接在段落上按照预设格式调整格式；\n"
              "未识别到题目且该部件启用，则自动插入占位文本。"},
    ],
    "关键词": [
        {"sub": "“关键词”标签",
         "D": "满足\n（包括且只包括连续的“关键词”文本）\n的段落，将会被识别为“关键词”标签。\n"
              "如果在下一个识别到的部件前不存在满足条件的段落，则判定不存在关键词。",
         "E": "允许“关键词”三字外存在任意的“【”“】”“[”“]”，不要求成对。",
         "F": "如果识别到本部件，则将直接在段落上按照预设格式调整格式，并强制将内容改为“【关键词】”；\n"
              "未识别到关键词且该部件启用，则自动插入占位文本。"},
        {"sub": "关键词正文",
         "D": "满足\n（紧贴“关键词”标签之后）且（位于下一个部件标签之前）\n的段落，将会被判定为关键词正文。",
         "E": None,
         "F": "如果识别到本部件，则将直接在段落上按照预设格式调整格式；\n"
              "未识别到关键词正文且该部件启用，则自动插入占位文本。"},
    ],
}


def read_existing(sheet) -> list[dict]:
    """把用户现有的行读成 [{no, obj, sub, D, E, F}]，文字原样取出。"""
    records: list[dict] = []
    current: dict | None = None
    for row in range(2, sheet.max_row + 1):
        no = sheet.cell(row=row, column=1).value
        obj = sheet.cell(row=row, column=2).value
        sub = sheet.cell(row=row, column=3).value
        if obj:
            current = {"no": no, "obj": str(obj), "rows": []}
            records.append(current)
        if current is None:
            continue
        current["rows"].append({
            "sub": (str(sub) if sub else None),
            "D": sheet.cell(row=row, column=4).value,
            "E": sheet.cell(row=row, column=5).value,
            "F": sheet.cell(row=row, column=6).value,
            "G": sheet.cell(row=row, column=7).value,
        })
    return records


def main() -> None:
    BACKUP.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(BOOK, BACKUP)

    book = openpyxl.load_workbook(BOOK)
    sheet = book[SHEET]
    records = read_existing(sheet)

    # 样式取自用户表头与已有内容格，保证重排后观感一致
    styles = {
        "A": copy.copy(sheet["A2"]._style),
        "B": copy.copy(sheet["B2"]._style),
        "C": copy.copy(sheet["C4"]._style),
        "D": copy.copy(sheet["D2"]._style),
        "E": copy.copy(sheet["E4"]._style),
        "F": copy.copy(sheet["F2"]._style),
        "G": copy.copy(sheet["G2"]._style),
    }

    for rng in list(sheet.merged_cells.ranges):
        sheet.unmerge_cells(str(rng))
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, max_col=7):
        for cell in row:
            cell.value = None

    plan: list[dict] = []
    for record in records:
        obj = record["obj"]
        if obj in PRESERVE:
            for row in PRESERVE[obj]:
                plan.append({"no": record["no"], "obj": obj, "sub": row["sub"],
                             "D": row["D"], "E": row["E"], "F": row["F"],
                             "G": None})
            continue
        subs = SPLIT.get(obj)
        if subs is None:
            # 用户本来就拆成多行的（摘要、关键词），原样保留他那一行一行
            for row in record["rows"]:
                plan.append({"no": record["no"], "obj": obj, "sub": row["sub"],
                             "D": row["D"], "E": row["E"], "F": row["F"],
                             "G": row["G"]})
            continue
        for name in subs:
            text = SUB_TEXT.get((obj, name), {})
            plan.append({"no": record["no"], "obj": obj, "sub": name,
                         "D": text.get("D"), "E": text.get("E"), "F": text.get("F"),
                         "G": None})

    row = 2
    first_row: dict[str, int] = {}
    last_row: dict[str, int] = {}
    order: list[str] = []
    for item in plan:
        if item["obj"] not in first_row:
            first_row[item["obj"]] = row
            order.append(item["obj"])
        last_row[item["obj"]] = row
        sheet.cell(row=row, column=1, value=item["no"]).value = item["no"]
        sheet.cell(row=row, column=2, value=item["obj"])
        if item["sub"] is not None:
            sheet.cell(row=row, column=3, value=item["sub"])
        for column, key in (("D", "D"), ("E", "E"), ("F", "F"), ("G", "G")):
            if item[key] is not None:
                sheet.cell(row=row, column=ord(column) - 64, value=item[key])
        for column in "ABCDEFG":
            sheet[f"{column}{row}"]._style = styles[column]
        row += 1

    for name in order:
        start, stop = first_row[name], last_row[name]
        if start == stop:
            sheet.merge_cells(start_row=start, start_column=2, end_row=stop, end_column=3)
        elif stop > start:
            sheet.merge_cells(start_row=start, start_column=2, end_row=stop, end_column=2)

    last = row - 1
    for r in range(2, last + 1):
        # 行高置空即"自动"（openpyxl 的 customHeight 是只读派生属性，不能直接写）
        sheet.row_dimensions[r].height = None

    # 用户原文的 F2 开头有一个空格；openpyxl 重存工作簿时会把这种开头空格吃掉，
    # 这里原样补回来。
    f2 = sheet["F2"].value
    if isinstance(f2, str) and not f2.startswith(" "):
        sheet["F2"] = " " + f2

    book.save(BOOK)
    MIRROR.parent.mkdir(parents=True, exist_ok=True)
    book.save(MIRROR)

    print(f"已重排：{BOOK.relative_to(ROOT)}")
    print(f"已另存：{MIRROR.relative_to(ROOT)}")
    print(f"  对象 {len(order)} 个、表行 {len(plan)} 行（第 2~{last} 行）；行高已设为自动")
    print(f"  备份：{BACKUP.relative_to(ROOT)}")
    split_names = [n for n in order if last_row[n] > first_row[n]]
    print(f"  拆成多行的对象：{split_names}")


if __name__ == "__main__":
    main()
