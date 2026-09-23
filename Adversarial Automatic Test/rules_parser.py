r"""规则表解析器（**2026-09-20 按重写后的窗口正文重建**）。

真源 = `essay_tool/gui/operation_logic_text.py` 的 `OPERATION_LOGIC_TEXT`
（上游真源，**只读**——另一侧还在改它）。

新版正文里的表**不止三张**，而且识别侧拆成了三个识别器各一张表。所以这份
解析器不再按"小节标题"给前缀，而是**按表的列签名 + 识别对象的内容**判定
它是哪张表，再给前缀：

    F  3.6.2（1）前向识别器·标签表（摘要/关键词/英文摘要/英文关键词）
    T  3.6.2（1）前向识别器·题目表（中文正标题/中文副标题/英文正标题/英文副标题）
    B  3.6.2（2）后向识别器·标签表（参考文献/附录/致谢）
    H  3.6.2（2）后向识别器·标题表（一～四级标题/引言/结论）
    S  3.6.2（3）补充识别器表（题注/目录三清单标签/目录三清单域/附录序号题名）
    P  3.6.3 格式应用表（样式应用规则）
    A  3.7 自动告警表
    M  6.2 其他告警表

ID = 前缀 + 表内「序号」（如 `F1`、`H3`、`S4`、`A10`）。用户明确：编号随便编，
只要稳定、报告里能指得回表里哪一行。

产物（本文件夹 `_out/` 下）：rule_tables.json / rule_tables.txt / rule_health.txt。
"""

from __future__ import annotations

import hashlib
import importlib.util
import json
import pathlib
import re
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent
SOURCE = ROOT / "essay_tool" / "gui" / "operation_logic_text.py"
OUT = HERE / "_out"


def load_window_text() -> str:
    """执行那份模块文件、取走 OPERATION_LOGIC_TEXT。

    用 importlib 按路径直接加载，**不走 `essay_tool.gui` 包**——免得为了读一段
    文本而把 tkinter 那一串拉进来。
    """
    spec = importlib.util.spec_from_file_location("_aat_operation_logic_text", SOURCE)
    if spec is None or spec.loader is None:      # pragma: no cover - 只在文件被删时
        raise SystemExit(f"读不到真源：{SOURCE}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.OPERATION_LOGIC_TEXT


def is_table_line(line: str) -> bool:
    return line.lstrip().startswith("|")


# 小节标题，如 `3.6.2 文档的识别`、`6.2 其他告警`。
# （窗口正文沿用生成器那套约定：标题里编号与文字之间必须有空格。）
SECTION_HEADING_RE = re.compile(r"^\d+(?:\.\d+)*\s+\S")


def heading_for(lines: list[str], table_line_index: int) -> str:
    """表格所属的小节标题：往上找**最近的小节标题行**；找不到才退回普通文字行。"""
    fallback = ""
    for back in range(table_line_index - 1, -1, -1):
        candidate = lines[back].strip()
        if not candidate or is_table_line(candidate):
            continue
        if SECTION_HEADING_RE.match(candidate):
            return candidate
        if not fallback:
            fallback = candidate
    return fallback


def is_separator_line(line: str) -> bool:
    """Markdown 表格的分隔行：只有 | - : 和空格。"""
    body = line.strip().strip("|")
    return bool(body) and all(ch in "|-: " for ch in body)


def split_row(line: str) -> list[str]:
    """一行拆成单元格。首尾的竖线去掉，每个格子 strip。

    ⚠️ 约定：正文里的表格**没有转义竖线 `\\|`**。以后若出现，这里会切多一格，
    `health` 里的"列数不一致"会立刻报出来。
    """
    return [cell.strip() for cell in line.strip().strip("|").split("|")]


def find_tables(text: str) -> list[dict]:
    """按行扫出所有 Markdown 管道表，连同它上面的小标题。"""
    lines = text.splitlines()
    tables: list[dict] = []
    index = 0
    while index < len(lines):
        if not is_table_line(lines[index]) or not is_separator_line(
            lines[min(index + 1, len(lines) - 1)]
        ):
            index += 1
            continue
        heading = heading_for(lines, index)
        headers = split_row(lines[index])
        rows: list[dict] = []
        cursor = index + 2
        while cursor < len(lines) and is_table_line(lines[cursor]):
            cells = split_row(lines[cursor])
            rows.append({"line": cursor + 1, "cells": cells})
            cursor += 1
        tables.append({
            "heading": heading,
            "start_line": index + 1,
            "headers": headers,
            "rows": rows,
        })
        index = cursor
    return tables


def classify(table: dict) -> str:
    """判定这张表是哪一张（见模块开头的前缀表）。只认**列签名 + 识别对象**，不认行号。"""
    headers = [header.strip() for header in table["headers"]]
    columns = set(headers)
    if "告警名称" in columns:
        heading = table["heading"]
        if "3.7" in heading:
            return "A"
        if "6.2" in heading:
            return "M"
        return "W"
    if "样式应用规则" in columns:
        return "P"
    if "识别规则" in columns:
        names = ""
        for row in table["rows"]:
            cells = dict(zip(table["headers"], row["cells"]))
            names += cells.get("识别对象", "")
        if "题注" in names or "附录序号题名" in names:
            return "S"
        if "正标题" in names or "副标题" in names:
            return "T"
        if any(token in names for token in ("一级标题", "二级标题", "三级标题", "四级标题",
                                            "引言", "结论")):
            return "H"
        if "参考文献" in names or "致谢" in names or "附录" in names:
            return "B"
        if "摘要" in names or "关键词" in names:
            return "F"
    return "U"


def rule_id(prefix: str, row: dict, headers: list[str]) -> str:
    """稳定 ID：前缀 + 「序号」列。样式应用表（P）里同一个序号会用「子项」
    （标签行 / 正文行）拆成两行、甚至三行，所以带上子项才不会撞。"""
    cells = dict(zip(headers, row["cells"]))
    ordinal = (cells.get("序号") or "").strip()
    sub_item = (cells.get("子项") or "").replace("“", "").replace("”", "").strip()
    if sub_item:
        return f"{prefix}{ordinal}.{sub_item}" if ordinal else f"{prefix}@{row['line']}.{sub_item}"
    return f"{prefix}{ordinal}" if ordinal else f"{prefix}@{row['line']}"


def pretty(cell: str) -> str:
    """表格里换行写成字面量 `\\n`（沿用生成器的老约定），显示时还原。"""
    return cell.replace("\\n", "\n")


# ---------------------------------------------------------------- 体检

# 受关注的中文数字字样（识别表会点名它们；2026-09-20 版是"以中文数字开头"
# 大原则 + "百、千不能作首字 / 廿、卅不被承认"两条禁令）。
NUMERAL_CHARS = "零一二三四五六七八九十百千万廿卅两〇亿"


def numeral_tokens(cell: str) -> set[str]:
    """从一个单元格里抠出"被点名的中文数字字符"（带引号的单字 + 顿号连排）。"""
    found = set(re.findall(rf"[“\"]([{NUMERAL_CHARS}])[”\"]", cell))
    for run in re.findall(rf"[{NUMERAL_CHARS}]+(?:、[{NUMERAL_CHARS}]+)+", cell):
        found.update(ch for ch in run if ch in NUMERAL_CHARS)
    return found


# 各表的"必填"列
REQUIRED = {
    "F": ("识别对象", "识别规则"), "T": ("识别对象", "识别规则"),
    "B": ("识别对象", "识别规则"), "H": ("识别对象", "识别规则"),
    "S": ("识别对象", "识别规则"), "P": ("识别对象", "样式应用规则"),
    "A": ("告警名称", "触发条件", "警告内容"), "M": ("告警名称", "触发条件", "警告内容"),
}


def health(tables: list[dict], text: str) -> list[str]:
    report: list[str] = []
    add = report.append

    add("=" * 72)
    add("真源指纹")
    add("=" * 72)
    raw = SOURCE.read_bytes()
    add(f"  路径：{SOURCE.relative_to(ROOT)}")
    add(f"  md5 ：{hashlib.md5(raw).hexdigest()}")
    add(f"  大小：{len(raw):,} 字节")
    add("")

    add("=" * 72)
    add("一、表清单")
    add("=" * 72)
    seen_ids: dict[str, str] = {}
    duplicate_ids: list[str] = []
    unknown: list[str] = []
    for table in tables:
        prefix = classify(table)
        add(f"  [{prefix}] {table['heading']}  —— 表头 {len(table['headers'])} 列、"
            f"{len(table['rows'])} 行（起始行 {table['start_line']}）")
        add(f"      列：{table['headers']}")
        if prefix == "U":
            unknown.append(f"第 {table['start_line']} 行起、小标题「{table['heading']}」")
        for row in table["rows"]:
            rid = rule_id(prefix, row, table["headers"])
            if rid in seen_ids:
                duplicate_ids.append(rid)
            seen_ids[rid] = table["heading"]
    if unknown:
        add(f"  ✘ 有 {len(unknown)} 张表认不出是哪一张（列签名与识别对象都对不上）：")
        for item in unknown:
            add(f"      {item}")
    add(f"  合计 {len(tables)} 张表、{len(seen_ids)} 条规则")
    if duplicate_ids:
        add(f"  ✘ 重复 ID：{sorted(set(duplicate_ids))}")
    add("")

    add("=" * 72)
    add("二、列数不一致的行（多半是格子里混进了竖线）")
    add("=" * 72)
    bad_columns = 0
    for table in tables:
        expected = len(table["headers"])
        for row in table["rows"]:
            if len(row["cells"]) != expected:
                bad_columns += 1
                add(f"  ✘ 第 {row['line']} 行：{len(row['cells'])} 格，应为 {expected} 格"
                    f"  {row['cells'][:2]}")
    if not bad_columns:
        add("  （无）")
    add("")

    add("=" * 72)
    add("三、必填单元格为空的条目")
    add("=" * 72)
    empty_hits = 0
    for table in tables:
        prefix = classify(table)
        for row in table["rows"]:
            cells = dict(zip(table["headers"], row["cells"]))
            for column in REQUIRED.get(prefix, ()):
                if not (cells.get(column) or "").strip():
                    empty_hits += 1
                    add(f"  ✘ {rule_id(prefix, row, table['headers'])} 的"
                        f"「{column}」是空的")
    if not empty_hits:
        add("  （无）")
    add("")

    add("=" * 72)
    add("四、关键常量对账（60 字上限 / 中文数字字样）")
    add("=" * 72)
    prose_60 = len(re.findall(r"60 个字符", text))
    a2_text = ""
    for table in tables:
        if classify(table) == "A":
            for row in table["rows"]:
                cells = dict(zip(table["headers"], row["cells"]))
                if "疑似标题" in (cells.get("告警名称") or ""):
                    a2_text = " ".join(cells.values())
    a2_60 = "六十字" in a2_text or "60" in a2_text
    add(f"  「60 个字符」在正文里出现 {prose_60} 处；A2 疑似标题行提到 60/六十字：{'是' if a2_60 else '否'}")
    if prose_60 == 0 or not a2_60:
        add("  ✘ 60 字上限在「标题长度」说明与 A2 触发条件之间对不上——两侧任一处缺失")
    forbidden: set[str] = set()
    allowed_note: list[str] = []
    for table in tables:
        if classify(table) != "H":
            continue
        for row in table["rows"]:
            cells = dict(zip(table["headers"], row["cells"]))
            cell = cells.get("识别规则", "")
            if "不能作为" in cell or "不会被承认" in cell:
                forbidden |= numeral_tokens(cell)
            if "中文数字" in cell:
                allowed_note.append(f"{rule_id('H', row, table['headers'])}："
                                    f"{'、'.join(sorted(numeral_tokens(cell))) or '（未点名）'}")
    add(f"  识别表点名的禁用字样：{'、'.join(sorted(forbidden)) or '（无）'}")
    for line in allowed_note:
        add(f"      {line}")
    add("")

    add("=" * 72)
    add("五、识别表点名的部件")
    add("=" * 72)
    for table in tables:
        prefix = classify(table)
        if prefix not in ("F", "T", "B", "H", "S"):
            continue
        names = []
        for row in table["rows"]:
            cells = dict(zip(table["headers"], row["cells"]))
            names.append(cells.get("识别对象", "").strip())
        add(f"  [{prefix}] {'、'.join(names)}")
    add("")

    add("=" * 72)
    add("六、待人工确认")
    add("=" * 72)
    add("  体检只做**机械可判**的那几类。表与引擎是否一致，要等答案生成器")
    add("  （aat_expect.py，按同一份表字面重建）算出期望、与真实管线对照后才知道。")
    return report


def enrich_tables(tables: list[dict]) -> list[dict]:
    """`find_tables()` 的原始结果 → 强化表数据（cells 成字典 + 稳定 ID）。"""
    payload: list[dict] = []
    for table in tables:
        prefix = classify(table)
        payload.append({
            "id_prefix": prefix,
            "heading": table["heading"],
            "start_line": table["start_line"],
            "headers": table["headers"],
            "rows": [
                {
                    "id": rule_id(prefix, row, table["headers"]),
                    "line": row["line"],
                    "cells": dict(zip(table["headers"], row["cells"])),
                }
                for row in table["rows"]
            ],
        })
    return payload


def parse_tables(text: str | None = None) -> list[dict]:
    """窗口正文 → 增强表数据。

    `main()` 与其它层（`gen_exceptions`）共用这一份。**别的层别再读
    `_out/rule_tables.json`**：那是产物（`_out` 不进版本库、打包/清理会删），
    2026-09-20 清 `_out` 之后 `gen_exceptions` 当场 FileNotFoundError。
    注意：`health()` 要的是 `find_tables()` 的**原始**表（cells 是列表），不是这里。
    """
    return enrich_tables(find_tables(text if text is not None else load_window_text()))


def main() -> int:
    text = load_window_text()
    raw = find_tables(text)
    if not raw:
        print("没解析出任何表格——真源结构可能变了，先人工看一眼。")
        return 2
    payload = enrich_tables(raw)

    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / "rule_tables.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    text_lines: list[str] = []
    for table in payload:
        text_lines.append("=" * 72)
        text_lines.append(f"[{table['id_prefix']}] {table['heading']}"
                          f"  （{len(table['rows'])} 条）")
        text_lines.append("=" * 72)
        for row in table["rows"]:
            text_lines.append("")
            text_lines.append(f"── {row['id']}")
            for column, value in row["cells"].items():
                if not value.strip():
                    continue
                body = pretty(value)
                if "\n" not in body:
                    text_lines.append(f"   {column}：{body}")
                else:
                    text_lines.append(f"   {column}：")
                    for piece in body.split("\n"):
                        if piece.strip():
                            text_lines.append(f"     {piece.strip()}")
    (OUT / "rule_tables.txt").write_text("\n".join(text_lines), encoding="utf-8")

    report = health(raw, text)
    (OUT / "rule_health.txt").write_text("\n".join(report), encoding="utf-8")

    print("\n".join(report))
    print()
    print(f"已写出：{OUT.relative_to(ROOT)}\\rule_tables.json / rule_tables.txt / rule_health.txt")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
