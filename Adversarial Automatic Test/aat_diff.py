r"""差分器：比两份 docx 的"逐段实际值"，出人到得了现场的报告。

两种用法：
  1. **两个 docx 互相比** —— 不需要任何期望值。改代码前后 / 换预设 / 翻一个开关，
     都能用。已实证：同源同预设跑两遍，22 个部件里 21 个逐字节相同、纯噪声 0，
     所以"有差异"就是真差异。
  2. **一个 docx 与期望值比** —— 期望值由别的模块算出，也走这里的报告格式。

比对单位（对齐规则，必须先声明清楚）：
  * 段：`w:body` 直接子元素里数出来的 `w:p` 下标。用 `difflib` 按**文字**对齐，
    这样"中间插了一段"不会让后面所有段都报成"变了"。
  * 表：按 (表序, 行, 列) 对齐。
  * 页眉/页脚/脚注：各自按出现顺序对齐。
  * 节：按出现顺序，逐字段比。

只读，不改任何东西。
"""

from __future__ import annotations

import difflib
import pathlib
from dataclasses import dataclass, field

from aat_docx import (DocShape, Para, effective_para_format, effective_run_format,
                      read_docx)


@dataclass
class FieldChange:
    where: str            # 段号 / 表-行-列 / 节n
    text: str             # 该段原文（截断）
    changes: dict[str, tuple[str, str]] = field(default_factory=dict)
    note: str = ""        # "新增段" / "删除段"

    def is_empty(self) -> bool:
        return not self.changes and not self.note


@dataclass
class DiffResult:
    left: pathlib.Path
    right: pathlib.Path
    part_changes: list[tuple[str, str, str]] = field(default_factory=list)
    section_changes: list[FieldChange] = field(default_factory=list)
    style_changes: list[FieldChange] = field(default_factory=list)
    body_changes: list[FieldChange] = field(default_factory=list)
    table_changes: list[FieldChange] = field(default_factory=list)
    header_changes: list[FieldChange] = field(default_factory=list)
    footer_changes: list[FieldChange] = field(default_factory=list)
    footnote_changes: list[FieldChange] = field(default_factory=list)
    counts: dict[str, int] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)

    @property
    def identical(self) -> bool:
        return not (self.part_changes or self.section_changes or self.style_changes
                    or self.body_changes
                    or self.table_changes or self.header_changes
                    or self.footer_changes or self.footnote_changes)


# ------------------------------------------------------------------ 逐段比


def _run_format_summary(para: Para) -> str:
    """这一段的 run 上还挂着哪些字符格式标签（看残留用）。

    只报"有没有"，不报每一段的值——同一段里 run 之间常常不一致，
    逐 run 铺开会让报告淹掉。
    """
    tags: set[str] = set()
    for _text, rpr in para.runs:
        for key in ("b", "i", "u", "vertAlign", "color", "rFonts", "sz"):
            if key in rpr:
                value = rpr[key].get("val", "")
                if key == "vertAlign" and not value:
                    continue
                tags.add(key if key != "vertAlign" else f"vertAlign={value}")
    return ",".join(sorted(tags))


def compare_para(left: Para, right: Para, shape_left: DocShape,
                 shape_right: DocShape, *, compare_text: bool = True) -> dict[str, tuple[str, str]]:
    """比两个段落，返回 `{字段: (左值, 右值)}`。"""
    changes: dict[str, tuple[str, str]] = {}

    def note(name: str, before: str, after: str) -> None:
        if before != after:
            changes[name] = (before, after)

    if compare_text:
        note("文字", left.text.strip(), right.text.strip())
    note("样式", f"{left.style_id}/{left.style_name}",
         f"{right.style_id}/{right.style_name}")

    left_para = effective_para_format(shape_left, left)
    right_para = effective_para_format(shape_right, right)
    for key in sorted(set(left_para) | set(right_para)):
        note(f"段落·{key}", left_para.get(key, ""), right_para.get(key, ""))

    left_run = effective_run_format(shape_left, left)
    right_run = effective_run_format(shape_right, right)
    for key in sorted(set(left_run) | set(right_run)):
        note(f"字符·{key}", left_run.get(key, ""), right_run.get(key, ""))

    note("run格式标签", _run_format_summary(left), _run_format_summary(right))
    note("域", " | ".join(left.fields), " | ".join(right.fields))
    note("含图", f"drawing={left.has_drawing} pict={left.has_pict}",
         f"drawing={right.has_drawing} pict={right.has_pict}")
    return changes


def align_and_compare(left_items: list[Para], right_items: list[Para],
                      shape_left: DocShape, shape_right: DocShape,
                      where_prefix: str, *, compare_text: bool = True,
                      limit: int = 400) -> list[FieldChange]:
    """按文字对齐两串段落，逐对比较。"""
    matcher = difflib.SequenceMatcher(
        None, [p.text.strip() for p in left_items], [p.text.strip() for p in right_items],
        autojunk=False)
    out: list[FieldChange] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for offset in range(i2 - i1):
                changes = compare_para(left_items[i1 + offset], right_items[j1 + offset],
                                       shape_left, shape_right, compare_text=compare_text)
                if changes:
                    para = right_items[j1 + offset]
                    out.append(FieldChange(
                        where=f"{where_prefix}{i1 + offset}", text=para.text[:40],
                        changes=changes))
        elif tag == "replace":
            for offset in range(max(i2 - i1, j2 - j1)):
                left_para = left_items[i1 + offset] if i1 + offset < i2 else None
                right_para = right_items[j1 + offset] if j1 + offset < j2 else None
                if left_para is None:
                    out.append(FieldChange(where=f"{where_prefix}{j1 + offset}",
                                           text=right_para.text[:40], note="新增段"))
                elif right_para is None:
                    out.append(FieldChange(where=f"{where_prefix}{i1 + offset}",
                                           text=left_para.text[:40], note="删除段"))
                else:
                    changes = compare_para(left_para, right_para, shape_left,
                                           shape_right, compare_text=compare_text)
                    if changes:
                        out.append(FieldChange(where=f"{where_prefix}{i1 + offset}",
                                               text=right_para.text[:40],
                                               changes=changes))
        else:                                       # delete / insert
            source = left_items if tag == "delete" else right_items
            start, stop = (i1, i2) if tag == "delete" else (j1, j2)
            for offset in range(start, stop):
                out.append(FieldChange(where=f"{where_prefix}{offset}",
                                       text=source[offset].text[:40],
                                       note="删除段" if tag == "delete" else "新增段"))
        if len(out) > limit:
            out.append(FieldChange(where="…", text="", note=f"差异过多，已截断（>{limit}）"))
            break
    return out


# ------------------------------------------------------------------ 整体比


def compare_shapes(left: DocShape, right: DocShape, *,
                   compare_text: bool = True) -> DiffResult:
    result = DiffResult(left=left.path, right=right.path)
    result.warnings = sorted(set(left.warnings) | set(right.warnings))

    for name in sorted(set(left.parts) | set(right.parts)):
        digest_left = left.parts.get(name, "（缺）")
        digest_right = right.parts.get(name, "（缺）")
        if digest_left != digest_right and name != "docProps/core.xml":
            result.part_changes.append((name, digest_left, digest_right))

    result.section_changes = align_sections(left, right)
    result.style_changes = align_styles(left, right)
    result.body_changes = align_and_compare(left.paras, right.paras, left, right,
                                            "段", compare_text=compare_text)
    result.header_changes = align_and_compare(left.header_paras, right.header_paras,
                                              left, right, "页眉段",
                                              compare_text=compare_text)
    result.footer_changes = align_and_compare(left.footer_paras, right.footer_paras,
                                              left, right, "页脚段",
                                              compare_text=compare_text)
    result.footnote_changes = align_and_compare(left.footnote_paras,
                                                right.footnote_paras, left, right,
                                                "脚注段", compare_text=compare_text)
    result.table_changes = align_tables(left, right, compare_text=compare_text)

    result.counts = {
        "部件": len(result.part_changes),
        "节": len(result.section_changes),
        "样式": len(result.style_changes),
        "正文段": len(result.body_changes),
        "表段": len(result.table_changes),
        "页眉段": len(result.header_changes),
        "页脚段": len(result.footer_changes),
        "脚注段": len(result.footnote_changes),
        "新增": sum(1 for item in result.body_changes if item.note == "新增段"),
        "删除": sum(1 for item in result.body_changes if item.note == "删除段"),
    }
    if left.table_repeat_header != right.table_repeat_header:
        result.counts["首页标题行重复"] = 1
    return result


def align_sections(left: DocShape, right: DocShape) -> list[FieldChange]:
    """比节设置。

    对齐规矩：**从末尾往前配**——最后一节是正文节，位置最稳；前面的节会因为
    插入/删除而增删（实测：勾了"重新为文档生成"之后节数会变），按序号从前
    往后配会把"前置"跟"正文"对上，报出一堆假差异。
    """
    out: list[FieldChange] = []
    count_left = len(left.sections)
    count_right = len(right.sections)
    if count_left != count_right:
        out.append(FieldChange(where="节数", text="",
                               note=f"{count_left} → {count_right}"
                                    f"（后面的节只能从末尾对齐着看）"))
    same_count = count_left == count_right
    for offset in range(1, min(count_left, count_right) + 1):
        index = count_left - offset
        other = count_right - offset
        left_section = left.sections[index]
        right_section = right.sections[other]
        changes = {
            key: (left_section.get(key, ""), right_section.get(key, ""))
            for key in sorted(set(left_section) | set(right_section))
            if left_section.get(key, "") != right_section.get(key, "")
        }
        if changes:
            where = (f"节{index}" if same_count
                     else f"节（左{index} / 右{other}）")
            out.append(FieldChange(where=where, text="", changes=changes))
    return out


def align_styles(left: DocShape, right: DocShape) -> list[FieldChange]:
    """比 styles.xml 里的样式定义（格式页改一个值，这里就该看到那一处变）。"""
    def flatten(cell: dict[str, dict[str, str]]) -> str:
        return ";".join(f"{k}={';'.join(f'{a}={b}' for a, b in sorted(v.items()))}"
                        for k, v in sorted((cell or {}).items()))

    out: list[FieldChange] = []
    for style_id in sorted(set(left.styles) | set(right.styles)):
        before = left.styles.get(style_id)
        after = right.styles.get(style_id)
        if before is None or after is None:
            present = before or after
            out.append(FieldChange(where=f"样式 {style_id}/{present.get('name', '')}",
                                   text="", note="样式新增" if before is None
                                   else "样式删除"))
            continue
        changes: dict[str, tuple[str, str]] = {}
        if before.get("name") != after.get("name"):
            changes["名称"] = (before.get("name", ""), after.get("name", ""))
        if before.get("basedOn") != after.get("basedOn"):
            changes["basedOn"] = (before.get("basedOn", ""), after.get("basedOn", ""))
        for group in ("pPr", "rPr"):
            keys = set(before.get(group, {})) | set(after.get(group, {}))
            for key in sorted(keys):
                before_value = flatten({key: before.get(group, {}).get(key, {})})
                after_value = flatten({key: after.get(group, {}).get(key, {})})
                if before_value != after_value:
                    changes[f"{group}·{key}"] = (before_value, after_value)
        if changes:
            out.append(FieldChange(where=f"样式 {style_id}/{after.get('name', '')}",
                                   text="", changes=changes))
    return out


def align_tables(left: DocShape, right: DocShape, *, compare_text: bool = True
                 ) -> list[FieldChange]:
    out: list[FieldChange] = []
    for table_index in range(max(len(left.tables), len(right.tables))):
        left_table = left.tables[table_index] if table_index < len(left.tables) else []
        right_table = right.tables[table_index] if table_index < len(right.tables) else []
        for row_index in range(max(len(left_table), len(right_table))):
            left_row = left_table[row_index] if row_index < len(left_table) else []
            right_row = right_table[row_index] if row_index < len(right_table) else []
            for cell_index in range(max(len(left_row), len(right_row))):
                left_cell = left_row[cell_index] if cell_index < len(left_row) else None
                right_cell = (right_row[cell_index] if cell_index < len(right_row)
                              else None)
                where = f"表{table_index}-{row_index}-{cell_index}"
                if left_cell is None or right_cell is None:
                    present = left_cell or right_cell
                    out.append(FieldChange(where=where, text=present.text[:40],
                                           note="单元格新增" if left_cell is None
                                           else "单元格删除"))
                    continue
                changes = compare_para(left_cell, right_cell, left, right,
                                       compare_text=compare_text)
                if changes:
                    out.append(FieldChange(where=where, text=right_cell.text[:40],
                                           changes=changes))
    return out


# ------------------------------------------------------------------ 报告


def render(result: DiffResult, *, max_items: int = 40, detail: bool = True) -> str:
    lines: list[str] = []
    lines.append("=" * 78)
    lines.append(f"左：{result.left}")
    lines.append(f"右：{result.right}")
    lines.append("=" * 78)
    for warning in result.warnings:
        lines.append(f"⚠️ {warning}")
    if result.identical:
        lines.append("结论：**两份产物没有差异**")
        return "\n".join(lines)

    summary = "、".join(f"{key} {value}" for key, value in result.counts.items()
                        if value)
    lines.append(f"差异概览：{summary}")
    if result.part_changes:
        lines.append("")
        lines.append("变了哪些包内部件：")
        for name, digest_left, digest_right in result.part_changes:
            lines.append(f"  {name:<34} {digest_left} → {digest_right}")

    for title, items in (("节设置", result.section_changes),
                         ("样式定义", result.style_changes),
                         ("正文", result.body_changes),
                         ("表格", result.table_changes),
                         ("页眉", result.header_changes),
                         ("页脚", result.footer_changes),
                         ("脚注", result.footnote_changes)):
        if not items:
            continue
        lines.append("")
        lines.append(f"—— {title}（{len(items)} 处）")
        for item in items[:max_items]:
            if item.note:
                lines.append(f"  {item.where}  [{item.note}]  {item.text!r}")
                continue
            lines.append(f"  {item.where}  {item.text!r}")
            if detail:
                for name, (before, after) in sorted(item.changes.items()):
                    lines.append(f"      {name:<16} {before or '（空）'}"
                                 f"  →  {after or '（空）'}")
        if len(items) > max_items:
            lines.append(f"  …（还有 {len(items) - max_items} 处，见 JSON）")
    return "\n".join(lines)


def diff_files(left_path, right_path, *, compare_text: bool = True) -> DiffResult:
    return compare_shapes(read_docx(pathlib.Path(left_path)),
                          read_docx(pathlib.Path(right_path)),
                          compare_text=compare_text)


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 3:
        raise SystemExit("用法：aat_diff.py <左docx> <右docx> [--no-text]")
    outcome = diff_files(sys.argv[1], sys.argv[2],
                         compare_text="--no-text" not in sys.argv)
    print(render(outcome))
