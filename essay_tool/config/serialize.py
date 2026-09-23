"""PresetConfig ↔ 可以直接 `json.dump` 的 dict。

字段里能存"人读得懂的写法"的都存那种写法（中文字号名、"居中"这类对齐名、
"是"/"否"……），跟 Excel 用的是**同一套词汇**；解析统一复用 `units.py`
已经写好的 `parse_*` 函数，不再维护第二份校验逻辑——这样 JSON 和 Excel
两条路径认的写法保证一致，也顺带保证了两条路径互转（导入 Excel 存成
JSON、把 JSON 导出成 Excel）不会走样。
"""

from __future__ import annotations

from docx.enum.text import WD_LINE_SPACING

from essay_tool.config import units
from essay_tool.config.schema import (
    CitationConfig,
    HeaderFooterRule,
    PageSetup,
    ParagraphStyleSpec,
    PresetConfig,
)

SCHEMA_VERSION = 1

_ALIGN_TO_LABEL = {v: k for k, v in units.ALIGNMENTS.items()}
_LINE_RULE_TO_LABEL = {v: k for k, v in units.LINE_RULES.items()}
_PT_TO_CN_SIZE = {v: k for k, v in units.CHINESE_FONT_SIZES.items()}


def _size_to_str(size) -> str:
    pt = size.pt
    return _PT_TO_CN_SIZE.get(pt, f"{pt:g}")


def _indent_to_pair(ind) -> tuple[str, str]:
    if ind.kind is units.IndentKind.CHARS:
        return "字符", f"{ind.chars:g}"
    if ind.kind is units.IndentKind.LENGTH and ind.length is not None:
        return "厘米", f"{ind.length.cm:g}"
    return "-", "0"


def _len_pt_to_str(v) -> str:
    return "" if v is None else f"{v.pt:g}"


def _line_value_to_str(rule, value) -> str:
    if value is None:
        return ""
    if rule == WD_LINE_SPACING.MULTIPLE:
        return f"{float(value):g}"
    return f"{value.pt:g}"


# ------------------------------------------------------------ 样式（单行）

def style_spec_to_dict(spec: ParagraphStyleSpec) -> dict:
    fl_unit, fl_val = _indent_to_pair(spec.first_line_indent)
    li_unit, li_val = _indent_to_pair(spec.left_indent)
    return {
        "id": spec.id,
        "name": spec.name,
        "font_east_asian": spec.font_east_asian,
        "font_ascii": spec.font_ascii,
        "size": _size_to_str(spec.size),
        "bold": "是" if spec.bold else "否",
        "alignment": _ALIGN_TO_LABEL.get(spec.alignment, "两端对齐"),
        "first_line_indent_unit": fl_unit,
        "first_line_indent_value": fl_val,
        "left_indent_unit": li_unit,
        "left_indent_value": li_val,
        "space_before": _len_pt_to_str(spec.space_before),
        "space_after": _len_pt_to_str(spec.space_after),
        "line_rule": _LINE_RULE_TO_LABEL.get(spec.line_rule, "固定值（磅）"),
        "line_value": _line_value_to_str(spec.line_rule, spec.line_value),
        "page_break_before": "是" if spec.page_break_before else "否",
    }


def style_spec_from_dict(d: dict) -> ParagraphStyleSpec:
    line_rule = units.parse_line_rule(d["line_rule"])
    line_value = units.parse_line_value(line_rule, d.get("line_value", ""))
    return ParagraphStyleSpec(
        id=d["id"],
        name=d["name"],
        font_east_asian=d["font_east_asian"],
        font_ascii=d["font_ascii"],
        size=units.parse_font_size(d["size"]),
        bold=units.parse_yn(d["bold"]),
        alignment=units.parse_alignment(d["alignment"]),
        first_line_indent=units.parse_indent_pair(
            d["first_line_indent_unit"], d["first_line_indent_value"]),
        left_indent=units.parse_indent_pair(
            d["left_indent_unit"], d["left_indent_value"]),
        space_before=units.parse_optional_length_pair("磅", d.get("space_before", "")),
        space_after=units.parse_optional_length_pair("磅", d.get("space_after", "")),
        line_rule=line_rule,
        line_value=line_value,
        page_break_before=units.parse_yn(d["page_break_before"]),
    )


# ------------------------------------------------------------ 页面设置

def page_setup_to_dict(p: PageSetup) -> dict:
    return {
        "paper_size": p.paper_size,
        "page_width_cm": f"{p.page_width.cm:g}",
        "page_height_cm": f"{p.page_height.cm:g}",
        "orientation": p.orientation,
        "margin_top_cm": f"{p.margin_top.cm:g}",
        "margin_bottom_cm": f"{p.margin_bottom.cm:g}",
        "margin_left_cm": f"{p.margin_left.cm:g}",
        "margin_right_cm": f"{p.margin_right.cm:g}",
        "gutter_cm": f"{p.gutter.cm:g}",
        "header_distance_cm": f"{p.header_distance.cm:g}",
        "footer_distance_cm": f"{p.footer_distance.cm:g}",
    }


def page_setup_from_dict(d: dict) -> PageSetup:
    paper_choice = d["paper_size"]
    if paper_choice in units.NAMED_PAPER_SIZES:
        page_width, page_height = units.NAMED_PAPER_SIZES[paper_choice]
    else:
        page_width = units.parse_cm(d["page_width_cm"])
        page_height = units.parse_cm(d["page_height_cm"])
    return PageSetup(
        paper_size=paper_choice,
        page_width=page_width,
        page_height=page_height,
        orientation=d["orientation"],
        margin_top=units.parse_cm(d["margin_top_cm"]),
        margin_bottom=units.parse_cm(d["margin_bottom_cm"]),
        margin_left=units.parse_cm(d["margin_left_cm"]),
        margin_right=units.parse_cm(d["margin_right_cm"]),
        gutter=units.parse_cm(d["gutter_cm"]),
        header_distance=units.parse_cm(d["header_distance_cm"]),
        footer_distance=units.parse_cm(d["footer_distance_cm"]),
    )


# ------------------------------------------------------------ 页眉页脚

def header_footer_rule_to_dict(r: HeaderFooterRule) -> dict:
    return {
        "section_name": r.section_name,
        "header_text": r.header_text,
        "page_number_enabled": "是" if r.page_number_enabled else "否",
        "page_number_format": r.page_number_format,
        "page_number_prefix": r.page_number_prefix,
        "page_number_suffix": r.page_number_suffix,
    }


def header_footer_rule_from_dict(d: dict) -> HeaderFooterRule:
    enabled = units.parse_yn(d["page_number_enabled"])
    return HeaderFooterRule(
        section_name=d["section_name"],
        header_text=d.get("header_text", ""),
        page_number_enabled=enabled,
        page_number_format=d.get("page_number_format") if enabled else None,
        page_number_prefix=d.get("page_number_prefix", ""),
        page_number_suffix=d.get("page_number_suffix", ""),
    )


# ------------------------------------------------------------ 引注

def citation_to_dict(c: CitationConfig | None) -> dict | None:
    if c is None:
        return None
    return {
        "footnote_marker_style": c.footnote_marker_style,
        "footnote_restart_each_page": "是" if c.footnote_restart_each_page else "否",
        "footnote_marker_font": c.footnote_marker_font,
        "footnote_marker_trailing_space": c.footnote_marker_trailing_space,
        "reference_numbering_style": c.reference_numbering_style,
    }


def citation_from_dict(d: dict | None) -> CitationConfig | None:
    if not d:
        return None
    return CitationConfig(
        footnote_marker_style=d["footnote_marker_style"],
        footnote_restart_each_page=units.parse_yn(d["footnote_restart_each_page"]),
        footnote_marker_font=d["footnote_marker_font"],
        footnote_marker_trailing_space=d["footnote_marker_trailing_space"],
        reference_numbering_style=d["reference_numbering_style"],
    )


# ------------------------------------------------------------ 整份预设

def preset_to_dict(cfg: PresetConfig) -> dict:
    return {
        "schema_version": SCHEMA_VERSION,
        "page": page_setup_to_dict(cfg.page),
        "styles": [style_spec_to_dict(s) for s in cfg.styles.values()],
        "header_footer": [header_footer_rule_to_dict(r) for r in cfg.header_footer],
        "citation": citation_to_dict(cfg.citation),
        "structure_required": dict(cfg.structure_required),
        "structure_insert_text": dict(cfg.structure_insert_text),
        # None（没配，代码走默认值）和 []（用户明确写"无"）含义不同，见 schema.py
        # 里 keep_run_formatting 的注释——这里原样保留这个区分，不能合并成一种。
        "keep_run_formatting": (
            sorted(cfg.keep_run_formatting) if cfg.keep_run_formatting is not None else None
        ),
        "header_bottom_line": cfg.header_bottom_line,
        "header_line_extend_left_cm": cfg.header_line_extend_left_cm,
        "header_line_extend_right_cm": cfg.header_line_extend_right_cm,
        "cover_hanging_indent": cfg.cover_hanging_indent,
        "cover_enabled": cfg.cover_enabled,
        "cover_form": cfg.cover_form,
        "clear_page_breaks": cfg.clear_page_breaks,
        "clear_blank_lines": cfg.clear_blank_lines,
        "label_page_break": dict(cfg.label_page_break),
    }


def preset_from_dict(d: dict) -> PresetConfig:
    styles_by_id: dict[str, ParagraphStyleSpec] = {}
    styles_by_name: dict[str, ParagraphStyleSpec] = {}
    for sd in d.get("styles", []):
        spec = style_spec_from_dict(sd)
        styles_by_id[spec.id] = spec
        styles_by_name[spec.name] = spec
    keep = d.get("keep_run_formatting")
    return PresetConfig(
        page=page_setup_from_dict(d["page"]),
        styles=styles_by_id,
        styles_by_name=styles_by_name,
        header_footer=[header_footer_rule_from_dict(x) for x in d.get("header_footer", [])],
        citation=citation_from_dict(d.get("citation")),
        structure_required=dict(d.get("structure_required", {})),
        structure_insert_text=dict(d.get("structure_insert_text", {})),
        keep_run_formatting=frozenset(keep) if keep is not None else None,
        header_bottom_line=bool(d.get("header_bottom_line", False)),
        # 旧预设没有这两项 → 0（不伸出，老行为）
        header_line_extend_left_cm=float(d.get("header_line_extend_left_cm", 0.0)),
        header_line_extend_right_cm=float(d.get("header_line_extend_right_cm", 0.0)),
        cover_hanging_indent=bool(d.get("cover_hanging_indent", True)),
        # 旧预设里没这两项 → 保持 None，由调用方（GUI）按老行为处理
        cover_enabled=(None if d.get("cover_enabled") is None else bool(d["cover_enabled"])),
        cover_form=d.get("cover_form"),
        # 旧预设没有这三项 → 都是"不清理、标签跟随样式"（老行为）
        clear_page_breaks=bool(d.get("clear_page_breaks", False)),
        clear_blank_lines=bool(d.get("clear_blank_lines", False)),
        label_page_break={str(k): bool(v) for k, v in (d.get("label_page_break") or {}).items()},
    )
