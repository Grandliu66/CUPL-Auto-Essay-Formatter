"""读取 presets/*.xlsx（单一工作表"格式设置"，九个小节堆叠）为 PresetConfig。

按小节标题定位每一块的位置，再按表头文字找列，而不是死记列号/行号，
这样用户在表里插入/调整行列也不容易读错。
"""

from __future__ import annotations

import pathlib

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from essay_tool.config import units
from essay_tool.config.schema import (
    COVER_FORM_NAMES,
    HEADING_LABEL_PARTS,
    CitationConfig,
    HeaderFooterRule,
    PageSetup,
    ParagraphStyleSpec,
    PresetConfig,
)

SHEET_NAME = "格式设置"
# 预设里“小节”的名字。**旧预设里出现过的小节若已废弃，不要往里加**：
# `_iter_sections` 只认名单里的，名单外的整节直接跳过（例如 2026-09-16 取消的
# “题注编号”：旧 Excel 里可能还有那节，读到时不当错误、也不往 PresetConfig 放）。
SECTION_TITLES = {"页面设置", "样式", "页眉页脚", "页眉线", "引注", "清理规则", "封面", "结构部件"}
MAX_SCAN_COLS = 20


class ConfigError(Exception):
    """预设表里某个单元格写法有问题，消息里带具体单元格坐标。"""


def _ref(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


def _iter_sections(ws: Worksheet):
    """产出 (小节标题, {表头文字: 列号}, [数据行行号, ...])。"""
    row = 1
    max_row = ws.max_row
    while row <= max_row:
        title = ws.cell(row=row, column=1).value
        if title in SECTION_TITLES:
            header_row = row + 1
            header_cols: dict[str, int] = {}
            for c in range(1, MAX_SCAN_COLS + 1):
                val = ws.cell(row=header_row, column=c).value
                if val is None:
                    break
                header_cols[str(val).strip()] = c
            r = header_row + 1
            data_rows = []
            # ⚠️ 以前遇到第一列是空的行就当"这一节到此为止"，把它当成边界——
            # 用户不小心删掉小节中间一行（留了个空行）之后，这一节剩下的行
            # 会被当成下一节的内容去扫，轻则读错、重则报"缺少某列"
            # （2026-09-22 审计揪出的真 bug）。改成只认**真正的边界**：碰到
            # 下一个小节标题、或者到表尾，才算这一节结束；中间纯粹是空行的，
            # 跳过不当数据行，但不停止扫描。
            while r <= max_row:
                first_cell = ws.cell(row=r, column=1).value
                if first_cell in SECTION_TITLES:
                    break
                if first_cell not in (None, ""):
                    data_rows.append(r)
                r += 1
            yield title, header_cols, data_rows
            row = r
        else:
            row += 1


def _cell(ws: Worksheet, row: int, header_cols: dict[str, int], header: str):
    col = header_cols.get(header)
    if col is None:
        raise ConfigError(f'"{ws.title}"缺少列"{header}"')
    return ws.cell(row=row, column=col).value, row, col


def _parse(ws, row, header_cols, header, fn):
    raw, r, c = _cell(ws, row, header_cols, header)
    try:
        return fn(raw)
    except ValueError as e:
        raise ConfigError(f"{_ref(r, c)}（{header}）：{e}") from e


def _parse2(ws, row, header_cols, unit_header, value_header, fn):
    """一个值拆成"单位"+"数值"两列时用：报错定位到数值那一格。"""
    unit_raw, _, _ = _cell(ws, row, header_cols, unit_header)
    value_raw, r, c = _cell(ws, row, header_cols, value_header)
    try:
        return fn(unit_raw, value_raw)
    except ValueError as e:
        raise ConfigError(f"{_ref(r, c)}（{value_header}）：{e}") from e


def _kv_map(ws: Worksheet, header_cols: dict[str, int], data_rows: list[int]) -> dict[str, tuple]:
    """把"项目/值"这类两列小节读成 {项目文字: (值, 行号, 列号)}。"""
    item_col = header_cols.get("项目")
    value_col = header_cols.get("值")
    if item_col is None or value_col is None:
        raise ConfigError(f'"{ws.title}" 缺少"项目"或"值"列')
    result = {}
    for r in data_rows:
        key = ws.cell(row=r, column=item_col).value
        val = ws.cell(row=r, column=value_col).value
        result[str(key).strip()] = (val, r, value_col)
    return result


def _kv_parse(kv: dict, label: str, fn):
    if label not in kv:
        raise ConfigError(f'缺少"{label}"这一行')
    raw, r, c = kv[label]
    try:
        return fn(raw)
    except ValueError as e:
        raise ConfigError(f"{_ref(r, c)}（{label}）：{e}") from e


def _kv_raw(kv: dict, label: str) -> str:
    if label not in kv:
        raise ConfigError(f'缺少"{label}"这一行')
    raw, _, _ = kv[label]
    # ⚠️ 单元格被清空时 openpyxl 读出来是 Python 的 `None`，`str(None)` 会得到
    # 字面字符串 `"None"`——以前这里直接 `str(raw)`，清空的格子会被当成真的
    # 填了"None"三个字母去解析，轻则报一句看不懂的"无法识别的写法：'None'"，
    # 重则某个字段正好把 "None" 当合法值接受、静默用错数据
    # （2026-09-22 审计揪出的真 bug）。跟 `_kv_raw_optional` 一样，`None` 先
    # 变成空串，让后面的校验按"这一格是空的"正常报错。
    return "" if raw is None else str(raw).strip()


def _kv_raw_optional(kv: dict, label: str) -> str:
    """像 `_kv_raw`，但整行不存在、或值是空的，都返回空串（不报错）。

    给后加的项目用：旧版预设里没有那一行，不能因此报"缺少某行"。
    """
    if label not in kv:
        return ""
    raw, _, _ = kv[label]
    return "" if raw is None else str(raw).strip()


def _build_page_setup(ws: Worksheet, header_cols, data_rows) -> PageSetup:
    kv = _kv_map(ws, header_cols, data_rows)
    paper_choice = _kv_raw(kv, "纸张")
    if paper_choice == "自定义":
        page_width = _kv_parse(kv, "自定义宽（厘米）", units.parse_cm)
        page_height = _kv_parse(kv, "自定义高（厘米）", units.parse_cm)
    elif paper_choice in units.NAMED_PAPER_SIZES:
        page_width, page_height = units.NAMED_PAPER_SIZES[paper_choice]
    else:
        raw, r, c = kv["纸张"]
        raise ConfigError(
            f'{_ref(r, c)}（纸张）：无法识别的纸张规格："{raw}"'
            f'（应为 {"/".join(units.NAMED_PAPER_SIZES)}/自定义）'
        )
    return PageSetup(
        paper_size=paper_choice,
        page_width=page_width,
        page_height=page_height,
        orientation=_kv_raw(kv, "纸张方向"),
        margin_top=_kv_parse(kv, "上边距（厘米）", units.parse_cm),
        margin_bottom=_kv_parse(kv, "下边距（厘米）", units.parse_cm),
        margin_left=_kv_parse(kv, "左边距（厘米）", units.parse_cm),
        margin_right=_kv_parse(kv, "右边距（厘米）", units.parse_cm),
        gutter=_kv_parse(kv, "装订线（厘米）", units.parse_cm),
        header_distance=_kv_parse(kv, "页眉距边（厘米）", units.parse_cm),
        footer_distance=_kv_parse(kv, "页脚距边（厘米）", units.parse_cm),
    )


def _raw(ws: Worksheet, row: int, header_cols: dict[str, int], header: str) -> str:
    """按列名取一个**不做解析**的原始字符串。

    以前这几处直接写 `header_cols["编号"]`，缺列时抛的是裸 `KeyError: '编号'`
    ——用户看不到是哪个表哪一列缺了（2026-09-16 审计报告 Bug#4）。现在统一
    走 `_cell()`，报错里带表名和列名；顺带把空单元格的 `None` 归一成空串
    （以前 `str(None)` 会写出一个名叫 `None` 的样式）。
    """
    value = _cell(ws, row, header_cols, header)[0]
    return str(value).strip() if value is not None else ""


def _build_styles(ws: Worksheet, header_cols, data_rows) -> tuple[dict, dict]:
    by_id: dict[str, ParagraphStyleSpec] = {}
    by_name: dict[str, ParagraphStyleSpec] = {}
    for r in data_rows:
        line_rule = _parse(ws, r, header_cols, "行距类型", units.parse_line_rule)
        line_value = _parse(ws, r, header_cols, "行距值", lambda raw: units.parse_line_value(line_rule, raw))
        spec = ParagraphStyleSpec(
            id=_raw(ws, r, header_cols, "编号"),
            name=_raw(ws, r, header_cols, "名称"),
            font_east_asian=_raw(ws, r, header_cols, "中文字体"),
            font_ascii=_raw(ws, r, header_cols, "西文字体"),
            size=_parse(ws, r, header_cols, "字号", units.parse_font_size),
            bold=_parse(ws, r, header_cols, "加粗", units.parse_bold),
            alignment=_parse(ws, r, header_cols, "对齐", units.parse_alignment),
            first_line_indent=_parse2(ws, r, header_cols, "首行缩进单位", "首行缩进值", units.parse_indent_pair),
            left_indent=_parse2(ws, r, header_cols, "左缩进单位", "左缩进值", units.parse_indent_pair),
            space_before=_parse2(ws, r, header_cols, "段前单位", "段前值", units.parse_optional_length_pair),
            space_after=_parse2(ws, r, header_cols, "段后单位", "段后值", units.parse_optional_length_pair),
            line_rule=line_rule,
            line_value=line_value,
            page_break_before=_parse(ws, r, header_cols, "段前分页", units.parse_yn),
        )
        by_id[spec.id] = spec
        by_name[spec.name] = spec
    return by_id, by_name


def _build_header_footer(ws: Worksheet, header_cols, data_rows) -> list[HeaderFooterRule]:
    rules = []
    for r in data_rows:
        enabled = _parse(ws, r, header_cols, "启用页码", units.parse_yn)
        fmt = _parse(ws, r, header_cols, "页码编码", units.parse_page_number_format) if enabled else None
        rules.append(
            HeaderFooterRule(
                section_name=_raw(ws, r, header_cols, "节"),
                header_text=_raw(ws, r, header_cols, "页眉文字"),
                page_number_enabled=enabled,
                page_number_format=fmt,
                page_number_prefix=_raw(ws, r, header_cols, "前缀"),
                page_number_suffix=_raw(ws, r, header_cols, "后缀"),
            )
        )
    return rules


def _build_citation(ws: Worksheet, header_cols, data_rows) -> CitationConfig:
    kv = _kv_map(ws, header_cols, data_rows)
    return CitationConfig(
        footnote_marker_style=_kv_raw(kv, "脚注序号样式"),
        footnote_restart_each_page=_kv_parse(kv, "脚注每页重新编号", units.parse_yn),
        footnote_marker_font=_kv_raw(kv, "脚注序号字体"),
        footnote_marker_trailing_space=_kv_raw(kv, "脚注序号后空白"),
        reference_numbering_style=_kv_raw(kv, "参考文献编号样式"),
    )


def _build_structure_parts(ws: Worksheet, header_cols, data_rows
                           ) -> tuple[dict[str, bool], dict[str, str], dict[str, bool]]:
    required: dict[str, bool] = {}
    insert_text: dict[str, str] = {}
    label_pb: dict[str, bool] = {}
    pb_col = header_cols.get("标签前另起一页")      # 后加的列：旧版预设没有 → 不读
    for r in data_rows:
        name = _raw(ws, r, header_cols, "部件")
        text = _raw(ws, r, header_cols, "插入文字")
        if pb_col is not None:
            raw_pb = ws.cell(row=r, column=pb_col).value
            raw_pb = "" if raw_pb is None else str(raw_pb).strip()
            if raw_pb and raw_pb != "—":            # 空＝跟随样式；"—"＝这个部件没有这一项
                try:
                    label_pb[name] = units.parse_yn(raw_pb)
                except ValueError as e:
                    raise ConfigError(f"{_ref(r, pb_col)}（标签前另起一页）：{e}") from e
        if name in HEADING_LABEL_PARTS:
            # 引言/结论：没有"是否启用"（那一格是"—"，不读），只有写法；空＝不配。
            # 旧版预设没有这两行，缺行同样是"不配"。
            if text:
                insert_text[name] = text
            continue
        required[name] = _parse(ws, r, header_cols, "是否启用", units.parse_yn)
        insert_text[name] = text
    return required, insert_text, label_pb


def _build_cleanup_flags(ws: Worksheet, header_cols, data_rows) -> tuple[bool, bool]:
    """「清理规则」里后加的两行：清除手打分页符／清除空行（旧版预设没有＝否）。"""
    kv = _kv_map(ws, header_cols, data_rows)
    flags = []
    for label in ("清除手打分页符", "清除空行"):
        raw = _kv_raw_optional(kv, label)
        try:
            flags.append(units.parse_yn(raw) if raw else False)
        except ValueError as e:
            _val, r, c = kv[label]
            raise ConfigError(f"{_ref(r, c)}（{label}）：{e}") from e
    return flags[0], flags[1]


# 「清理规则」小节里"保留的字符格式"能认的写法 -> `w:rPr` 的子元素名。
# 认不出来的写法**直接报错**（不像别处那样退回默认值）：这是用户主动写下的
# 配置，写错了却悄默声不生效，比报错难查得多。
_KEEPABLE_RUN_FORMATS = {
    "加粗": "b",
    "斜体": "i",
    "下划线": "u",
    "上标下标": "vertAlign",
    "删除线": "strike",
    "双删除线": "dstrike",
    "颜色": "color",
    "高亮": "highlight",
    "隐藏文字": "vanish",
    "着重号": "em",
    "字符间距": "spacing",
}

# 填这几个词表示"一个都不保留"
_KEEP_NOTHING = {"无", "不保留", "全部清掉"}


def _build_keep_run_formatting(ws: Worksheet, header_cols, data_rows) -> frozenset[str]:
    """读「清理规则」小节，返回要保留的 `w:rPr` 子元素名。

    返回**空集合**是合法且有意义的：用户写了"无"，意思就是一个都不留。
    （"没配"是另一回事，由 `load_preset` 用 None 表示，会走代码里的默认值。）
    """
    kv = _kv_map(ws, header_cols, data_rows)
    raw = _kv_raw(kv, "保留的字符格式")
    if raw in _KEEP_NOTHING:
        return frozenset()

    # 分隔符随便用哪个：顿号、逗号、分号都行
    separator_normalized = raw
    for mark in ("，", ",", "；", ";"):
        separator_normalized = separator_normalized.replace(mark, "、")

    names: set[str] = set()
    for token in separator_normalized.split("、"):
        token = token.strip()
        if not token:
            continue
        if token not in _KEEPABLE_RUN_FORMATS:
            raise ConfigError(
                f'「保留的字符格式」里的"{token}"不认识。'
                f'可以写：{"、".join(_KEEPABLE_RUN_FORMATS)}；'
                f'或者写"无"表示全部清掉'
            )
        names.add(_KEEPABLE_RUN_FORMATS[token])
    return frozenset(names)


# 「页眉线」小节："页眉文字下画横线"一格——页眉文字下面那条横线要不要画。
# 官方做法是把它做在"页眉"样式的下边框上（不是逐页画形状）；填"不要"时
# 会把样式上已有的下边框摘掉（源文档自带的线也算）。认不出直接报错
# （同「清理规则」：用户主动写的配置，写错了不出声比报错难查）。
_HEADER_LINE_VALUES = {
    "要": True,
    "是": True,
    "有": True,
    "不要": False,
    "否": False,
    "无": False,
}


def _build_header_line(ws: Worksheet, header_cols, data_rows) -> tuple[bool, float, float]:
    """读「页眉线」小节，返回 `(要不要画线, 左伸出厘米, 右伸出厘米)`。

    两个"伸出"是后加的行，旧版预设里没有：缺行、留空都当 0（＝不伸出）。"""
    kv = _kv_map(ws, header_cols, data_rows)
    raw = _kv_raw(kv, "页眉文字下画横线")
    value = _HEADER_LINE_VALUES.get(raw)
    if value is None:
        raise ConfigError(
            f'「页眉文字下画横线」里填的"{raw}"不认识：'
            f'可以写"要"（页眉文字下面画横线，官方模板做法）或"不要"（不画线）'
        )
    extends = []
    for label in ("页眉线左伸出（厘米）", "页眉线右伸出（厘米）"):
        if not _kv_raw_optional(kv, label):
            extends.append(0.0)
            continue
        extends.append(_kv_parse(kv, label, lambda t: units.parse_cm(t).cm))
    return value, extends[0], extends[1]


# 「封面」小节："超长值悬挂缩进"一格——封面信息行的值太长换行时，续行对齐
# 到提示词右侧（首行不变）。默认开（用户 2026-09-12 要求"加开关、默认开启"）；
# 认不出直接报错（同其他小节）。
_HANGING_INDENT_VALUES = {
    "要": True,
    "是": True,
    "有": True,
    "不要": False,
    "否": False,
    "无": False,
}


def _build_cover(ws: Worksheet, header_cols, data_rows) -> tuple[bool, bool | None, str | None]:
    """读「封面」小节，返回 `(悬挂缩进, 要不要封面, 封面形态)`。

    「是否插入封面」/「封面形态」是 2026-09-15 补上的两格（审计核实：这两项
    在 JSON 预设里存得住，Excel 预英里却根本没有位置）。**旧预英里没有这两行**，
    所以缺行、留空都当"预设里没写"（`None`）——套用时保持界面上当前的选择不改。
    """
    kv = _kv_map(ws, header_cols, data_rows)
    raw = _kv_raw(kv, "超长值悬挂缩进")
    value = _HANGING_INDENT_VALUES.get(raw)
    if value is None:
        raise ConfigError(
            f'「超长值悬挂缩进」里填的"{raw}"不认识：'
            f'可以写"要"（续行对齐到提示词右侧）或"不要"（续行回到最左边）'
        )
    enabled_raw = _kv_raw_optional(kv, "是否插入封面")
    enabled = None
    if enabled_raw:
        enabled = _HANGING_INDENT_VALUES.get(enabled_raw)
        if enabled is None:
            raise ConfigError(
                f'「是否插入封面」里填的"{enabled_raw}"不认识：'
                f'可以写"要"（套用时把封面开关打开）/"不要"（关掉），或留空'
                f'（套用时不改这一项）'
            )
    form = _kv_raw_optional(kv, "封面形态") or None
    if form is not None and form not in COVER_FORM_NAMES:
        raise ConfigError(
            f'「封面形态」里填的"{form}"不认识：'
            f'可以写 {"/".join(COVER_FORM_NAMES)}，或留空（套用时不改这一项）'
        )
    return value, enabled, form


def load_preset(path: str | pathlib.Path) -> PresetConfig:
    # `wb.close()` 不能忘（审计 2026-09-15 核实）：openpyxl 不关的话在 Windows
    # 上会一直占着这个文件（用户"导出→又立刻导回同一个文件"时会吃权限错误），
    # 而且打开的 zip 句柄要等垃圾回收才释放。
    wb = load_workbook(path, data_only=True)
    try:
        return _load_preset_workbook(wb, path)
    finally:
        wb.close()


def _load_preset_workbook(wb, path: str | pathlib.Path) -> PresetConfig:
    if SHEET_NAME not in wb.sheetnames:
        raise ConfigError(f'找不到工作表"{SHEET_NAME}"，实际有：{", ".join(wb.sheetnames)}')
    ws = wb[SHEET_NAME]

    page = None
    styles_by_id: dict = {}
    styles_by_name: dict = {}
    header_footer: list[HeaderFooterRule] = []
    citation = None
    keep_run_formatting: frozenset[str] | None = None
    header_bottom_line = False          # 不套用预设时的默认值：不画页眉装饰线
    header_extend_left = header_extend_right = 0.0   # 不伸出
    cover_hanging_indent = True
    cover_enabled: bool | None = None   # 预设里没写（旧预设）
    cover_form: str | None = None
    structure_required: dict[str, bool] = {}
    structure_insert_text: dict[str, str] = {}
    label_page_break: dict[str, bool] = {}
    clear_page_breaks = clear_blank_lines = False

    for title, header_cols, data_rows in _iter_sections(ws):
        if title == "页面设置":
            page = _build_page_setup(ws, header_cols, data_rows)
        elif title == "样式":
            styles_by_id, styles_by_name = _build_styles(ws, header_cols, data_rows)
        elif title == "页眉页脚":
            header_footer = _build_header_footer(ws, header_cols, data_rows)
        elif title == "引注":
            citation = _build_citation(ws, header_cols, data_rows)
        elif title == "清理规则":
            # 旧版预设里没有这一节，所以只在真出现时才读，不加进下面的 missing 检查
            keep_run_formatting = _build_keep_run_formatting(ws, header_cols, data_rows)
            clear_page_breaks, clear_blank_lines = _build_cleanup_flags(ws, header_cols, data_rows)
        elif title == "页眉线":
            # 同上：新加的小节，旧版预设没有 → 保持默认值（画线）
            header_bottom_line, header_extend_left, header_extend_right = _build_header_line(
                ws, header_cols, data_rows)
        elif title == "封面":
            # 同上：新加的小节，旧版预设没有 → 保持默认值（悬挂开）
            cover_hanging_indent, cover_enabled, cover_form = _build_cover(
                ws, header_cols, data_rows)
        elif title == "结构部件":
            structure_required, structure_insert_text, label_page_break = _build_structure_parts(
                ws, header_cols, data_rows)

    # 必需小节（审计 2026-09-15 核实后收紧）：除了最早就有的「页面设置/引注/
    # 样式」，把「页眉页脚」与「结构部件」也算进来——少了它们不会报错，只会
    # 静默拿默认值（页眉文字全空、部件启用全按默认），用户看到的成品与表里
    # 写的东西对不上却没有任何提示。其余小节（清理规则/页眉线/
    # 封面）仍然可选：那几个是后加的，旧版预设里真的没有。
    missing = [n for n, v in [("页面设置", page), ("引注", citation),
                              ("页眉页脚", header_footer),
                              ("结构部件", structure_required)] if not v]
    if missing or not styles_by_id:
        missing_names = missing + (["样式"] if not styles_by_id else [])
        raise ConfigError(f'表里缺少小节：{"、".join(missing_names)}')

    return PresetConfig(
        page=page,
        styles=styles_by_id,
        styles_by_name=styles_by_name,
        header_footer=header_footer,
        citation=citation,
        structure_required=structure_required,
        structure_insert_text=structure_insert_text,
        keep_run_formatting=keep_run_formatting,
        header_bottom_line=header_bottom_line,
        header_line_extend_left_cm=header_extend_left,
        header_line_extend_right_cm=header_extend_right,
        cover_hanging_indent=cover_hanging_indent,
        cover_enabled=cover_enabled,
        cover_form=cover_form,
        clear_page_breaks=clear_page_breaks,
        clear_blank_lines=clear_blank_lines,
        label_page_break=label_page_break,
    )
