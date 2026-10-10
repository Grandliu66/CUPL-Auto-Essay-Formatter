"""命令行入口：读预设（JSON/Excel，走 preset_io）+ 用户 docx/txt，产出排好版的工作副本。

2026-09-18 起接的是新架构：识别（`recognize.recognize()`）直接产出一份
索引法的 `Entry` 列表，标签/标题的格式套用（`apply_structure.apply_from_index`）
和缺失部件的生成（`doc_index.auto_ops` + `rebuild_body`）都从这份索引读，
不再各自维护一套"识别结果"。旧的 `detect.py` 探测器（`detect_headings`/
`detect_document_title`/……）与 `apply_structure.apply_missing_parts` 那条
锚链插入的老路子已经整体退休。

2026-09-19 起 GUI 的"调整识别结果"窗口（`adjust_preview.py`）也接上了这套
架构：`build_rows()` 直接消费 `recognize()` 产出的 `list[Entry]`，导出时
`doc_index.build_ops(rows, index, ...)` 把面板当前的行序换算成对 `index`
的身份更新 + 一份指令表，`run()` 的 `index`/`ops` 两个形参就是接这个的
（见两者各自的说明）。旧的 `detect.py` 探测器与 `ManualStructure` 只剩
`essay_tool/gui` 里题注识别（`detect_captions`）、附录序号题名识别
（`detect_appendix_entries_from_index`）这两处非"部件身份"用途还在用。
"""

from __future__ import annotations

import argparse
import dataclasses
import pathlib

from docx import Document
from docx.oxml.ns import qn

from essay_tool.config.excel_io import ConfigError
from essay_tool.config.preset_io import PresetIOError, load_preset
from essay_tool.config.schema import PresetConfig
from essay_tool.engine.alerts import AlertCollector
from essay_tool.engine.apply_structure import (
    apply_appendix_entries,
    apply_captions,
    split_pictures_out_of_captions,
    apply_existing_toc_entry_styles,
    apply_from_index,
    apply_long_quotes,
    apply_table_text_styles,
    build_generated_elements,
    caption_style_indices,
    merge_keyword_paragraphs,
    normalize_image_paragraphs,
    repeat_table_header_rows,
)
from essay_tool.engine.cleanup import (
    drop_regenerated_fields,
    normalize_toc_field_range,
    strip_document_direct_formatting,
    strip_page_breaks,
    cover_top_limit,
    drop_blank_paragraph_ops,
    tables_in_prefix,
    unwrap_section_break_sdt,
)
from essay_tool.engine.captions import captions_from_index
from essay_tool.engine.detect import detect_appendix_entries_from_index
from essay_tool.engine.doc_index import (
    Entry,
    LIST_PART_KINDS,
    Op,
    auto_ops,
    body_list_field_kinds,
    body_toc_field_parts,
    has_toc_field,
    is_toc_paragraph,
    paragraph_ordinals,
    rebuild_body,
)
from essay_tool.engine.footnotes import count_real_footnotes, normalize_footnotes
from essay_tool.engine.loader import LoadError, load_as_working_copy
from essay_tool.engine.normalize import normalize_document
from essay_tool.engine.recognize import LEVEL_PART_NAMES, grammar_for, recognize
from essay_tool.engine.cover import CoverFields, add_cover
from essay_tool.engine.page_setup import apply_page_setup
from essay_tool.engine.sections import (
    apply_header_footer,
    enable_field_update_on_open,
    expand_header_tokens,
    split_body_section,
)
from essay_tool.engine.styles import StyleRegistry, refresh_toc_tab_stops

DEFAULT_PRESET = pathlib.Path("presets/中国政法大学-本科毕业论文.json")


def _toc_anchor(document):
    """目录域（或第一条目录条目）那一段——用来定位"目录落在哪一节"。"""
    for paragraph in document.paragraphs:
        if has_toc_field(paragraph):
            return paragraph
    for paragraph in document.paragraphs:
        if is_toc_paragraph(paragraph):
            return paragraph
    return None


def collect_source_alerts(
    *,
    document,
    index: list[Entry],
    structure_required: dict[str, bool],
    existing_list_kinds: set[str],
    cover_end: int | None = None,
    regenerate_fields: set[str] = frozenset(),
    collector: AlertCollector | None = None,
    include_source_checks: bool = True,
) -> tuple[AlertCollector, dict[str, set[str]], list]:
    """把所有告警规则跑一遍，返回 `(收集器, 清单插入计划, captions)`。

    **唯一一份告警清单**：`run()`（导出）和 `review()`（打开结构页时）都
    走这里——两边的结论必须一致。`captions` 一起返回：`apply_captions()`/
    `apply_appendix_entries()`/正文兜底样式那一步都要用它，省得重新识别
    一遍图题/表题。
    """
    collector = collector if collector is not None else AlertCollector()
    # 题注不再现场重新识别：识别层已经把结果写进索引（`meta["caption"]`），
    # 结构页里手动勾选的图/表/续也是写在同一个地方（`build_ops`）。
    captions = captions_from_index(index)

    if include_source_checks:
        collector.check_duplicate_parts(index)
        collector.check_duplicate_front_labels(index)
        collector.check_sequence_issues(index)
        collector.check_demoted_titles(index)
        collector.check_unaccepted_revisions(document, cover_end=cover_end)
        collector.check_front_matter_body(document, index)
        collector.check_unnumbered_objects(document, captions, cover_end=cover_end)

    caption_kinds = {c.kind for c in captions if c.wears_caption_style}
    plan = collector.check_list_parts(
        structure_required, caption_kinds, existing_list_kinds,
        regenerate_fields=set(regenerate_fields))
    collector.consolidate()
    return collector, plan, captions


def review(
    source: pathlib.Path,
    preset_config: PresetConfig,
    cover_end: int | None = None,
    heading_scheme: str | None = None,
) -> AlertCollector:
    """只做识别 + 告警，**一个文件都不生成**（打开「论文结构」页签时用）。

    `heading_scheme`：原文的标题编号方案（"chinese"/"dotted"，见
    `recognize.grammar_for`）；None 走默认的中文方案。
    """
    document = Document(source)
    unwrap_section_break_sdt(document, keep_content=True)
    index = recognize(document, grammar=grammar_for(heading_scheme), cover_end=cover_end)

    collector, _plan, _captions = collect_source_alerts(
        document=document,
        index=index,
        structure_required=preset_config.structure_required,
        existing_list_kinds=body_list_field_kinds(document),
        cover_end=cover_end,
    )
    return collector


def run(
    source: pathlib.Path,
    preset_path: pathlib.Path | None,
    work_dir: pathlib.Path,
    alerts: AlertCollector | None = None,
    cover: CoverFields | None = None,
    cover_type: str = "thesis",
    structure_override: dict[str, bool] | None = None,
    index: list[Entry] | None = None,
    ops: list[Op] | None = None,
    preset_config: PresetConfig | None = None,
    cover_end: int | None = None,
    delete_cover: bool = False,
    check_source_alerts: bool = True,
    renumber_references: bool = True,
    quote_paragraphs: list[int] = (),
    regenerate_fields: frozenset[str] = frozenset(),
    heading_scheme: str | None = None,
    session: str = "",
) -> pathlib.Path:
    """跑完整条管线，返回工作副本的路径。

    `session`：届数（如 "二十四"）。预设页眉文字里的 `{届数}` 占位换成它（每字间空一格，
    见 `sections.expand_header_tokens`）；封面上的届数走 `CoverFields.session`。

    `index`：预先识别（并可能经过人工调整）好的 `list[Entry]`——传了就直接
    用，不再自己跑一遍 `recognize()`。GUI 结构页调整过身份地图之后传这个
    （`adjust_preview.AdjustWindow.index`），先用 `doc_index.build_ops()` 把
    面板当前的行序换算成 `entry.part`/`.meta` 的更新——**这一步必须在调用
    `run()` 之前做**，`normalize_document`/`apply_from_index`/告警检查全都
    读这份 `index`，晚了就看不到用户的调整。不传（默认）就自动识别。

    `ops`：预先算好的指令表（同一次调用要连着传，通常是
    `doc_index.build_ops(rows, index, ...)` 的结果）——传了就直接拿去
    `rebuild_body`，不再自己跑 `auto_ops(index, ...)`。不传（默认）就照
    `index` 自动补全缺失部件。

    `cover_end`/`delete_cover`：封面区间与"导出时删除封面"。`cover_end`
    是排除识别/格式化的范围；`delete_cover=True`（且给了 `cover_end`）时，
    重建之后把封面那一片（含表格、含挂在封面末段上的分节符）整体删掉
    （2026-09-20 接上，此前只留接口位置）。

    `quote_paragraphs`：长引文段号（GUI 手动勾选☐引用的结果）。

    `regenerate_fields`：GUI 勾了"重新为文档生成"的目录/清单部件名——传了
    才会覆盖源文件里已有的对应域，见收尾那段 `drop_regenerated_fields`。

    `heading_scheme`：原文的标题编号方案（"chinese"/"dotted"，见
    `recognize.grammar_for`）；只在这里自己识别（没传 `index`）时用得上，
    传了 `index` 说明识别已经在调用方做完了。None 走默认的中文方案。
    """
    if preset_config is not None:
        preset = preset_config
    elif preset_path is not None:
        preset = load_preset(preset_path)
    else:
        raise ValueError("run() 需要 preset_config 或者 preset_path 至少给一个")
    if structure_override:
        # 不能原地 `.update()`：`preset` 在 `preset_config` 分支只是调用方
        # 那个对象的引用，不是副本——原地改会把这次的覆盖值永久写进调用方
        # 手上那份预设，同一个对象要是被复用着再调一次 `run()`（哪怕那次
        # 没传 `structure_override`），上次的覆盖值还留着（2026-09-22 审计
        # 揪出的：GUI 现在每次导出都现造一份新对象，暂时不会踩到，但这是
        # 潜在风险）。换成 `dataclasses.replace`，只有 `structure_required`
        # 这个字段换成新字典，不动调用方原来的对象（`PresetConfig` 不能
        # `copy.deepcopy`——会把 `Length` 这类字段的内部换算值再转一遍，
        # 2026-09-19 踩过的坑，见别处同类注释）。
        preset = dataclasses.replace(
            preset, structure_required={**preset.structure_required, **structure_override})
    working_path = load_as_working_copy(source, work_dir)

    try:
        document = Document(working_path)
    except Exception as error:  # noqa: BLE001 —— 保守起见全包住，总比甩 traceback 强
        raise LoadError(
            f"打不开这个文件：{source.name}。它可能不是真正的 .docx"
            f"（比如 .doc / .wps 改了个扩展名），请在 Word 里另存为 .docx 再试。"
        ) from error

    # 源文件里是不是已经有真正的目录域／清单域——必须在任何改写之前探明
    # （用户 2026-09-15："如果用户已有 toc，则不动"）。
    regenerate: set[str] = set(regenerate_fields)
    fields_left = body_toc_field_parts(document) - regenerate
    existing_toc = "目录" in fields_left
    keep_field_content = bool(fields_left)
    existing_list_kinds = body_list_field_kinds(document) - set()

    unwrap_section_break_sdt(document, keep_content=keep_field_content)
    # 已有目录域的标题级别范围校正到 1-3（用户 2026-09-16 裁决）。
    normalize_toc_field_range(document)

    # 识别要趁 cleanup 清掉 pPr 之前做：大纲级别、Word 自动编号（w:numPr）
    # 都挂在 pPr 上，清完就没了。
    index = index if index is not None else recognize(
        document, grammar=grammar_for(heading_scheme), cover_end=cover_end)
    ordinals = paragraph_ordinals(index)

    # "导出时删除封面"：先把封面那一片（body 开头到封面最后一段，含表格）的元素
    # 记下来，重建之后再删——删段落会让所有段落下标平移，提前删会让后面按
    # 下标做的事全部对不上号（同 `drop_regenerated_fields` 的说明）。
    # 封面在 body 顶层占到哪儿（"清除空行"等清理规则要避开封面，同样要在改动文档前算好）
    cover_limit = cover_top_limit(document, cover_end)
    cover_elements: list = []
    if delete_cover and cover_end is not None and 0 <= cover_end < len(document.paragraphs):
        # 封面 ＝ body 开头到"封面之后第一段"**之前**的全部元素（封面末段后面紧跟的
        # 表格也算，跟 `tables_in_prefix` 同一口径）；封面之后再没有段落就是到
        # body 末尾那个分节定义之前。
        following = (document.paragraphs[cover_end + 1]._p
                     if cover_end + 1 < len(document.paragraphs) else None)
        # body 自己的分节定义正常应该是最后一个直属子元素，但 WPS 存出来的
        # 文档可能把它放在别的位置（2026-09-22 审计：平时作业.docx 这份 WPS
        # 文档把它放在 body 第一个子元素）——以前"碰到任何 w:sectPr 就 break"
        # 会让整个封面一段都收不到、且没有任何提示。现在先精确定位这一个
        # 元素，扫描时只跳过它（它不是封面内容，也不能删），不再拿它当"到头
        # 了"的信号；真正的终点交给 `following`（没有 following 就收到 body
        # 末尾，只是路上跳过这一个元素）。
        body_sectPr = document.element.body.find(qn("w:sectPr"))
        for child in document.element.body:
            if child is following:
                break
            if child is body_sectPr:
                continue
            cover_elements.append(child)

    collector = alerts if alerts is not None else AlertCollector()
    collector, list_plan, captions = collect_source_alerts(
        document=document,
        index=index,
        structure_required=preset.structure_required,
        existing_list_kinds=existing_list_kinds,
        collector=collector,
        include_source_checks=check_source_alerts,
        regenerate_fields=regenerate,
        cover_end=cover_end,
    )

    # 识别是宽进的，套版不能原样照搬——统一改写成规范写法（`1、`→`1. `、
    # `引言`/`绪论`→`引　言`、`摘要`→`【摘要】` 之类）。
    normalize_document(document, index, preset.structure_insert_text)

    # 保留哪些字符格式由预设「清理规则」小节决定。正文标题例外：完全听样式。
    heading_only_indices = {i for i, e in ordinals.items()
                            if e.part in LEVEL_PART_NAMES.values()}
    cover_indices = set(range(0, cover_end + 1)) if cover_end is not None else set()
    cover_tables = (tables_in_prefix(document, cover_end + 1)
                    if cover_end is not None else set())
    # 预设「清理规则」：清除所有手打分页符（封面范围除外，用户 2026-09-20）
    if preset.clear_page_breaks:
        strip_page_breaks(document, skip_indices=cover_indices, skip_tables=cover_tables)
    # 参考文献里的链接（超链接/网址）保持原有的颜色与划线（用户 2026-09-20）
    reference_indices = {i for i, e in ordinals.items() if e.part == "参考文献"}
    strip_document_direct_formatting(
        document, preset.keep_run_formatting, style_only_indices=heading_only_indices,
        skip_indices=cover_indices, skip_tables=cover_tables,
        link_indices=reference_indices,
    )
    apply_page_setup(document, preset.page)
    built_styles = StyleRegistry(document, preset.styles)

    normal_style = built_styles["S01"]
    # 没被识别成任何部件的段落一律按正文处理（用户 2026-09-16）；题注类样式
    # 段落、目录/清单域与它的条目除外。
    keep_body_style = caption_style_indices(document) | {
        c.paragraph_index for c in captions}
    for i, paragraph in enumerate(document.paragraphs):
        if i in cover_indices or i in keep_body_style:
            continue
        if has_toc_field(paragraph) or is_toc_paragraph(paragraph):
            continue
        paragraph.style = normal_style

    normalize_image_paragraphs(document, skip_indices=cover_indices)

    # 图题/表题不能吃"内容样式"——直接格式会盖掉图题/表题样式自己的居中、
    # 行距。识别不到的题注类段落也要保护（样式是 题注/Caption/图题/表题）。
    protected_indices = ({c.paragraph_index for c in captions}
                         | caption_style_indices(document))
    # 指令表要在套格式之前定下来：关键词"标签一段、关键词一段"的老写法在这里并成
    # 一段（去掉被并走那一段的"照抄"指令），套格式才能按同段写法给两截文字分别上字体。
    if ops is None:
        ops = auto_ops(
            index, preset.structure_required,
            list_kinds=list_plan, insert_toc=not existing_toc,
        )
    ops = merge_keyword_paragraphs(document, index, ops)
    apply_from_index(
        document, index, built_styles, skip_indices=protected_indices,
        renumber_references=renumber_references,
        reference_style=preset.citation.reference_numbering_style,
        label_page_break=preset.label_page_break,
    )
    apply_appendix_entries(
        document, detect_appendix_entries_from_index(document, index), built_styles
    )
    apply_table_text_styles(document, built_styles, skip_tables=cover_tables)
    repeat_table_header_rows(document, skip_tables=cover_tables)
    apply_captions(document, captions, built_styles)
    if quote_paragraphs:
        apply_long_quotes(document, list(quote_paragraphs), built_styles)

    if count_real_footnotes(document):
        body_spec = preset.styles.get("S01")
        body_fonts = (
            (body_spec.font_east_asian, body_spec.font_ascii)
            if body_spec is not None
            else None
        )
        normalize_footnotes(document, built_styles["S23"].style_id, preset.citation,
                            body_fonts=body_fonts)

    # ---- 缺失部件的自动补充：索引法执行器，取代旧的锚链插入 ----
    # 到这一步为止，`document` 上所有"已存在"的段落都已经清理/套好格式；
    # `auto_ops` 只决定"启用但缺失的部件该插在哪条指令的位置"（指令表在上面
    # `apply_from_index` 之前就算好了），`rebuild_body` 按这份指令表把正文整体
    # 重排一遍——已存在的段落原样搬走（元素没被复制，携带着刚套好的格式），
    # 缺失的部件现造现插。

    def builder(op):
        return build_generated_elements(
            document, op, preset.structure_insert_text, built_styles,
            list_kinds=frozenset(),
            renumber_references=renumber_references,
            reference_style=preset.citation.reference_numbering_style,
            label_page_break=preset.label_page_break,
        )

    # 预设「清理规则」：清除空行（确实为空的段落；封面范围除外，用户 2026-09-20）。
    # 在指令表里去掉这些段落的"照抄"指令——重建时它们就不会被搬到成品里，
    # 识别索引、结构页/告警用的原文段号都不受影响。
    if preset.clear_blank_lines:
        ops, _n_blank = drop_blank_paragraph_ops(
            ops, document, cover_limit, page_breaks_removed=preset.clear_page_breaks)

    track: dict[int, int] = {}
    new_field_elements: list = []
    rebuild_body(document, ops, builder=builder, track=track,
                 new_field_elements=new_field_elements)

    # 挂了图题/表题样式的段落里夹着的嵌入图，要挪到前面单独一段——清单域按
    # 样式收整段，不挪的话图片会被一起搬进清单条目。这里先按重建后的序号取
    # 出段落对象（对象在插段后仍然有效，序号不行），真正拆分放在下面分节之后。
    rebuilt_paragraphs = document.paragraphs
    picture_caption_paragraphs = [
        rebuilt_paragraphs[track[ordinals[c.paragraph_index].top_index]]
        for c in captions
        if c.wears_caption_style
        and ordinals[c.paragraph_index].top_index in track
        and track[ordinals[c.paragraph_index].top_index] < len(rebuilt_paragraphs)
    ]

    if cover_elements:
        # 封面元素（含分节符所在段）整体删除——封面本来就被排除在识别之外，
        # 删掉不影响任何身份；后面的段落序号整体前移，`track` 跟着换算。
        removed_paragraphs = sum(1 for el in cover_elements
                                 if el.tag == qn("w:p") and el.getparent() is not None)
        for el in cover_elements:
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)
        track = {k: v - removed_paragraphs for k, v in track.items()
                 if v >= removed_paragraphs}

    # 分节点：正常是"引言"那一段——重建之后原来的段落下标全部作废，靠
    # `track`（源 top_index → 重建后段落序号）换算。引言缺失时退回"正文
    # 第一个已识别标题"（重建之后同样要用 track 换算）。
    intro_entry = next((e for e in ordinals.values() if e.part == "引言"), None)
    body_start_entry = min(
        (e for e in ordinals.values() if e.part in LEVEL_PART_NAMES.values()),
        key=lambda e: e.top_index, default=None,
    )
    body_start = None
    for candidate in (intro_entry, body_start_entry):
        if candidate is not None and candidate.top_index in track:
            body_start = track[candidate.top_index]
            break
    body_anchor = (document.paragraphs[body_start]
                   if body_start is not None
                   and 0 <= body_start < len(document.paragraphs) else None)
    body_section_index = split_body_section(document, body_anchor)
    split_pictures_out_of_captions(picture_caption_paragraphs, built_styles)
    header_rules = [dataclasses.replace(r, header_text=expand_header_tokens(r.header_text, session))
                    for r in preset.header_footer]
    apply_header_footer(
        document, header_rules, built_styles, preset.header_bottom_line,
        (preset.header_line_extend_left_cm, preset.header_line_extend_right_cm),
        body_section_index=body_section_index)
    refresh_toc_tab_stops(document, _toc_anchor(document))
    enable_field_update_on_open(document)

    if cover is not None:
        add_cover(document, cover_type, cover, hanging_indent=preset.cover_hanging_indent)

    # 用户勾了"重新为文档生成"的目录/清单（`regenerate_fields` 形参）。
    if regenerate:
        replaced: set[str] = set()
        if "目录" in regenerate and not existing_toc:
            replaced.add("目录")
        for part, kinds in list_plan.items():
            if part not in regenerate:
                continue
            replaced.add(part)
            replaced |= {name for name, need in LIST_PART_KINDS.items()
                         if need <= set(kinds)}
        if replaced:
            # `keep`＝这次刚生成的新域（`new_field_elements`）：不能不传——不传
            # 的话新旧域按同一套"像不像目录域"的判据一起被扫中，新插的域也会被
            # 删掉，成品变成"一条域都没有"（2026-09-22 审计揪出的真 bug，实测
            # 复现过：目录域 1 个→0 个）。
            drop_regenerated_fields(document, replaced, keep=new_field_elements)

    # 已有目录域（没插新的）时，条目样式仍要按预设建、并真的挂到条目上。
    if existing_toc:
        for style_id in ("S15", "S16", "S17"):
            built_styles.ensure(style_id)
        apply_existing_toc_entry_styles(document, built_styles)

    document.save(working_path)
    if alerts is None:
        collector.print_report()
    return working_path


def main():
    parser = argparse.ArgumentParser(description="按预设给论文排版")
    parser.add_argument("source", type=pathlib.Path, help="用户的 .docx 或 .txt")
    parser.add_argument("--preset", type=pathlib.Path, default=DEFAULT_PRESET)
    parser.add_argument("--work-dir", type=pathlib.Path, default=pathlib.Path("work"))
    args = parser.parse_args()

    alerts = AlertCollector()
    try:
        out = run(args.source, args.preset, args.work_dir, alerts=alerts)
    except (ConfigError, PresetIOError, LoadError) as e:
        print(f"出错：{e}")
        raise SystemExit(1)
    print(f"已生成：{out}")
    alerts.print_report()


if __name__ == "__main__":
    main()
