r"""单因素扫描：**逐个翻一个开关，跑一遍，比产物差异。**

这一块直接回答"这个按钮到底有没有用"：
  * 每次只改一个变量，其余全不动；
  * 跑基线一次（复用），然后每个开关各跑一次；
  * 差异比对（`aat_diff`）告诉你"变了什么"。

三种结论，自动分好：
  1. **完全没变化** → 这个开关没用，或者这份夹具触发不到它（后者很常见，见 `hint`）
  2. 有变化、且都在声明的范围内 → 有效
  3. **越界** → 还动了不该动的地方（误伤），比"没用"严重

范围声明（`Toggle.scope`）是**粗粒度**的：只声明"允许变化出现在哪几个桶"
（节 / 样式 / 正文 / 表格 / 页眉 / 页脚 / 脚注）。这是故意的——精细声明要等
「答案生成器」上线才做得了，粗粒度已经能抓住大部分误伤。

只读生产代码，产物写在 AAT 文件夹自己的 `_out/` 下。
"""

from __future__ import annotations

import copy
import dataclasses
import pathlib
import shutil
import sys
import time
from dataclasses import dataclass
from typing import Callable

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent
for entry in (str(ROOT), str(HERE)):
    if entry not in sys.path:
        sys.path.insert(0, entry)

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from docx.shared import Emu, Pt                       # noqa: E402

import aat_diff                                       # noqa: E402
from aat_docx import read_docx                        # noqa: E402
from essay_tool.config.preset_io import load_preset    # noqa: E402
from essay_tool.config.units import (ALIGNMENTS, LINE_RULES,  # noqa: E402
                                     parse_indent_pair)
from essay_tool.engine import pipeline                 # noqa: E402

ONE_CM = Emu(360000)
PAGE_NUMBER_FORMATS = ("decimal", "upperRoman", "lowerRoman", "upperLetter",
                       "lowerLetter")

# 会影响**版心宽度**的开关，都会连长目录条目的右对齐制表位——因为
# `styles.refresh_toc_tab_stops` 是按目录所在节的版心宽算的。这是引擎的
# **有意耦合**（2026-09-17 加的），不是误伤。
WIDTH_COUPLING = {
    "样式": "版心宽度变了 → 目录条目的制表位跟着变（有意耦合）",
    "正文": "目录条目段落上的制表位跟着变（有意耦合）",
}


@dataclass
class Toggle:
    """一个"只动一处"的实验。"""

    group: str
    name: str
    scope: set[str]                # 允许变化的桶
    mutate: Callable[[object, dict], str]     # (preset_config, extras) -> 说明
    expect_styles: set[str] | None = None     # 只允许变这几个样式（None=不检查）
    needs_cover: bool = False                 # 要带自动封面跑
    hint: str = ""                            # "没变化"时给出的提示
    # `expect_keys`：格式开关专用——**只允许这几个属性变**。
    # 为什么不用“只允许这个样式变”：预设里的角色编号与 Word 样式不是一套
    # （角色「正文」落到内置 `Normal`，`目录一级` 落到 `toc 1`，`脚注文本` 落到
    # `footnote text`，`页眉` 落到 `header`…）。按名字比会在这些角色上**全报假越界**
    # （实测）。改成按属性比：改了字号，那么"任何样式的字号跟着变"都合理，
    # 而"某个样式的对齐跟着变了"才可疑。
    expect_keys: set[str] | None = None
    expect_styles: set[str] | None = None     # 备而不用（按名字比太脆）
    # `known`：**已知耦合**——变了、但不算缺陷的桶（值写理由）。
    # 例：改页边距 → 版心宽度变 → 目录条目的制表位跟着变，这是引擎的有意设计
    # （`styles.refresh_toc_tab_stops`）。不这么登记的话，每个"有意耦合"都会
    # 被报成误伤，报告很快就没人看了。
    known: dict[str, str] = dataclasses.field(default_factory=dict)

    def apply(self, config, extras) -> str:
        return self.mutate(config, extras)


# ------------------------------------------------------------------ 取值翻转


def _different(current, candidates):
    """从候选里挑一个跟当前不同的值。"""
    for candidate in candidates:
        if candidate != current:
            return candidate
    return current


def _cm(value) -> str:
    return f"{Emu(int(value)).cm:.2f}厘米"


def _text(value) -> str:
    return "（空）" if value in (None, "") else repr(value)


def safe_rel(path) -> str:
    """能相对就相对、不能就绝对——命令行常给相对路径，硬转会 ValueError。"""
    resolved = pathlib.Path(path).resolve()
    try:
        return str(resolved.relative_to(ROOT))
    except ValueError:
        return str(resolved)


# ------------------------------------------------------------------ 开关清单


def page_toggles(config, *, source=None) -> list[Toggle]:
    out: list[Toggle] = []
    for field, label in (("margin_left", "左页边距"), ("margin_right", "右页边距"),
                         ("margin_top", "上页边距"), ("margin_bottom", "下页边距"),
                         ("gutter", "装订线"), ("header_distance", "页眉距边"),
                         ("footer_distance", "页脚距边")):
        def mutate(cfg, _extras, field=field, label=label):
            before = getattr(cfg.page, field)
            after = Emu(int(before) + int(ONE_CM))
            setattr(cfg.page, field, after)
            return f"{label} {_cm(before)} → {_cm(after)}"
        # 只有左右页边距会改变版心宽度，才会连长目录制表位
        changes_width = field in ("margin_left", "margin_right")
        out.append(Toggle("页面", f"页边距·{label}", {"节"}, mutate,
                          known=dict(WIDTH_COUPLING) if changes_width else {}))

    def flip_line(cfg, _extras):
        before = bool(cfg.header_bottom_line)
        cfg.header_bottom_line = not before
        return f"页眉装饰线 {before} → {not before}"
    out.append(Toggle("页面", "页眉装饰线", {"样式", "页眉"}, flip_line,
                      hint="页眉线做在「页眉」样式的下边框上，要看 styles.xml"))

    # —— 2026-09-19 对账器查出来的漏项：页眉线的左右伸缩量 ——
    for field, label in (("header_line_extend_left_cm", "页眉线左伸"),
                         ("header_line_extend_right_cm", "页眉线右伸")):
        def extend(cfg, _extras, field=field, label=label):
            before = float(getattr(cfg, field) or 0.0)
            after = 0.0 if before else 0.5
            setattr(cfg, field, after)
            return f"{label} {before} → {after} 厘米"
        out.append(Toggle("页面", f"页眉线·{label}", {"样式", "页眉"}, extend,
                          hint="预设里没开「页眉装饰线」时这项不生效"
                               "（这份预设默认关着）"))

    for index, rule in enumerate(getattr(config, "header_footer", []) or []):
        section = rule.section_name or f"第{index}节"

        def text_mutate(cfg, _extras, index=index, section=section):
            target = cfg.header_footer[index]
            before = target.header_text
            target.header_text = "测试页眉文字"
            return f"{section} 页眉文字 {before!r} → '测试页眉文字'"
        out.append(Toggle("页面", f"页眉文字·{section}", {"页眉", "节"}, text_mutate))

        def page_mutate(cfg, _extras, index=index, section=section):
            target = cfg.header_footer[index]
            before = bool(target.page_number_enabled)
            target.page_number_enabled = not before
            if not target.page_number_enabled:
                target.page_number_format = None
            else:
                target.page_number_format = target.page_number_format or "decimal"
            return f"{section} 页码 {before} → {not before}"
        out.append(Toggle("页面", f"页码开关·{section}", {"页脚", "节"}, page_mutate))

        def format_mutate(cfg, _extras, index=index, section=section):
            target = cfg.header_footer[index]
            if not target.page_number_enabled:
                return "（页码未启用，跳过）"
            before = target.page_number_format or "decimal"
            target.page_number_format = _different(before, PAGE_NUMBER_FORMATS)
            return f"{section} 页码格式 {before} → {target.page_number_format}"
        out.append(Toggle("页面", f"页码格式·{section}", {"页脚", "节"}, format_mutate,
                          hint="页码没启用时这项不生效"))

        def prefix_mutate(cfg, _extras, index=index, section=section):
            target = cfg.header_footer[index]
            before = target.page_number_prefix
            target.page_number_prefix = "第"
            return f"{section} 页码前缀 {before!r} → '第'"
        out.append(Toggle("页面", f"页码前缀·{section}", {"页脚", "节"}, prefix_mutate,
                          hint="页码没启用时这项不生效"))

        def suffix_mutate(cfg, _extras, index=index, section=section):
            target = cfg.header_footer[index]
            before = target.page_number_suffix
            target.page_number_suffix = "页"
            return f"{section} 页码后缀 {before!r} → '页'"
        out.append(Toggle("页面", f"页码后缀·{section}", {"页脚", "节"}, suffix_mutate,
                          hint="页码没启用时这项不生效"))

    # —— 纸张与方向（对账器 `aat_coverage.py` 查出来的漏项）——
    def size_mutate(cfg, _extras, field="page_width", label="页宽"):
        before = getattr(cfg.page, field)
        after = Emu(int(before) + 2 * int(ONE_CM))
        setattr(cfg.page, field, after)
        return f"{label} {_cm(before)} → {_cm(after)}"
    out.append(Toggle("页面", "页宽", {"节"},
                      lambda c, e: size_mutate(c, e, "page_width", "页宽"),
                      known=dict(WIDTH_COUPLING)))
    out.append(Toggle("页面", "页高", {"节"},
                      lambda c, e: size_mutate(c, e, "page_height", "页高")))

    def orientation_mutate(cfg, _extras):
        before = cfg.page.orientation
        cfg.page.orientation = "横向" if before != "横向" else "纵向"
        return f"方向 {before!r} → {cfg.page.orientation!r}"
    out.append(Toggle("页面", "纸张方向", {"节"}, orientation_mutate,
                      known=dict(WIDTH_COUPLING)))

    def paper_mutate(cfg, _extras):
        before = cfg.page.paper_size
        cfg.page.paper_size = "A3" if before != "A3" else "A4"
        return f"纸张名 {before!r} → {cfg.page.paper_size!r}"
    out.append(Toggle("页面", "纸张名", {"节"}, paper_mutate,
                      hint="这一项只是预设里的显示名，实际尺寸在「页宽/页高」那两项"))
    return out


FORMAT_FIELDS: tuple[tuple[str, str, tuple], ...] = (
    ("字体（中文）", "font_east_asian", ("楷体", "黑体", "宋体", "仿宋")),
    ("字体（西文）", "font_ascii", ("Times New Roman", "Arial", "Calibri")),
    ("字号", "size", (Pt(12), Pt(10.5), Pt(14), Pt(16), Pt(9))),
    ("加粗", "bold", (True, False)),
    ("对齐", "alignment", tuple(ALIGNMENTS.values())),
    ("首行缩进", "first_line_indent", (parse_indent_pair("字符", 2),
                                   parse_indent_pair("字符", 0),
                                   parse_indent_pair("厘米", 0.85))),
    ("左缩进", "left_indent", (parse_indent_pair("字符", 2),
                            parse_indent_pair("字符", 0))),
    ("段前", "space_before", (Pt(0), Pt(6), Pt(12))),
    ("段后", "space_after", (Pt(0), Pt(6), Pt(12))),
    ("行距类型", "line_rule", tuple(LINE_RULES.values())),
    ("段前分页", "page_break_before", (True, False)),
    # 行距值的含义取决于行距类型（倍数 or 磅），候选值在下面按当前类型现算
    ("行距值", "line_value", ()),
)

# 预设里的字段名 → 它落在样式定义里的哪个属性（用于判断"该不该动"）
FIELD_TO_STYLE_KEYS: dict[str, set[str]] = {
    "font_east_asian": {"rPr·rFonts"},
    "font_ascii": {"rPr·rFonts"},
    "size": {"rPr·sz", "rPr·szCs"},
    "bold": {"rPr·b"},
    "alignment": {"pPr·jc"},
    "first_line_indent": {"pPr·ind"},
    "left_indent": {"pPr·ind"},
    "space_before": {"pPr·spacing"},
    "space_after": {"pPr·spacing"},
    "line_rule": {"pPr·spacing"},
    "line_value": {"pPr·spacing"},
    "page_break_before": {"pPr·pageBreakBefore"},
}


def format_toggles(config, *, styles=None, fields=None, source=None) -> list[Toggle]:
    """格式页：每个样式 × 每个取值，各生成一个开关。

    这是"人力做不到"的那部分——三十来个样式 × 十一个取值，几百个格子。
    """
    out: list[Toggle] = []
    for style_id in sorted(config.styles):
        if styles and style_id not in styles:
            continue
        spec = config.styles[style_id]
        for label, field, candidates in FORMAT_FIELDS:
            if fields and field not in fields:
                continue
            if not hasattr(spec, field):
                continue
            current = getattr(spec, field)
            if not candidates and field == "line_value":
                # 多倍行距时值是个倍数（float），固定值/最小值时是个磅值
                candidates = ((1.0, 1.5, 2.0, 2.5)
                              if spec.line_rule is not None
                              and "MULTIPLE" in str(spec.line_rule)
                              else (Pt(20), Pt(24), Pt(18), Pt(28)))
            new_value = _different(current, candidates)
            if new_value == current:
                continue
            allowed = FIELD_TO_STYLE_KEYS.get(field)

            def mutate(cfg, _extras, style_id=style_id, field=field,
                       new_value=new_value, label=label):
                target = cfg.styles[style_id]
                before = getattr(target, field)
                setattr(target, field, new_value)
                return (f"{style_id}·{target.name} · {label} "
                        f"{before!r} → {new_value!r}")
            out.append(Toggle("格式", f"{style_id}·{spec.name}·{label}",
                              {"样式", "正文", "表格", "页眉", "页脚", "脚注", "节"},
                              mutate, expect_keys=allowed,
                              hint="这份夹具里没用到这个样式时，改它不会有任何变化"
                                   "（引擎按需建样式：没用到就不写进文档）"))
    return out


def _pick_quote_index(source: pathlib.Path | None):
    """给「长引文手选」挑一个段号（源文档段落下标，与 GUI 传的 `doc_index` 同义）。

    挑法：跳过开头 8 段（题目/封面区），要第一个长度 ≥30 字的段落；挑不到
    返回 None（那条记「本夹具不适用」）。覆盖率对账没有夹具（source=None），
    给个占位值——那边只关心 extras 里有没有这个键。
    """
    if source is None:
        return 4
    from docx import Document as _Document
    document = _Document(str(source))
    for index, paragraph in enumerate(document.paragraphs):
        if index < 8:
            continue
        if len(paragraph.text.strip()) >= 30:
            return index
    return None


def _direct_index(source: pathlib.Path | None):
    """复刻 `pipeline.run()` 自识别那条路，造一份**与之相同**的 `index`。

    run() 的顺序（见其源码）：先按“有没有域”定 keep_content 拆内容控件 →
    `normalize_toc_field_range` → `recognize(grammar, cover_end)`。这里照抄
    同样的顺序，直传回去应当**逐字节不变**——这条开关的预期结论就是“没变化”。
    对账（source=None）时给空列表占位，只为把 `param.index` 这个键设上。
    """
    if source is None:
        return []
    from docx import Document as _Document
    from essay_tool.engine.cleanup import (normalize_toc_field_range,
                                           unwrap_section_break_sdt)
    from essay_tool.engine.doc_index import body_toc_field_parts
    from essay_tool.engine.recognize import recognize
    document = _Document(str(source))
    keep_content = bool(body_toc_field_parts(document))
    unwrap_section_break_sdt(document, keep_content=keep_content)
    normalize_toc_field_range(document)
    return recognize(document)


def structure_toggles(config, *, source=None) -> list[Toggle]:
    out: list[Toggle] = []
    for part in sorted(config.structure_required):
        def mutate(cfg, _extras, part=part):
            before = bool(cfg.structure_required[part])
            cfg.structure_required[part] = not before
            return f"部件「{part}」启用 {before} → {not before}"
        # 启用一个部件会把它自己的样式建出来（引擎按需建样式），所以"样式"也在范围内。
        out.append(Toggle("结构", f"部件·{part}", {"正文", "表格", "样式"}, mutate,
                          hint=f"这份夹具里可能本来就没有「{part}」，"
                               "或它已经存在于原文、开关只决定「缺了要不要补」"))

    # —— 2026-09-19 接新版 `pipeline.run()` 形参：老 ManualStructure 那条路
    #    下线了，结构页那几个开关现在直接改 run() 的参数（对账器的漏项在这里闭合）——
    def renumber(_cfg, extras):
        before = bool(extras.get("renumber_references", True))
        extras["renumber_references"] = not before
        return f"参考文献自动重新编号 {before} → {not before}"
    out.append(Toggle("结构", "参考文献重编号", {"正文", "样式"}, renumber,
                      hint="夹具里没有手打编号的参考文献时看不出来"))

    def regenerate(_cfg, extras):
        before = set(extras.get("regenerate_fields") or ())
        extras["regenerate_fields"] = frozenset() if before else frozenset({"目录"})
        return f"「重新为文档生成」 {sorted(before) or '（空）'} → " \
               f"{sorted(extras['regenerate_fields']) or '（空）'}"
    out.append(Toggle("结构", "目录重新生成", {"正文"}, regenerate,
                      hint="这份夹具的源文件里可能没有真目录域"))

    def scheme(_cfg, extras):
        before = extras.get("heading_scheme") or "chinese"
        extras["heading_scheme"] = "chinese" if before == "dotted" else "dotted"
        return f"标题编号方案 {before} → {extras['heading_scheme']}"
    out.append(Toggle("结构", "标题编号方案", {"正文", "样式"}, scheme,
                      hint="两种方案认出来的标题集合一样时看不出来"
                           "（点分方案只管『1.2.3.』这类编号的识别）",
                      known={
                          "节": "编号方案改变「哪些段是标题」⇒ 正文分节点/节数跟着变"
                                "（2026-09-19 在 PaperStructureTest.docx 实测："
                                "首段带「一、」时两种方案节数不同）",
                          "页眉": "同上：节数变了，节的页眉引用跟着换",
                          "页脚": "同上：节数变了，节的页脚引用跟着换",
                      }))

    def quote(_cfg, extras):
        target = _pick_quote_index(source)
        if target is None:
            return None
        extras["quote_paragraphs"] = ([] if extras.get("quote_paragraphs")
                                      else [target])
        return f"长引文手选 → {extras['quote_paragraphs']}"
    out.append(Toggle("结构", "长引文手选", {"正文", "样式"}, quote,
                      hint="挑不到 ≥30 字的正文段时这条记「不适用」"))

    def direct(_cfg, extras):
        extras["index"] = _direct_index(source)
        extras["ops"] = None          # 显式传 None＝走 run 自己的 auto_ops 分支
        return "index/ops 直传（与自动路径同值，预期逐字节不变）"
    out.append(Toggle("结构", "index/ops 直传", set(), direct,
                      hint="这条的预期结论就是「没变化」——直传与自动路径等价"))

    # ⚠️ 用户 2026-09-17 裁决：下面两项**不给用户调**，所以不做开关：
    #   * `keep_run_formatting`（保留字符格式）——纯粹的代码层面规则，
    #     保留什么写死在《工具操作逻辑》3.6 表 R18 的「样式应用规则」列里
    #     （保留加粗/斜体/下划线/上下标 + 段落样式与自动编号）。
    #   * `structure_insert_text`（缺失时插入的文字）——同样不对外调整。
    # 若将来它们又被改成可调的，记得在这里补回开关。

    def override(cfg, extras):
        before = extras.get("structure_override")
        extras["structure_override"] = {"附录": True}
        return f"强推部件开关 structure_override {_text(before)} → {{'附录': True}}"
    out.append(Toggle("结构", "强推启用全部部件", {"正文", "表格", "样式"}, override,
                      hint="这份夹具没有缺部件时，强推也不会多出东西"))
    return out


def cover_toggles(config, *, source=None) -> list[Toggle]:
    out: list[Toggle] = []

    def enabled(cfg, _extras):
        before = cfg.cover_enabled
        cfg.cover_enabled = not (before is None or before)
        return f"封面总开关 {before} → {cfg.cover_enabled}"
    out.append(Toggle("封面", "封面总开关", {"正文", "节", "页眉", "页脚"}, enabled,
                      needs_cover=True))

    def form(cfg, _extras):
        before = cfg.cover_form
        cfg.cover_form = ("课程考核封面" if before != "课程考核封面"
                          else "毕业论文封面")
        return f"封面形态 {before!r} → {cfg.cover_form!r}"
    out.append(Toggle("封面", "封面形态", {"正文", "节", "页眉", "页脚"}, form,
                      needs_cover=True))

    def hanging(cfg, _extras):
        before = bool(cfg.cover_hanging_indent)
        cfg.cover_hanging_indent = not before
        return f"超长值悬挂缩进 {before} → {not before}"
    out.append(Toggle("封面", "悬挂缩进", {"正文"}, hanging, needs_cover=True,
                      hint="封面里没有超长信息时，勾不勾都一样"))

    def cover_range(cfg, extras):
        before = extras.get("cover_end")
        extras["cover_end"] = 0 if before is None else None
        return f"封面范围 cover_end {before!r} → {extras['cover_end']!r}"
    out.append(Toggle("封面", "封面范围", {"正文"}, cover_range,
                      hint="源文件没有封面时，设不设前三段为封面都看不出差别"))

    def delete_cover(cfg, extras):
        before = bool(extras.get("delete_cover"))
        extras["delete_cover"] = not before
        if extras.get("cover_end") is None:
            extras["cover_end"] = 0
        return f"导出时删除封面 {before} → {not before}"
    out.append(Toggle("封面", "删除封面", {"正文", "表格"}, delete_cover,
                      hint="要先有封面范围，且源文件前几段得真有内容"))

    # —— 对账器 `aat_coverage.py` 查出来的漏项 ——
    def cover_kind(cfg, extras):
        before = extras.get("cover_type", "thesis")
        extras["cover_type"] = "course" if before != "course" else "thesis"
        return f"封面形态(参数 cover_type) {before!r} → {extras['cover_type']!r}"
    out.append(Toggle("封面", "封面形态·参数", {"正文", "节", "页眉", "页脚"},
                      cover_kind, needs_cover=True,
                      hint="注意：封面形态走的是 pipeline 参数 cover_type，"
                           "不是预设里的 cover_form"))

    from essay_tool.engine.cover import CoverFields as _CoverFields
    for spec in dataclasses.fields(_CoverFields):
        def fill(cfg, extras, name=spec.name):
            cover = extras.get("cover") or _CoverFields()
            before = getattr(cover, name)
            setattr(cover, name, "测试内容")
            extras["cover"] = cover
            return f"封面填写·{name} {_text(before)} → '测试内容'"
        out.append(Toggle("封面", f"填写·{spec.name}", {"正文"}, fill,
                          needs_cover=True,
                          hint="这一栏留空时封面就是空位留白，填了才看得出来"))
    return out


def citation_toggles(config, *, source=None) -> list[Toggle]:
    out: list[Toggle] = []
    if getattr(config, "citation", None) is None:
        return out

    fields = (
        ("脚注序号后空白", "footnote_marker_trailing_space", ("半角空格", "全角空格",
                                                          "不空格")),
        ("脚注序号字体", "footnote_marker_font", ("楷体", "黑体", "正文(非上标)")),
        ("脚注序号样式", "footnote_marker_style", ("①②③", "1.2.3.")),
        ("参考文献编号写法", "reference_numbering_style", ("[1][2][3]", "1.2.3.",
                                                     "(1)(2)(3)")),
    )
    for label, field, candidates in fields:
        if not hasattr(config.citation, field):
            continue
        current = getattr(config.citation, field)
        new_value = _different(current, candidates)

        def mutate(cfg, _extras, field=field, new_value=new_value, label=label):
            target = cfg.citation
            before = getattr(target, field)
            setattr(target, field, new_value)
            return f"{label} {before!r} → {new_value!r}"
        out.append(Toggle("引注", label, {"脚注", "正文", "样式"}, mutate,
                          hint="这份夹具里没有脚注/参考文献，这项就看不出来"))

    def restart(cfg, _extras):
        before = bool(cfg.citation.footnote_restart_each_page)
        cfg.citation.footnote_restart_each_page = not before
        return f"每页重新编号 {before} → {not before}"
    out.append(Toggle("引注", "脚注每页重编号", {"脚注"}, restart,
                      hint="这份夹具里没有脚注时看不出来"))
    return out


GROUPS = {
    "页面": page_toggles,
    "格式": format_toggles,
    "结构": structure_toggles,
    "封面": cover_toggles,
    "引注": citation_toggles,
}


# ------------------------------------------------------------------ 跑批


# （原 `build_manual_structure` 已删：2026-09-19 管线接口从 `manual_structure`
#   换成 `index/ops` 后它没有出口了——结构页那几个开关改走 `pipeline.run()`
#   的形参，见 `structure_toggles`。）


@dataclass
class CaseOutcome:
    toggle: Toggle
    description: str
    result: aat_diff.DiffResult | None
    seconds: float
    skipped: bool = False
    skip_reason: str = ""

    def buckets(self) -> dict[str, int]:
        if self.result is None:
            return {}
        counts = self.result.counts
        return {
            "节": counts.get("节", 0),
            "样式": counts.get("样式", 0),
            "正文": counts.get("正文段", 0),
            "表格": counts.get("表段", 0),
            "页眉": counts.get("页眉段", 0),
            "页脚": counts.get("页脚段", 0),
            "脚注": counts.get("脚注段", 0),
        }

    def verdict(self) -> tuple[str, list[str], list[str]]:
        """返回 `(结论, 越界说明, 已知耦合说明)`。"""
        if self.skipped:
            return "skip", [self.skip_reason], []
        changed = {name: value for name, value in self.buckets().items() if value}
        if not changed:
            return "no-effect", [], []
        escaped: list[str] = []
        coupled: list[str] = []
        for name in changed:
            if name in self.toggle.scope:
                continue
            if name in self.toggle.known:
                coupled.append(f"{name}：{self.toggle.known[name]}")
                continue
            escaped.append(
                f"{name} 变了，但它不在声明范围（{'/'.join(sorted(self.toggle.scope))}）里")
        if self.toggle.expect_keys is not None and self.result is not None:
            for item in self.result.style_changes:
                for key in item.changes:
                    if key in ("名称", "basedOn"):
                        continue
                    if key not in self.toggle.expect_keys:
                        escaped.append(
                            f"顺带动了不该动的属性：{item.where} → {key}"
                            f"（本次只该动 {'/'.join(sorted(self.toggle.expect_keys))}）")
        if self.toggle.expect_styles is not None and self.result is not None:
            for item in self.result.style_changes:
                _, _, style_name = item.where.partition("/")
                if (style_name or item.where.split()[-1]) not in self.toggle.expect_styles:
                    escaped.append(f"顺带动了别的样式：{item.where}")
        return ("escaped" if escaped else "in-scope"), escaped, coupled


def run_case(source: pathlib.Path, preset_path: pathlib.Path, out_dir: pathlib.Path,
             toggle: Toggle, *, needs_cover: bool = False):
    """跑一个实验：重新读一份预设 → 只改一处 → 跑管线 → 读产物。

    ⚠️ **绝不能 `copy.deepcopy(PresetConfig)`**：预设里的页边距是 `docx.shared.Cm`
    这类 `Length` 子类（int 子类，`__new__` 里做了“厘米→EMU”换算），deepcopy 时
    会拿已经换算过的 EMU 值再当厘米过一遍 `__new__` ⇒ 2 厘米变成 3.24e11
    （实测：`ValueError: value must be in range … got 324000000000`）。
    所以一律重新从磁盘读一份，省了又对。
    """
    from essay_tool.engine.alerts import AlertCollector

    config = load_preset(preset_path)
    extras: dict = {}
    description = toggle.apply(config, extras)
    if description is None:
        return None, "本夹具不适用"
    kwargs: dict = {}
    if needs_cover:
        from essay_tool.engine.cover import CoverFields
        kwargs["cover"] = CoverFields()
    kwargs.update(extras)
    shutil.rmtree(out_dir, ignore_errors=True)
    product = pipeline.run(source, None, out_dir, preset_config=config,
                           alerts=AlertCollector(), **kwargs)
    return product, description


def build_toggles(config, groups, styles, fields, source=None) -> list[Toggle]:
    """按分组把开关清单拼起来。`source` 只有两条开关要（长引文段号、index 直传）。"""
    toggles: list[Toggle] = []
    for group in groups:
        factory = GROUPS.get(group)
        if factory is None:
            print(f"（忽略未知分组：{group}）")
            continue
        if group == "格式":
            toggles.extend(factory(config, styles=styles, fields=fields,
                                   source=source))
        else:
            toggles.extend(factory(config, source=source))
    return toggles


def scan(source: pathlib.Path, preset_path: pathlib.Path, *,
         groups: list[str] | None = None, styles: set[str] | None = None,
         fields: set[str] | None = None, limit: int | None = None,
         keep_identical: bool = False,
         out_root: pathlib.Path | None = None) -> list[CaseOutcome]:
    """`out_root`：中间产物放哪（默认 `_out/scan`）。

    多进程批跑（`scan_all.py --jobs`）时**每个进程给一个自己的目录**——
    这里开头会 `rmtree`，共用目录会互相删现场（2026-09-19 加）。
    """
    config = load_preset(preset_path)
    toggles = build_toggles(config, groups or list(GROUPS), styles, fields,
                            source=source)
    if limit:
        toggles = toggles[:limit]
    print(f"共 {len(toggles)} 个开关待测", flush=True)

    out_root = pathlib.Path(out_root) if out_root else HERE / "_out" / "scan"
    shutil.rmtree(out_root, ignore_errors=True)      # 上一轮产物不留
    outcomes: list[CaseOutcome] = []
    sequence = 0

    for needs_cover in (False, True):
        batch = [t for t in toggles if t.needs_cover == needs_cover]
        if not batch:
            continue
        print(f"→ 跑基线（{'带自动封面' if needs_cover else '不带封面'}）…", flush=True)
        started = time.perf_counter()
        base_dir = out_root / f"_base{'_cover' if needs_cover else ''}"
        from essay_tool.engine.alerts import AlertCollector
        base_kwargs: dict = {}
        if needs_cover:
            from essay_tool.engine.cover import CoverFields
            base_kwargs["cover"] = CoverFields()
        shutil.rmtree(base_dir, ignore_errors=True)
        base_product = pipeline.run(source, None, base_dir,
                                    preset_config=load_preset(preset_path),
                                    alerts=AlertCollector(), **base_kwargs)
        base_shape = read_docx(base_product)
        print(f"   基线用时 {time.perf_counter() - started:.1f}s，"
              f"{len(base_shape.paras)} 段", flush=True)

        for toggle in batch:
            sequence += 1
            case_dir = out_root / f"{sequence:04d}-{toggle.group}"
            started = time.perf_counter()
            try:
                product, description = run_case(source, preset_path, case_dir, toggle,
                                                needs_cover=needs_cover)
            except Exception as error:                       # noqa: BLE001
                outcomes.append(CaseOutcome(toggle, f"{toggle.name}：{error}", None,
                                            0.0, skipped=True,
                                            skip_reason=f"跑失败：{error}"))
                continue
            elapsed = time.perf_counter() - started
            if product is None:
                outcomes.append(CaseOutcome(toggle, description, None, elapsed,
                                            skipped=True,
                                            skip_reason="本夹具不适用"))
                continue
            shape = read_docx(product)
            result = aat_diff.compare_shapes(base_shape, shape)
            outcome = CaseOutcome(toggle, description, result, elapsed)
            verdict, _, _ = outcome.verdict()
            # 只留"越界"那种当证据（"没变化"的产物跟基线一模一样，留着只占地方）
            if not (keep_identical or verdict == "escaped"):
                shutil.rmtree(case_dir, ignore_errors=True)
            outcomes.append(outcome)
            print(f"   [{toggle.group}] {description}  → {verdict} "
                  f"（{elapsed:.1f}s）", flush=True)
    shutil.rmtree(out_root / "_base_cover", ignore_errors=True)
    shutil.rmtree(out_root / "_base", ignore_errors=True)
    return outcomes


# ------------------------------------------------------------------ 报告


def render(outcomes: list[CaseOutcome], source: pathlib.Path,
           preset_path: pathlib.Path) -> str:
    lines: list[str] = []
    lines.append("=" * 78)
    lines.append("AAT 单因素扫描报告")
    lines.append(f"夹具：{safe_rel(source)}")
    lines.append(f"预设：{safe_rel(preset_path)}")
    lines.append("=" * 78)

    grouped: dict[str, list[CaseOutcome]] = {}
    for outcome in outcomes:
        grouped.setdefault(outcome.toggle.group, []).append(outcome)

    tally = {"in-scope": 0, "no-effect": 0, "escaped": 0, "skip": 0}
    for group in ("页面", "格式", "结构", "封面", "引注"):
        items = grouped.get(group)
        if not items:
            continue
        lines.append("")
        lines.append(f"———— {group}（{len(items)} 个开关）————")
        for outcome in items:
            verdict, escaped, coupled = outcome.verdict()
            tally[verdict] += 1
            buckets = {k: v for k, v in outcome.buckets().items() if v}
            if verdict == "skip":
                lines.append(f"  ⊘ {outcome.description} —— {outcome.skip_reason}")
                continue
            mark = {"in-scope": "✓", "no-effect": "⚠️", "escaped": "✘"}[verdict]
            lines.append(f"  {mark} {outcome.description}")
            if verdict == "no-effect":
                lines.append("      （完全没变化）"
                             + (f" {outcome.toggle.hint}" if outcome.toggle.hint else ""))
            else:
                summary = "、".join(f"{k} {v}" for k, v in sorted(buckets.items()))
                lines.append(f"      变化：{summary}")
            for item in escaped:
                lines.append(f"      ✘ 越界：{item}")
            for item in coupled:
                lines.append(f"      ○ 已知耦合：{item}")
            if verdict == "escaped" and outcome.result is not None:
                # 越界时把「到底是什么变了」打出来——光报"节变了"看不出所以然
                samples = (list(outcome.result.section_changes[:2])
                           + list(outcome.result.style_changes[:2])
                           + list(outcome.result.body_changes[:3]))
                for change in samples:
                    if change.note:
                        lines.append(f"         {change.where} [{change.note}]")
                        continue
                    pairs = "；".join(
                        f"{key}: {before or '（空）'} → {after or '（空）'}"
                        for key, (before, after)
                        in list(sorted(change.changes.items()))[:3])
                    lines.append(f"         {change.where} {change.text[:16]!r} {pairs}")

    lines.append("")
    lines.append("=" * 78)
    lines.append(f"汇总：有效 {tally['in-scope']} ／ 完全没变化 {tally['no-effect']} ／ "
                 f"越界 {tally['escaped']} ／ 跳过 {tally['skip']}")
    lines.append("说明：「完全没变化」不一定是 bug——也可能是这份夹具触发不到它"
                 "（每个开关要配一份能触发它的文档）。")
    return "\n".join(lines)


def main(argv: list[str]) -> int:
    import argparse
    parser = argparse.ArgumentParser(description="AAT 单因素扫描")
    parser.add_argument("source", type=pathlib.Path)
    parser.add_argument("--preset", type=pathlib.Path,
                        default=ROOT / "presets" / "中国政法大学-本科毕业论文.json")
    parser.add_argument("--groups", default="",
                        help="逗号分隔：页面,格式,结构,封面,引注（默认全部）")
    parser.add_argument("--styles", default="", help="只测这几个样式编号，如 S01,S02")
    parser.add_argument("--fields", default="", help="只测这几个字段，如 size,bold")
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--report", type=pathlib.Path, default=None)
    parser.add_argument("--list", action="store_true", help="只列开关，不跑")
    args = parser.parse_args(argv)

    groups = [g for g in args.groups.split(",") if g] or None
    styles = {s for s in args.styles.split(",") if s} or None
    fields = {f for f in args.fields.split(",") if f} or None
    if not args.source.exists():
        raise SystemExit(f"找不到夹具：{args.source}")

    if args.list:
        config = load_preset(args.preset)
        toggles = build_toggles(config, groups or list(GROUPS), styles, fields)
        for toggle in toggles:
            print(f"  [{toggle.group}] {toggle.name}  范围={'/'.join(sorted(toggle.scope))}")
        print(f"共 {len(toggles)} 个开关")
        return 0

    started = time.perf_counter()
    outcomes = scan(args.source, args.preset, groups=groups, styles=styles,
                    fields=fields, limit=args.limit)
    text = render(outcomes, args.source, args.preset)
    print()
    print(text)
    target = args.report or (HERE / "_out" / "scan" / "报告.txt")
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(text, encoding="utf-8")
    print()
    print(f"用时 {time.perf_counter() - started:.1f}s；报告：{safe_rel(target)}")
    return 1 if any(o.verdict()[0] == "escaped" for o in outcomes) else 0

if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
