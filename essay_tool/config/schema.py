"""预设配置在内存里的样子：读完预设文件（JSON 或 Excel，见 preset_io.py）之后得到的结构化数据。"""

from __future__ import annotations

from dataclasses import dataclass, field

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Emu

from essay_tool.config.units import Indent

# 「封面」小节里「封面形态」一格允许的两个值（也是界面上那两个选项的名字，
# 见 `gui.main_window.COVER_FORMS`）。放到这里是因为 Excel 预设的写/读
# （`excel_writer` / `excel_io`）都要用同一份清单对一下——以前这两个词只
# 写在 JSON 预设里，Excel 那条路上封面开关/形态根本传不过去（审计
# 2026-09-15 核实）。
COVER_FORM_NAMES = ("毕业论文封面", "学年论文封面", "课程考核封面", "学术十星封面")

# 只有"标题写法"、没有"是否启用"的部件：引言/结论是正文里的一级标题，位置即
# 身份，没有开关（2026-09-16），但它们的**写法**（"引　言""绪　论"这类字间空格、
# 有没有【】）可以在预设的 `structure_insert_text` 里给一格；不给＝保持作者原文。
# 这两项**不进** `structure_required`——进去引擎就会当成"缺了要补"的部件。
HEADING_LABEL_PARTS = ("引言", "结论")


def heading_label_with_gap(name: str, gap: int) -> str:
    """引言/结论的写法模板：两字之间空 `gap` 个空格（"引  言"）。规范的说法就是
    "引言、结论两字之间空两格"，界面上给一个数字，存进预设时生成这个模板。"""
    return name[0] + " " * gap + name[1:]


def heading_label_gap(text: str) -> int | None:
    """写法模板里两字之间空了几格；不是"两个字＋中间只有空白"的形状（如
    "【引言】"）就返回 None——界面上表达不了，由调用方决定怎么处理。
    只数空白字符的个数，半角、全角空格一视同仁；模板里的词不看（引言家族认出来
    后本来就保留作者写的词）。"""
    text = (text or "").strip()
    if len(text) < 2 or text[0].isspace() or text[-1].isspace():
        return None
    middle = text[1:-1]
    return len(middle) if all(ch.isspace() for ch in middle) else None

# 封面形态名（界面/预设里的中文名）→ 引擎里的封面类型键（`engine.cover.add_cover`
# / `pipeline.run(cover_type=…)` 认的那几个）。形态名单与这张表**只此一份**，
# 界面卡片、Excel 下拉、导出调用都从这里取，不要各写各的。
COVER_TYPE_BY_FORM = {
    "毕业论文封面": "thesis",
    "学年论文封面": "year",
    "课程考核封面": "course",
    "学术十星封面": "tenstar",
}


@dataclass
class PageSetup:
    paper_size: str  # 原始写法（"A4" 或自定义尺寸文字），仅供显示/调试
    page_width: Emu  # 纵向下的标准宽（自定义尺寸时按填写顺序）
    page_height: Emu
    orientation: str
    margin_top: Emu
    margin_bottom: Emu
    margin_left: Emu
    margin_right: Emu
    gutter: Emu
    header_distance: Emu
    footer_distance: Emu


@dataclass
class ParagraphStyleSpec:
    """样式表里的一行，对应一个 python-docx 会用到的段落/字符样式。"""

    id: str
    name: str
    font_east_asian: str
    font_ascii: str
    size: Emu  # 实为 Pt，但类型上是 Length
    bold: bool
    alignment: WD_ALIGN_PARAGRAPH
    first_line_indent: Indent
    left_indent: Indent
    space_before: Emu | None
    space_after: Emu | None
    line_rule: WD_LINE_SPACING
    line_value: Emu | float | None
    page_break_before: bool


@dataclass
class HeaderFooterRule:
    section_name: str
    header_text: str
    page_number_enabled: bool
    page_number_format: str | None  # OOXML w:pgNumType fmt 值（decimal/upperRoman/...），未启用时为 None
    page_number_prefix: str
    page_number_suffix: str


@dataclass
class CitationConfig:
    footnote_marker_style: str
    footnote_restart_each_page: bool
    footnote_marker_font: str
    footnote_marker_trailing_space: str
    reference_numbering_style: str


@dataclass
class PresetConfig:
    page: PageSetup
    styles: dict[str, ParagraphStyleSpec] = field(default_factory=dict)  # 按编号(S01...)索引
    styles_by_name: dict[str, ParagraphStyleSpec] = field(default_factory=dict)
    header_footer: list[HeaderFooterRule] = field(default_factory=list)
    citation: CitationConfig | None = None
    structure_required: dict[str, bool] = field(default_factory=dict)  # 部件名 -> 是否启用
    # 部件名 -> 标签/标题的写法（缺失时插入的文字，也是识别到后改写成的写法）。
    # **只有"标签"类部件的写法可以自定义**（摘要/关键词/英文摘要/英文关键词/目录/
    # 三张清单/参考文献/附录/致谢，外加 `HEADING_LABEL_PARTS` 的引言/结论）；
    # 题目/副标题/英文题目/英文副标题的"内容占位语"是工具行为，不是排版规范，
    # 不许自定义——它们这一格恒为空。格子空＝识别到后保持作者原文。
    structure_insert_text: dict[str, str] = field(default_factory=dict)
    # 源文档正文里要保留的字符格式（`w:rPr` 的子元素名，见 cleanup.py）。
    # None = 这次没配（旧版预设 Excel 里没这一节）→ 用代码里的默认值。
    # 空集合 = 用户明确说了"全清"。这两个是**不同**的意思，别合并。
    keep_run_formatting: frozenset[str] | None = None
    # 页眉文字下面那条横线要不要画（用户 2026-09-12 要求加开关）。官方模板
    # 把线做在"页眉"样式的下边框上；True = 保证有；False = 摘掉已有的下边框
    # （源文档自带的也算）。**封面不受影响**——封面节不放页眉，本来就没有
    # 这条线。见 styles.set_header_bottom_border。
    #
    # 默认值从 True 改成 False（用户 2026-09-14 定"不套用任何预设时"的
    # 默认值：页眉装饰线选否）——预设里写了这一项的照预设来（内置毕业论文
    # 那份写的还是"要"）。
    header_bottom_line: bool = False
    # 页眉线向正文栏**外**伸出多少厘米（左、右各一个；0＝不伸出，线跟正文栏
    # 等宽，老行为）。做法是给"页眉"样式设负的左右缩进——边框画在段落的缩进
    # 范围内，缩进变负，线就伸出去。学术十星的页眉线比正文栏宽（实测左伸
    # 1.5cm、右伸 1.35cm，取对称的 1.5/1.5）。只在 `header_bottom_line` 为真时才生效。
    header_line_extend_left_cm: float = 0.0
    header_line_extend_right_cm: float = 0.0
    # 封面信息行的"悬挂缩进"：值太长换行时，续行对齐到提示词右侧（首行不
    # 变）。默认开（用户 2026-09-12 接板"加开关、默认开启"），见「封面」小节。
    # 实现逻辑在 tools/make_cover_samples.build_thesis_cover(hanging_indent=…)；
    # **封面生成接入流水线后在那边生效**。
    cover_hanging_indent: bool = True
    # 这两项 2026-09-14 新加："这个预设要不要带封面、用哪种封面"（用户做
    # 课程考核预设时要求"封面开关打开，样式为课程考核"）。封面里**填写的
    # 内容**照旧不进预设（每次论文都不一样，界面上也写着"该内容将不会保存
    # 进预设"）。
    #   cover_enabled: True=应用这个预设时把封面开关打开；False=关掉；
    #                  None=预设里没写（旧预设）→ 按 True 处理（用封面）
    #   cover_form: `COVER_FORM_NAMES` 里的一个；None=预设里没写 →
    #               按预设名里有没有"课程"来猜（旧预设的行为）
    cover_enabled: bool | None = None
    cover_form: str | None = None
    # ---- 清理规则（用户 2026-09-20，格式页「清理规则」块）。**都不作用于封面范围。**
    # 清除文档里所有手打分页符（Ctrl+Enter 的 `w:br type="page"`；分节符不算）。
    clear_page_breaks: bool = False
    # 清除文档里所有"确实是空的"段落（没有任何字符、图片/对象、域、分节符、自动编号）。
    clear_blank_lines: bool = False
    # 标签前"另起一页"的按部件覆盖：{部件名: 要不要}。**没写的部件跟随它套用的样式**
    # （样式表"段前分页"那一列）——这个字典只记"跟样式不一样"的那几个。摘要/关键词/
    # 英文摘要/英文关键词/目录/三张清单/引言/结论/参考文献/附录/致谢。
    label_page_break: dict[str, bool] = field(default_factory=dict)

    def style(self, id_or_name: str) -> ParagraphStyleSpec:
        if id_or_name in self.styles:
            return self.styles[id_or_name]
        if id_or_name in self.styles_by_name:
            return self.styles_by_name[id_or_name]
        raise KeyError(f'预设里没有样式 "{id_or_name}"')
