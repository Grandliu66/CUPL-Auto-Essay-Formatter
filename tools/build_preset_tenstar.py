"""生成 presets/中国政法大学-学术十星.json。

依据：《第二十四届“学术十星”论文大赛参赛作品撰写格式规范》（附件二，
`reference/附件二 ：第二十四届“学术十星”论文大赛参赛作品撰写格式.pdf`）。

做法：以毕业论文预设为底，**只改规范明文规定了的项**，每一处都在下面标出
规范原句；规范没说的一律沿用毕业论文预设（不擅自发明），最后由本脚本回读
打印"与毕业论文预设的差异清单"供核对。

用法：.venv\\Scripts\\python.exe tools\\build_preset_tenstar.py
"""

from __future__ import annotations

import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING  # noqa: E402
from docx.shared import Cm, Pt  # noqa: E402

from essay_tool.config import preset_io, serialize  # noqa: E402
from essay_tool.config.units import Indent, IndentKind  # noqa: E402

PRESETS = ROOT / "presets"
BASE = PRESETS / "中国政法大学-本科毕业论文.json"
OUT = PRESETS / "中国政法大学-学术十星.json"

# 页眉：规范"用小四号宋体，居中排列，每字间空一格。页眉内容：中国政法大学“中伦杯”
# 第二十四届“学术十星”论文大赛参赛论文"。写法照官方封面范例里页眉的实物：汉字之间
# 各空一格，引号紧贴相邻汉字（全角引号自带空隙，不再加空格）。届数写成占位 `{届数}`：
# 导出时换成 GUI「封面」页签里用户填的届数（每字间空一格，如“二 十 四”），见
# `sections.expand_header_tokens`。
HEADER_TEXT = "中 国 政 法 大 学“中 伦 杯”第 {届数} 届“学 术 十 星”论 文 大 赛 参 赛 论 文"


def _chars(n: float) -> Indent:
    return Indent(IndentKind.CHARS, chars=n)


NO_INDENT = Indent(IndentKind.NONE)


def main() -> None:
    cfg = preset_io.load_preset(BASE)

    # ---- 页面：“页边距要求上下边距为2.5cm，左右边距为2.0cm”，A4 纵向
    p = cfg.page
    p.paper_size, p.orientation = "A4", "纵向"
    p.margin_top = p.margin_bottom = Cm(2.5)
    p.margin_left = p.margin_right = Cm(2.0)

    def style(sid: str):
        return cfg.styles[sid]

    def font(sid: str, east: str, ascii_: str, size_pt: float, bold: bool | None = None):
        s = style(sid)
        s.font_east_asian, s.font_ascii, s.size = east, ascii_, Pt(size_pt)
        if bold is not None:
            s.bold = bold

    # ---- 正文：“每段起行空两个字符（小四号），回行顶格。用小四号宋体字打印；单倍行距。”
    # 中文宋体；西文 Times New Roman（用户 2026-09-19 裁决：规范只写“宋体”，西文改回 TNR）
    font("S01", "宋体", "Times New Roman", 12)
    style("S01").first_line_indent = _chars(2)
    style("S01").left_indent = NO_INDENT
    style("S01").line_rule, style("S01").line_value = WD_LINE_SPACING.MULTIPLE, 1.0

    # ---- 标题：“正文一级标题位于上部居中位置，字体为三号加粗黑体。”
    font("S02", "黑体", "黑体", 16, True)
    style("S02").alignment = WD_ALIGN_PARAGRAPH.CENTER
    #      “正文二级标题，位于页面左端，顶格放置……四号加粗黑体字”
    font("S03", "黑体", "黑体", 14, True)
    style("S03").alignment = WD_ALIGN_PARAGRAPH.LEFT
    style("S03").first_line_indent = style("S03").left_indent = NO_INDENT
    #      “正文三级标题的字体为小四号加粗黑体，……空两格放置序数”。用户 2026-09-19 裁决：
    #      “按空两个半角空格来”＝序数前缩进两个半角空格的宽度＝1 个汉字宽（1 字符），
    #      三、四级同。
    font("S04", "黑体", "黑体", 12, True)
    style("S04").alignment = WD_ALIGN_PARAGRAPH.LEFT
    style("S04").first_line_indent = _chars(1)
    style("S04").left_indent = NO_INDENT
    #      “以下级别的标题字号均为小四号宋体，空两格放置序数”
    font("S05", "宋体", "宋体", 12, False)
    style("S05").alignment = WD_ALIGN_PARAGRAPH.LEFT
    style("S05").first_line_indent = _chars(1)
    style("S05").left_indent = NO_INDENT

    # ---- 中文题目：“中文论文题目字体为三号黑体字，居中打印”
    font("S06", "黑体", "黑体", 16, False)
    style("S06").alignment = WD_ALIGN_PARAGRAPH.CENTER
    # ---- 摘要：标签“四号加粗黑体”顶左格；正文“空两个中文字符……小四号宋体”
    font("S07", "黑体", "黑体", 14, True)
    font("S08", "宋体", "宋体", 12)
    style("S08").first_line_indent = _chars(2)
    # ---- 关键词：“内容下空一行用四号黑体字打印‘关键词’，其后是关键词，用小四号宋体”
    #      （四号黑体，**不加粗**；“空一行”＝段前一行，小四单倍行距约 14 磅）
    font("S09", "黑体", "黑体", 14, False)
    style("S09").space_before, style("S09").space_after = Pt(14), Pt(0)
    font("S10", "宋体", "宋体", 12)
    # ---- 英文题目：“三号字加粗 Times New Roman，居中打印”；“自成一页”→段前分页
    font("S11", "Times New Roman", "Times New Roman", 16, True)
    style("S11").alignment = WD_ALIGN_PARAGRAPH.CENTER
    style("S11").page_break_before = True
    # ---- Abstract：“四号加粗黑体”；西文字体统一 Times New Roman → 中文字体黑体、西文 TNR
    font("S12", "黑体", "Times New Roman", 14, True)
    #      摘要正文“小四号 Times New Roman”，空两个字符
    font("S13", "Times New Roman", "Times New Roman", 12)
    style("S13").first_line_indent = _chars(2)
    #      “Key Words”：“四号黑体字”（不加粗）；关键词“小四号 Times New Roman”
    font("S30", "黑体", "Times New Roman", 14, False)
    style("S30").space_before, style("S30").space_after = Pt(14), Pt(0)
    font("S31", "Times New Roman", "Times New Roman", 12)
    # ---- 目录：“‘目录’题头用三号加粗黑体字居中……目录独立成页”“目录正文字体为小四号宋体”
    font("S14", "黑体", "黑体", 16, True)
    style("S14").alignment = WD_ALIGN_PARAGRAPH.CENTER
    style("S14").page_break_before = True
    for sid in ("S15", "S16", "S17"):
        font(sid, "宋体", "宋体", 12)
    # ---- 图表：“图序及图名置于下方，字体为小四号加粗宋体”“表序及表名位于表上方居中位置，
    #      字体为小四号加粗宋体”
    font("S24", "宋体", "宋体", 12, True)
    font("S25", "宋体", "宋体", 12, True)
    # ---- 页眉/页脚：页眉“小四号宋体，居中”（预设 S27 已是）；页脚“五号宋体，居中”
    font("S27", "宋体", "宋体", 12, False)
    font("S28", "宋体", "宋体", 10.5, False)
    style("S28").alignment = WD_ALIGN_PARAGRAPH.CENTER
    for sid in ("S02", "S03", "S04", "S05", "S06", "S07", "S08", "S09", "S10", "S11",
                "S12", "S13", "S14", "S24", "S25", "S27", "S28", "S30", "S31"):
        cfg.styles_by_name[cfg.styles[sid].name] = cfg.styles[sid]
    for sid in ("S01", "S15", "S16", "S17"):
        cfg.styles_by_name[cfg.styles[sid].name] = cfg.styles[sid]

    # ---- 页眉页脚：两节内容相同；页码只有“页脚内容为页码（封面不计入页数）”，
    #      前置页罗马数字是用户的暂定裁决（规范未要求）
    for rule in cfg.header_footer:
        rule.header_text = HEADER_TEXT
    cfg.header_bottom_line = True                       # 官方范例页眉有下边框线
    cfg.header_line_extend_left_cm = cfg.header_line_extend_right_cm = 1.5

    # ---- 结构：规范“由封面、题目、中英文摘要、关键词、目录、正文、附录七部分构成”
    #      附录“（必要时）”“不是必需的”；正文含“附加的文献”（参考文献）；无致谢、无清单、
    #      无副标题（副标题只在封面）
    on = {"题目", "摘要", "关键词", "英文题目", "英文摘要", "英文关键词", "目录", "参考文献"}
    for name in cfg.structure_required:
        cfg.structure_required[name] = name in on
    # ---- 标签写法：摘要/关键词无【】；“Abstract”“Key Words”；“目录”“绪论”两字中间
    #      空两个字符（沿用本项目"空两格＝两个空格"的写法，如"致  谢"）
    cfg.structure_insert_text.update({
        "摘要": "摘要", "关键词": "关键词",
        "英文摘要": "Abstract", "英文关键词": "Key Words",
        "目录": "目  录", "引言": "绪  论",
        "参考文献": "参考文献", "附录": "附录",
    })
    # ---- 封面：学术十星封面
    cfg.cover_enabled, cfg.cover_form = True, "学术十星封面"

    preset_io.save_preset_json(cfg, OUT)
    print("写出:", OUT)

    # ---- 回读，逐项打印与毕业论文预设的差异
    a = serialize.preset_to_dict(preset_io.load_preset(BASE))
    b = serialize.preset_to_dict(preset_io.load_preset(OUT))

    def walk(x, y, path=""):
        if isinstance(x, dict) and isinstance(y, dict):
            for k in sorted(set(x) | set(y)):
                walk(x.get(k), y.get(k), f"{path}/{k}")
        elif isinstance(x, list) and isinstance(y, list) and len(x) == len(y):
            for i, (u, v) in enumerate(zip(x, y)):
                key = u.get("id") if isinstance(u, dict) and "id" in u else i
                walk(u, v, f"{path}[{key}]")
        elif x != y:
            print(f"  {path}: {x!r} → {y!r}")

    walk(a, b)


if __name__ == "__main__":
    main()

