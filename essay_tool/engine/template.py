"""空白模板生成：不碰任何源文件，按预设（JSON/Excel 均可，走 preset_io）直接造一份"填着占位提示的空壳"。

原理：引擎本来就是"缺什么补什么"——一份**完全空白**的 docx 就是"所有部件
都缺"的极端情况。把空白文档直接喂给 `pipeline.run()`：识别阶段什么都找不到，
随后 `apply_missing_parts` 按规范顺序把题目/副标题/摘要/关键词/英文摘要/目录/
引言/结论/参考文献/附录/致谢……全部补上，每处塞一句占位提示
（"［请在此处填写……］"），页边距、页码、页眉页脚、目录域、正文分节也都照
正常管线配置——**模板文件就这样免费得到**，不需要为"无源文件"写第二套逻辑。

运行产物：题目/摘要/关键词/英文摘要页/目录（带目录域）/引言/结论/参考文献……
每处一句占位提示；中间还有一块**正文示例区**（一～四级标题＋一段正文，
照规范写法示范），页边距、页码、页眉页脚也都按管线正常配置。

用法：
    .venv\\Scripts\\python.exe essay_tool\\engine\\template.py
    .venv\\Scripts\\python.exe essay_tool\\engine\\template.py --out work\\我的空白模板.docx
    .venv\\Scripts\\python.exe essay_tool\\engine\\template.py --cover    # 顺带拼封面（占位文字）

给主 GUI 接线：
    from essay_tool.engine.template import build_template
    build_template(preset_path, out_path)        # "导出空白模板"按钮里这一行
"""

from __future__ import annotations

import argparse
import pathlib
import shutil
import sys

ROOT = pathlib.Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from docx import Document  # noqa: E402

from essay_tool.config.preset_io import load_preset  # noqa: E402
from essay_tool.engine.alerts import AlertCollector  # noqa: E402
from essay_tool.engine.cover import CoverFields  # noqa: E402
from essay_tool.engine.pipeline import DEFAULT_PRESET, run  # noqa: E402

# 模板里的"正文示例区"：完整演示一～四级标题＋正文的规范写法。故意把示例
# 写进"源文件"里、让管线当真实内容处理——标题识别、编号规范化、样式套用
# 全走正常路径，示例段的标题样式（S02～S05）也就自然跟着建进文档里了。
# 编号对齐 `recognize.ChineseHeadingGrammar` 的规范写法：一级"一、"后不空格，
# 二三四级全角括号（半角句点）＋一个半角空格。
_SAMPLE_BODY = (
    "一、此处写一级标题",
    "（一） 此处写二级标题",
    "1. 此处写三级标题",
    "（1） 此处写四级标题",
    "此处写正文内容；正文按自然段书写。",
)


def build_template(
    preset_path: pathlib.Path | None = None,
    out_path: pathlib.Path = pathlib.Path("work/空白模板.docx"),
    work_dir: pathlib.Path = pathlib.Path("work"),
    with_cover: bool = False,
    cover_type: str = "thesis",
    cover_fields: CoverFields | None = None,
    all_parts: bool = False,
    preset_config=None,
    alerts: AlertCollector | None = None,
    session: str = "",
) -> pathlib.Path:
    """按预设生成空白模板到 out_path，返回实际写出的路径。

    - 中间产物（空白"源文件"和它的工作副本）落在 work_dir 里，可随时重跑；
    - with_cover=True 时封面一并拼上；`cover_type` 选封面形态（"thesis"/
      "year"/"course"/"tenstar"，同 `pipeline.run`）。`cover_fields` 传了就用它（GUI「封面」
      页签"直接填写"的那些真实值）；不传（比如命令行直接跑）就用"［……］"
      占位文字，换一套字段名跟着 cover_type 走；
    - 默认"模板跟着预设配置走"（没启用的部件不出现）；all_parts=True 则
      无视预设开关、把所有结构都补上（查看完整骨架用）。三张图表清单例外：
      没题注时引擎不生成空清单（空域刷新后会显示"没有指定样式的文字"
      报错字样，很难看），等论文里有图/表了再自动补。
    - `preset_config` 传了就直接用（GUI 页签当场的表单值），不再按
      `preset_path` 重新读文件——跟 `pipeline.run()` 的同名参数是同一个
      问题、同一个修法：只认预设文件的话，用户在页签里临时改的东西会被
      无声吃掉。两者传一个即可，`preset_config` 优先。
    """
    if preset_config is None and preset_path is None:
        # 两个都不传时以前直接 `load_preset(None)` → `pathlib.Path(None)` 抛
        # TypeError（审计 2026-09-16 报的）：签名里的默认值其实不可用。这里
        # 给一句能看懂的报错。
        raise ValueError(
            "build_template() 需要 preset_config 或 preset_path 之一："
            "前者是 GUI 页签当场的表单值，后者是预设文件路径")
    preset = preset_config if preset_config is not None else load_preset(preset_path)
    override = None
    if all_parts:
        override = {part: True for part in preset.structure_required}
    work_dir.mkdir(parents=True, exist_ok=True)
    # 空白"源文件"放在子目录里：loader 会把源文件原样复制成 work_dir/源文件名，
    # 直接放 work_dir 里会和自己撞名（SameFileError）。
    src_dir = work_dir / "_template"
    src_dir.mkdir(parents=True, exist_ok=True)
    blank_source = src_dir / "_blank.docx"
    doc = Document()
    for text in _SAMPLE_BODY:                   # 正文示例区（见 _SAMPLE_BODY 注释）
        doc.add_paragraph(text)
    doc.save(str(blank_source))

    cover = None
    if with_cover:
        if cover_fields is not None:
            cover = cover_fields
        elif cover_type == "course":
            cover = CoverFields(
                course_name="［课程名称］", course_code="［课程代码］",
                course_seq="［课程序号］", teacher_name="［任课教师］",
            )
        elif cover_type == "tenstar":
            cover = CoverFields(title="［论文题目］", session=session)   # 十星封面只有题目/副标题/届数
        else:                                           # thesis / year 字段相同
            cover = CoverFields(
                title="［论文题目］", student_name="［姓名］", student_id="［学号］",
                college="［学院］", major="［专业］", grade="［年级］",
                advisor_name="［指导教师］",
            )

    result = run(blank_source, None, work_dir,
                 alerts=alerts if alerts is not None else AlertCollector(),
                 cover=cover, cover_type=cover_type, structure_override=override,
                 preset_config=preset, session=session,
                 # 这里的"源文件"是上面现造的示例正文（四个孤立标题那种），
                 # 围着它报的告警全是噪声——2026-09-14 接线时实测过：空白
                 # 模板会稳定报 4 条"标题孤立"，纯扰人。
                 check_source_alerts=False)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(result, out_path)
    return out_path


def main() -> None:
    parser = argparse.ArgumentParser(description="按预设生成空白写作模板")
    parser.add_argument("--preset", type=pathlib.Path, default=DEFAULT_PRESET)
    parser.add_argument("--out", type=pathlib.Path, default=pathlib.Path("work/空白模板.docx"))
    parser.add_argument("--work-dir", type=pathlib.Path, default=pathlib.Path("work"))
    parser.add_argument("--cover", action="store_true", help="顺带拼封面（占位文字）")
    parser.add_argument("--cover-type", choices=("thesis", "year", "course", "tenstar"),
                        default="thesis",
                        help="封面形态：毕业论文（默认）/ 学年论文 / 课程考核 / 学术十星")
    parser.add_argument("--all-parts", action="store_true",
                        help="无视预设开关，把所有结构都补上（查看完整骨架用）")
    parser.add_argument("--session", default="二十四",
                        help="届数（学术十星的封面和页眉用），如“二十四”；默认二十四")
    args = parser.parse_args()

    # 告警走自己的收集器：生成完顺手打一份报告，命令行跑也能看见
    # （GUI 那条路径是自己建收集器、弹窗显示，见 on_export_blank）。
    alerts = AlertCollector()
    out = build_template(args.preset, args.out, args.work_dir,
                         with_cover=args.cover, cover_type=args.cover_type,
                         all_parts=args.all_parts, alerts=alerts, session=args.session)
    print(f"已生成：{out}")
    if alerts.items:
        alerts.print_report()


if __name__ == "__main__":
    main()
