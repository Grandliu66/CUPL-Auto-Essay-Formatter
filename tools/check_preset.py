"""用法：.venv\\Scripts\\python.exe tools\\check_preset.py

读一遍 presets/*.xlsx，把解析结果打出来，检查换算对不对。
"""

import dataclasses
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))

from essay_tool.config.excel_io import ConfigError, load_preset

PATH = pathlib.Path(__file__).resolve().parent.parent / "presets" / "中国政法大学-本科毕业论文.xlsx"


def main():
    try:
        cfg = load_preset(PATH)
    except ConfigError as e:
        print(f"配置错误：{e}")
        raise SystemExit(1)

    p = cfg.page
    print(f"页面：{p.paper_size} {p.orientation}，上下左右边距 {p.margin_top} {p.margin_bottom} {p.margin_left} {p.margin_right}")
    print(f"样式数量：{len(cfg.styles)}")
    for s in cfg.styles.values():
        print(f"  {s.id} {s.name}")
    print("  前 3 个的细节：")
    for s in list(cfg.styles.values())[:3]:
        print(f"  {s.id} {s.name}: {s.font_east_asian}/{s.font_ascii} {s.size} 加粗={s.bold} "
              f"对齐={s.alignment} 首行缩进={s.first_line_indent} 行距={s.line_rule}/{s.line_value}")
    print(f"页眉页脚规则数：{len(cfg.header_footer)}")
    for hf in cfg.header_footer:
        print(
            f"  {hf.section_name}: 页眉「{hf.header_text}」 "
            f"启用页码={hf.page_number_enabled} 编码={hf.page_number_format} "
            f"前缀「{hf.page_number_prefix}」 后缀「{hf.page_number_suffix}」"
        )
    c = cfg.citation
    # 引注这几列以前只打两个，结果“序号后空多少”“标记用什么字体”这些列
    # 埋了很久没人记得。全部打出来。
    print("引注配置：")
    for field in dataclasses.fields(c):
        print(f"  {field.name} = {getattr(c, field.name)!r}")

    kept = cfg.keep_run_formatting
    if kept is None:
        print("清理规则：预设里没配这一节（用代码默认值）")
    elif not kept:
        print("清理规则：源文档的字符格式**一个都不保留**（全清）")
    else:
        print(f"清理规则：源文档里保留这些字符格式 → {sorted(kept)}")
    print(f"页眉线：{'要（页眉文字下画横线）' if cfg.header_bottom_line else '不要（不画线）'}")
    print(f"封面悬挂缩进：{'要（超长值续行对齐提示词右侧）' if cfg.cover_hanging_indent else '不要（续行回最左）'}")
    print(f"结构部件启用数：{sum(cfg.structure_required.values())}/{len(cfg.structure_required)}")
    # 明细很重要：只报个总数的话，判断不了"缺某个部件时会怎么补"——
    # 没启用的部件根本不会被补，代码里那条插入路径也就走不到。
    for name, enabled in cfg.structure_required.items():
        print(f"  {'✔' if enabled else '✘'} {name}  → 插入文字 {cfg.structure_insert_text.get(name, '')!r}")
    print("OK")


if __name__ == "__main__":
    main()
