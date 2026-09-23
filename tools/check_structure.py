"""用法：.venv\\Scripts\\python.exe tools\\check_structure.py <docx路径> [--preset xlsx路径]

打印文档结构清单：哪些部件找到了、哪些必需/可选的部件缺失。
"""

import argparse
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))

from docx import Document

from essay_tool.config.excel_io import ConfigError, load_preset
from essay_tool.engine.structure import build_structure_inventory

DEFAULT_PRESET = pathlib.Path("presets/中国政法大学-本科毕业论文.xlsx")


def main():
    parser = argparse.ArgumentParser(description="打印文档结构清单")
    parser.add_argument("docx", type=pathlib.Path, help="用户的 .docx")
    parser.add_argument("--preset", type=pathlib.Path, default=DEFAULT_PRESET)
    args = parser.parse_args()

    try:
        preset = load_preset(args.preset)
    except ConfigError as e:
        print(f"出错：{e}")
        raise SystemExit(1)

    doc = Document(args.docx)
    inv = build_structure_inventory(doc, preset.structure_required)

    def line(p):
        loc = f"第{p.paragraph_index}段" if p.found else "缺失"
        req = "必需" if p.required else "可选"
        return f"  {p.name:6s} [{req}] {loc}"

    print(line(inv.title))
    print(line(inv.subtitle))
    for p in inv.parts:
        print(line(p))

    print()
    print("缺失的必需部件:", [p.name for p in inv.missing_required] or "无")
    print("缺失的可选部件:", [p.name for p in inv.missing_optional] or "无")


if __name__ == "__main__":
    main()
