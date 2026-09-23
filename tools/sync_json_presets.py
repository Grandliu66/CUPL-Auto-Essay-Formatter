"""把 presets/*.xlsx 转成/更新对应的 presets/*.json。

**JSON 才是本工具默认读取/存盘的格式**（GUI 的配置文件列表、「保存当前
设置为新配置」都是落 JSON；Excel 只是给人可视化查看/修改的边缘产品）。
`presets/` 目录下随程序分发的内置默认预设因此需要两份：`build_presets.py`
生成的 `.xlsx`（人可视化编辑用）和跟它内容一致的 `.json`（程序实际读取的
那份）。改了 `tools/build_presets.py` 之后，除了要重新跑一遍生成新的
`.xlsx`，还要接着跑这个脚本，把内置 JSON 预设同步更新——两份文件不会
自动保持一致。

用法：
    .venv\\Scripts\\python.exe tools\\build_presets.py     # 先重新生成 xlsx
    .venv\\Scripts\\python.exe tools\\sync_json_presets.py  # 再同步出 json
"""

from __future__ import annotations

import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from essay_tool.config.preset_io import import_from_excel  # noqa: E402

PRESET_DIR = ROOT / "presets"


def main() -> None:
    xlsx_files = sorted(PRESET_DIR.glob("*.xlsx"))
    if not xlsx_files:
        print("presets/ 下没有找到 .xlsx 文件。")
        return
    for xlsx in xlsx_files:
        json_path = xlsx.with_suffix(".json")
        import_from_excel(xlsx, json_path)
        print(f"已同步：{xlsx.name} -> {json_path.name}")


if __name__ == "__main__":
    main()
