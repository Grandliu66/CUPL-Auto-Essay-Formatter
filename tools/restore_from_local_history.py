"""从 VS Code 本地历史恢复被旧缓冲区覆盖的 5 个文件。

事故（2026-09-16 19:24:12）：detect.py / cleanup.py / styles.py / structure.py /
spec_viewer.py 五个文件被同时写回**旧一代内容**（编辑器里开着的陈旧缓冲区被保存），
当前 detect.py 只有 40542 字节，而 19:24 之前的好版本是 67740 字节。

本脚本默认只做**只读比对**（列函数清单差异、检查关键函数是否齐全），加 `--apply`
才真正写回，写回前把当前（受损）文件备份到 work/_damaged_20260916_1924/。

跑法：.venv\\Scripts\\python.exe tools\\restore_from_local_history.py [--apply]
"""

from __future__ import annotations

import ast
import datetime
import os
import pathlib
import shutil
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = pathlib.Path(__file__).resolve().parents[1]
HISTORY = pathlib.Path(os.environ["APPDATA"]) / "Code" / "User" / "History"
DAMAGED_DIR = ROOT / "work" / "_damaged_20260916_1924"

# 目标文件 -> (本地历史快照, 快照时间说明)
RESTORE = {
    "essay_tool/engine/detect.py": ("-14581d9", "bDVD.py", "09-16 18:00:57"),
    "essay_tool/engine/cleanup.py": ("1b50bf00", "dSfi.py", "09-15 16:37:14"),
    "essay_tool/engine/styles.py": ("-26849078", "mo5G.py", "09-16 10:31:30"),
    "essay_tool/engine/structure.py": ("a1d9a91", "sk4r.py", "09-16 12:45:29"),
    "essay_tool/gui/spec_viewer.py": ("-6ce4b44b", "oeqS.py", "09-16 12:45:29"),
}

REQUIRED = {
    "essay_tool/engine/detect.py": (
        "has_toc_field", "is_toc_paragraph", "list_field_kinds", "body_toc_field_parts",
        "body_list_field_kinds", "detect_appendix_entries", "detect_english_subtitle",
        "detect_english_title", "detect_front_matter_labels", "TOC_PAGE_PARTS",
        "hand_typed_toc_entries", "wears_caption_style", "detect_captions",
    ),
    "essay_tool/engine/cleanup.py": (
        "tables_in_prefix", "drop_regenerated_fields", "unwrap_section_break_sdt",
        "strip_paragraph_direct_formatting", "set_run_text",
    ),
}


def defs(path: pathlib.Path) -> set[str]:
    """模块里所有 def / class 的名字（拿不到 AST 就退回正则）。"""
    text = path.read_text(encoding="utf-8-sig", errors="replace")
    try:
        tree = ast.parse(text)
    except SyntaxError:
        import re
        return set(re.findall(r"^(?:def|class)\s+(\w+)", text, re.M))
    names: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            names.add(node.name)
    return names


def main() -> None:
    apply = "--apply" in sys.argv
    print("模式：" + ("写回" if apply else "只读比对（加 --apply 才写回）") + "\n")

    for resource, (folder, snap, when) in RESTORE.items():
        live = ROOT / resource.replace("/", os.sep)
        source = HISTORY / folder / snap
        if not source.exists():
            print(f"！{resource}：快照不存在 {source}")
            continue

        print(f"===== {resource} =====")
        print(f"  当前（受损）{live.stat().st_size} 字节   快照（{when}）{source.stat().st_size} 字节")
        old, new = defs(live), defs(source)
        back = sorted(old - new)
        gain = sorted(new - old)
        print(f"  快照比当前多出 {len(gain)} 个定义：{gain}")
        print(f"  当前比快照多出 {len(back)} 个定义：{back}")

        text = source.read_text(encoding="utf-8-sig", errors="replace")
        need = REQUIRED.get(resource)
        if need:
            missing = [name for name in need if f"def {name}" not in text]
            print(f"  关键函数检查：{'全部在位' if not missing else '缺 ' + str(missing)}")
        if "廿" in text or "hundred" in text:
            print("  含廿/卅 那套中文数字处理：", "廿" in text)

        if apply:
            DAMAGED_DIR.mkdir(parents=True, exist_ok=True)
            keep = DAMAGED_DIR / (live.name + "_damaged.py")
            shutil.copy2(live, keep)
            shutil.copy2(source, live)
            print(f"  已写回；受损版本存 {keep.relative_to(ROOT)}")
        print()

    if not apply:
        print("（只读比对结束，未改动任何文件）")


if __name__ == "__main__":
    main()
