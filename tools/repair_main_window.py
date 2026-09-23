"""一键修复 main_window.py：编辑器把陈旧缓冲区保存回来时，用它救回并重做那两处改动。

判据（不靠记时间）：
  * 「新一代」文件必须有 `_refresh_log_bar`、`_close_cover_zoom` 之类 09-15 之后才有的东西，
    且**不向 AdjustWindow 传 `status_callback=`**（该参数 09-15 已删）；
  * 候选快照＝本地历史里满足上述条件、且时间戳最早的一次事故（2026-09-16 19:24:12）之前的最近一份。

修复动作：
  1. 用那份快照覆盖当前文件（当前受损版本先备份到 work/_damaged_*/）；
  2. 若缺 `operation_logic_text` 的导入就补上；
  3. 若还有内联的 `OPERATION_LOGIC_TEXT = \"\"\"…\"\"\"` 就剔掉（改成指路注释）。

跑法：.venv\\Scripts\\python.exe tools\\repair_main_window.py [--dry-run]
"""

from __future__ import annotations

import datetime
import hashlib
import json
import os
import pathlib
import re
import shutil
import sys
import urllib.parse

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = pathlib.Path(__file__).resolve().parents[1]
HISTORY = pathlib.Path(os.environ["APPDATA"]) / "Code" / "User" / "History"
LIVE = ROOT / "essay_tool" / "gui" / "main_window.py"
RESOURCE = "/essay_tool/gui/main_window.py"
CLOBBER = datetime.datetime(2026, 9, 16, 19, 24, 12).timestamp() * 1000
DAMAGED_DIR = ROOT / "work" / "_damaged_main_window"

NEEDED = ("_refresh_log_bar", "REVIEW_ALERTS_HINT", "wrap_text")
FORBIDDEN = ("status_callback=",)
IMPORT_LINE = "from essay_tool.gui.operation_logic_text import OPERATION_LOGIC_TEXT  # noqa: E402"
IMPORT_ANCHOR = "from essay_tool.gui.text_wrap import wrap_text  # noqa: E402"
INLINE_BLOCK = re.compile(r"# 「工具操作逻辑」正文.*?OPERATION_LOGIC_TEXT = \"\"\".*?\"\"\"\n", re.S)
NEW_COMMENT = (
    "# 「工具操作逻辑」正文由工作簿生成，见 essay_tool/gui/operation_logic_text.py；\n"
    "# 那份文件由 tools/make_operation_logic_text.py 从\n"
    "# reference/工具操作逻辑写作表.xlsx 重出。要改文字就改工作簿再重跑那个脚本。\n"
)


def wanted(text: str) -> bool:
    return all(m in text for m in NEEDED) and not any(m in text for m in FORBIDDEN)


def snapshots() -> list[tuple[int, pathlib.Path]]:
    rows: list[tuple[int, pathlib.Path]] = []
    for folder in HISTORY.iterdir():
        entries_file = folder / "entries.json"
        if not entries_file.exists():
            continue
        try:
            data = json.loads(entries_file.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        if not urllib.parse.unquote(data.get("resource", "")).endswith(RESOURCE):
            continue
        for entry in data.get("entries", []):
            path = folder / entry["id"]
            if path.exists():
                rows.append((entry["timestamp"], path))
    return sorted(rows)


def main() -> None:
    dry = "--dry-run" in sys.argv
    live_text = LIVE.read_text(encoding="utf-8-sig", errors="replace")
    print(f"当前 main_window.py：{len(live_text.splitlines())} 行 / {LIVE.stat().st_size} 字节")
    print(f"  含 {NEEDED}：{[m in live_text for m in NEEDED]}；"
          f"含 {FORBIDDEN}：{[m in live_text for m in FORBIDDEN]}")

    candidates = [(ts, path) for ts, path in snapshots()
                  if ts < CLOBBER
                  and wanted(path.read_text(encoding="utf-8-sig", errors="replace"))]
    if not candidates:
        raise SystemExit("！本地历史里找不到「新一代」快照，无法自动修复")
    ts, source = candidates[-1]
    print(f"选用快照：{datetime.datetime.fromtimestamp(ts / 1000):%m-%d %H:%M:%S}  "
          f"{source.stat().st_size} 字节  {source.parent.name}/{source.name}")

    text = source.read_text(encoding="utf-8-sig", errors="replace")
    if "operation_logic_text" not in text:
        text = text.replace(IMPORT_ANCHOR, f"{IMPORT_ANCHOR}\n{IMPORT_LINE}", 1)
        print("  补上 operation_logic_text 的导入")
    if INLINE_BLOCK.search(text):
        text = INLINE_BLOCK.sub(NEW_COMMENT, text, count=1)
        print("  剔掉内联的 OPERATION_LOGIC_TEXT 常量")
    if not wanted(text):
        raise SystemExit("！拼出来的结果仍不满足判据，放弃")

    if dry:
        print(f"\n--dry-run：将会写回 {len(text.splitlines())} 行（未改动）")
        return

    DAMAGED_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%H%M%S")
    keep = DAMAGED_DIR / f"main_window.py_damaged_{stamp}"
    shutil.copy2(LIVE, keep)
    LIVE.write_text(text, encoding="utf-8")
    print(f"\n已写回 {len(text.splitlines())} 行；受损版本存 {keep.relative_to(ROOT)}")
    print(f"摘要：{hashlib.md5(LIVE.read_bytes()).hexdigest()[:12]}")


if __name__ == "__main__":
    main()
