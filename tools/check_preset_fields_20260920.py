"""预设新增字段的往返检查：清除手打分页符/清除空行/标签前另起一页（JSON、Excel 都要能往返，旧预设要照常读）。

跑法：.venv\\Scripts\\python.exe tools\\check_preset_fields_20260920.py    （退出码 0＝全部通过）
"""

import sys, pathlib, tempfile, dataclasses
sys.stdout.reconfigure(encoding="utf-8")
ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from essay_tool.config.preset_io import load_preset, save_preset_json, export_to_excel
bad = 0
def check(n, c, extra=""):
    global bad
    print(("OK   " if c else "FAIL ") + n, "" if c else extra); bad += (not c)

T = pathlib.Path(tempfile.mkdtemp())
# 1 旧预设（没有新字段）照常读，默认值＝老行为
for p in sorted((ROOT / "presets").glob("*.json")):
    c = load_preset(p)
    check(f"旧 JSON 预设 {p.stem}：新字段取默认", (c.clear_page_breaks, c.clear_blank_lines, c.label_page_break) == (False, False, {}))
old_xlsx = ROOT / "presets" / "中国政法大学-本科毕业论文.xlsx"
c = load_preset(old_xlsx)
check("旧 Excel 预设（没有新行/新列）照常读", (c.clear_page_breaks, c.clear_blank_lines, c.label_page_break) == (False, False, {}))

# 2 JSON 往返
base = load_preset(ROOT / "presets" / "中国政法大学-本科毕业论文.json")
cfg = dataclasses.replace(base, clear_page_breaks=True, clear_blank_lines=True,
                          label_page_break={"摘要": True, "参考文献": False, "引言": True},
                          keep_run_formatting=frozenset({"b", "i", "u", "vertAlign", "color"}))
j = T / "x.json"; save_preset_json(cfg, j); back = load_preset(j)
check("JSON 往返：清除分页符/空行", (back.clear_page_breaks, back.clear_blank_lines) == (True, True))
check("JSON 往返：标签另起一页覆盖", back.label_page_break == {"摘要": True, "参考文献": False, "引言": True}, back.label_page_break)
check("JSON 往返：保留颜色", back.keep_run_formatting == frozenset({"b", "i", "u", "vertAlign", "color"}))

# 3 Excel 往返
x = T / "x.xlsx"; export_to_excel(cfg, x); back = load_preset(x)
check("Excel 往返：清除分页符/空行", (back.clear_page_breaks, back.clear_blank_lines) == (True, True), (back.clear_page_breaks, back.clear_blank_lines))
check("Excel 往返：标签另起一页覆盖（是/否/空＝跟随）", back.label_page_break == {"摘要": True, "参考文献": False, "引言": True}, back.label_page_break)
check("Excel 往返：保留颜色", "color" in (back.keep_run_formatting or ()), back.keep_run_formatting)
# 4 默认（全关）也能往返
x2 = T / "d.xlsx"; export_to_excel(base, x2); b2 = load_preset(x2)
check("Excel 往返：默认值（全关、无覆盖）", (b2.clear_page_breaks, b2.clear_blank_lines, b2.label_page_break) == (False, False, {}))
print("失败:", bad)

sys.exit(1 if bad else 0)
