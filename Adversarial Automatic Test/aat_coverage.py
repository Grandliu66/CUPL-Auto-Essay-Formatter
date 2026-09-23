r"""覆盖率对账：**"程序里所有能动的字段" vs "扫描实际动过的字段"**，把漏掉的列出来。

为什么要它：单因素扫描"有没有漏掉某个按钮"不能靠人想。这里把 `PresetConfig`
的几个 dataclass 与 `pipeline.run()` 的参数**自动枚举**出来，再跟扫描器里每个开关
**实际改了什么**（用"读盘两份预设、跑一遍 mutate、逐字段比对"的办法自动发现，
不用手写元数据）对账。

三种"没覆盖"要分开说，因为它们性质完全不同：
  * **漏了**：这个字段存在、也能自动测，但扫描器没给它开关 → **我的活儿没干完**
  * **测不了**：这个按钮的效果不落在 docx 里（导入文件、备份、查阅窗口…）→ 扫描原理上管不着
  * **无穷**：结构页拖拽（把任意一段拖成任意身份）→ 组合无穷，不属于"单因素"

用法：.venv\Scripts\python.exe "Adversarial Automatic Test\aat_coverage.py"
"""

from __future__ import annotations

import dataclasses
import inspect
import pathlib
import re
import sys

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent
for entry in (str(ROOT), str(HERE)):
    if entry not in sys.path:
        sys.path.insert(0, entry)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from essay_tool.config.preset_io import load_preset          # noqa: E402
from essay_tool.config.schema import (CitationConfig,       # noqa: E402
                                      HeaderFooterRule, PageSetup,
                                      ParagraphStyleSpec, PresetConfig)
from essay_tool.engine import pipeline                        # noqa: E402

import aat_scan                                               # noqa: E402

PRESET = ROOT / "presets" / "中国政法大学-本科毕业论文.json"

# 只看这些复合字段的"里面"；它们本身不是可调项
COMPOSITE = {"page", "styles", "styles_by_name", "header_footer", "citation",
             "structure_required", "structure_insert_text"}

# 只递归这几种**配置类**的 dataclass。像 `Indent`/`Length` 这种"值对象"要当叶子
# 直接比——不然会多出 `first_line_indent.kind` 这种子路径，把"已覆盖"误报成"漏了"。
CONFIG_TYPES = (PresetConfig, PageSetup, ParagraphStyleSpec, HeaderFooterRule,
                CitationConfig)

# 这些"存在但故意不该动"——不算漏，列出来只是为了把话说清
NOT_APPLICABLE = {
    "styles[*].id": "样式编号，是身份不是配置",
    "styles[*].name": "样式名，是身份不是配置",
    "header_footer[*].section_name": "预设里只用于显示，不参与配对（2026-09-16 审计核实）",
    "param.check_source_alerts": "只影响告警，不影响 docx（已归入「测不了」）",
    "keep_run_formatting": "用户 2026-09-17 裁决：保留字符格式是代码层面的固定规则，"
                           "不对外调整（保留什么写在《工具操作逻辑》3.6 表 R18 的"
                           "「样式应用规则」列里）",
    "structure_insert_text[*]": "用户 2026-09-17 裁决：缺失时插入的文字不对外调整",
}

INDEXED = re.compile(r"^([A-Za-z_]+)\[[^\]]*\](.*)$")


def normalize(path: str) -> str:
    """`styles[S01].size` / `header_footer[0].header_text` → `xxx[*].yyy`。

    `styles_by_name` 是 `styles` 的按名镜像（预设内部的双份索引），改一处
    另一处必然跟着变，不单独算一项，直接折算到 `styles[*]`。
    """
    if path.startswith("styles_by_name["):
        path = "styles[" + path[len("styles_by_name["):]
    match = INDEXED.match(path)
    return f"{match.group(1)}[*]{match.group(2)}" if match else path


def walk_changes(before, after, prefix: str, out: set[str]) -> None:
    """找出 before/after 两个 dataclass 之间**哪些字段变了**（递归）。"""
    if isinstance(before, CONFIG_TYPES) and isinstance(after, CONFIG_TYPES):
        for spec in dataclasses.fields(before):
            walk_changes(getattr(before, spec.name), getattr(after, spec.name),
                         f"{prefix}.{spec.name}" if prefix else spec.name, out)
        return
    if isinstance(before, list) and isinstance(after, list):
        for index in range(min(len(before), len(after))):
            walk_changes(before[index], after[index], f"{prefix}[{index}]", out)
        return
    if isinstance(before, dict) and isinstance(after, dict):
        for key in sorted(set(before) | set(after)):
            walk_changes(before.get(key), after.get(key), f"{prefix}[{key}]", out)
        return
    if before != after:
        out.add(prefix)


def universe() -> tuple[dict[str, str], list[tuple[str, str]], list[tuple[str, str]]]:
    """返回 `(可自动测的字段, 原理上测不了的, 数量无穷的)`。"""
    auto: dict[str, str] = {}
    for name, kind in (("page", PageSetup), ("citation", CitationConfig)):
        for spec in dataclasses.fields(kind):
            auto[f"{name}.{spec.name}"] = f"预设「{name}」的 {spec.name}"
    for spec in dataclasses.fields(ParagraphStyleSpec):
        auto[f"styles[*].{spec.name}"] = f"格式页每一行的 {spec.name}"
    for spec in dataclasses.fields(HeaderFooterRule):
        auto[f"header_footer[*].{spec.name}"] = f"页眉页脚规则的 {spec.name}"
    for spec in dataclasses.fields(PresetConfig):
        if spec.name in COMPOSITE:
            continue
        auto[spec.name] = f"预设的 {spec.name}"
    auto["structure_required[*]"] = "结构页每个部件的启用勾选"
    auto["structure_insert_text[*]"] = "预设里「缺失时插入的文字」"
    for name, param in inspect.signature(pipeline.run).parameters.items():
        if name in ("source", "preset_path", "work_dir", "preset_config", "alerts"):
            continue
        auto[f"param.{name}"] = f"pipeline.run() 的参数 {name}"

    unreachable = [
        ("文件页·导入文件 / 备份开关 / 备份路径 / 自动清理", "效果落在文件系统，不进 docx"),
        ("文件页·保存/导入/导出/删除预设", "效果落在 %APPDATA% 的 JSON，不进 docx"),
        ("关于页·查阅规范 / 操作逻辑 / 另存官方版 / 联系开发者", "只开窗口，不进 docx"),
        ("空白模板导出", "入口不同（template.build_template），不是同一份成品"),
        ("日志区 / 崩溃弹窗 / 进度", "界面产物"),
        ("控件 → 参数 这一段（点一下按钮，参数真的变了吗）", "要 GUI 自动化，本扫描只测「参数 → 产物」"),
    ]
    infinite = [
        ("结构页拖拽：任意一段拖成任意身份（标题级别 / 部件 / 图题 / 表题 / 长引文）",
         "组合数无穷，不属于「单因素」；要另做等价类与抽样"),
        ("结构页多选拖拽、右键插空行、Ctrl+Z 撤回", "同上"),
        ("开关之间的组合（成对 / 成组）", "本扫描是单因素；成对组合只登记了已知耦合，还没跑"),
    ]
    return auto, unreachable, infinite


def main() -> int:
    auto, unreachable, infinite = universe()
    config = load_preset(PRESET)
    toggles = aat_scan.build_toggles(config, list(aat_scan.GROUPS), None, None)

    covered: dict[str, list[str]] = {}
    params_touched: set[str] = set()
    for toggle in toggles:
        before = load_preset(PRESET)          # 两份都从盘上读，保证起点一致
        after = load_preset(PRESET)
        extras: dict = {}
        toggle.apply(after, extras)
        found: set[str] = set()
        walk_changes(before, after, "", found)
        for name in found:
            covered.setdefault(normalize(name), []).append(toggle.name)
        params_touched |= {f"param.{key}" for key in extras}

    print("=" * 78)
    print("AAT 覆盖率对账")
    print("=" * 78)
    print(f"  可自动测的字段：{len(auto)} 个")
    print(f"  扫描器实际动到：{len(set(covered) | params_touched)} 个")

    missing = sorted(set(auto) - set(covered) - params_touched - set(NOT_APPLICABLE))
    print()
    print("=" * 78)
    print(f"① 漏了（存在、能测、但我没做开关）：{len(missing)} 个")
    print("=" * 78)
    for name in missing:
        print(f"  ✘ {name:<34} {auto[name]}")
    if not missing:
        print("  （无）")

    print()
    print("-" * 78)
    print("（下面这几项存在但故意不动，不算漏）")
    for name, why in NOT_APPLICABLE.items():
        print(f"  · {name}：{why}")

    print()
    print("=" * 78)
    print("② 测不了（效果不落在 docx 里）：")
    print("=" * 78)
    for name, why in unreachable:
        print(f"  – {name}\n      {why}")

    print()
    print("=" * 78)
    print("③ 数量无穷（不属于「单因素」）：")
    print("=" * 78)
    for name, why in infinite:
        print(f"  – {name}\n      {why}")

    extra = sorted((set(covered) | params_touched) - set(auto))
    if extra:
        print()
        print("（扫描器动了、但不在上面的枚举里："
              + "、".join(extra) + "）")
    return 1 if missing else 0


if __name__ == "__main__":
    raise SystemExit(main())
