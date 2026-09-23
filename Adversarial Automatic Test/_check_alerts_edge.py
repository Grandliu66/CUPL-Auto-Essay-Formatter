r"""告警层边界检查：**该报的报、不该报的别报**。

`aat_alerts.py` 那两轮验的是"表的文案与代码对不对得上""夹具上真报过哪些"。
这里补的是另外两件（都是行为，不能靠读代码确认）：

  ① **从没触发过的那条**：A9 `不存在的清单对象`（要"启用清单 + 文档里没有对应题注"）；
  ② **静默要求**——表里明写"不报"的情形，实际到底报没报：
       * 封面区间里的一切不参与识别 ⇒ 封面里的【摘要】**不该**触发"重复的部件"；
       * 手打目录条目退回正文，但正文里那个真标题只有一个 ⇒ A6“孤立”**应当**报；
       * A2 新表只认“超长”：`1.`＋61 字报、`一、`＋61 字不报、
         `万、`/`两、`/`廿一、`/带句号的整句都不报。

夹具由 `_make_alert_fixtures.py` 生成在 `reference\AAT夹具\`（子目录，不动
`reference\` 根目录的计数）。

用法：
    .venv\Scripts\python.exe "Adversarial Automatic Test\_check_alerts_edge.py"
"""
from __future__ import annotations

import pathlib
import sys

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from essay_tool.config.preset_io import load_preset               # noqa: E402
from essay_tool.engine import pipeline                            # noqa: E402
from essay_tool.engine.alerts import AlertCollector               # noqa: E402

FIXTURES = ROOT / "reference" / "AAT夹具"
PRESET = ROOT / "presets" / "中国政法大学-本科毕业论文.json"
OUT = HERE / "_out" / "alert_edge"

failures: list[str] = []
checked = 0


def check(label: str, ok: bool, detail: str = "") -> None:
    global checked
    checked += 1
    print(f"  {'✓' if ok else '✘'} {label}" + (f"    {detail}" if detail else ""))
    if not ok:
        failures.append(f"{label}    {detail}")


def run(name: str, *, cover_end: int | None = None, enable: tuple[str, ...] = ()):
    """跑一份夹具，返回 `(告警行列表, 收集器)`。"""
    preset = load_preset(PRESET)
    for part in enable:
        preset.structure_required[part] = True
    collector = AlertCollector()
    pipeline.run(FIXTURES / name, None, OUT / pathlib.Path(name).stem,
                 preset_config=preset, alerts=collector, cover_end=cover_end)
    collector.consolidate()
    return [alert.line for alert in collector.items], collector


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    if not FIXTURES.exists():
        print("夹具目录不存在，先跑 `_make_alert_fixtures.py`")
        return 1

    print("① A9 不存在的清单对象（启用插图清单 + 文档里没有图题）")
    lines, collector = run("无题注.docx", enable=("插图清单",))
    hits = [alert for alert in collector.items if alert.name == "不存在的清单对象"]
    check("报出了 A9", bool(hits), f"共 {len(hits)} 条")
    check("概要点名了缺哪种题注",
          any("启用了插图清单，但不存在图题" in alert.message for alert in hits),
          "；".join(alert.message for alert in hits) or "（一条都没有）")

    print("② 封面里的东西不参与识别 ⇒ 不该报“重复的部件”")
    lines, collector = run("封面含摘要.docx", cover_end=2)
    dupes = [alert for alert in collector.items if alert.name == "重复的部件"]
    check("设了封面范围时不报“重复的部件”", not dupes,
          "；".join(alert.line for alert in dupes) or "（干净）")
    lines, collector = run("封面含摘要.docx")          # 不设封面范围作对照
    dupes = [alert for alert in collector.items if alert.name == "重复的部件"]
    check("不设封面范围时**才**报（说明这条判据真的跟封面范围挂钩）",
          bool(dupes), "；".join(alert.line for alert in dupes) or "（竟然也没报）")

    print("③ 手打目录条目：退回正文；正文里那一个真标题该报“孤立”")
    lines, collector = run("手打目录带页码.docx")
    names = {alert.name for alert in collector.items}
    check("没有“疑似标题”", "疑似标题" not in names,
          "；".join(alert.line for alert in collector.items if alert.name == "疑似标题") or "（干净）")
    check("报出“孤立”（段 5 那个唯一的一级标题）",
          any("孤立" in alert.name and "（5" in alert.line for alert in collector.items),
          "；".join(alert.line for alert in collector.items if alert.name == "孤立") or "（没报）")
    check("但**报告**了“多余的正文”（域外手打条目，与表 A8 一致）",
          "多余的正文" in names)

    print("④ 新表 A2 只认超长：1.＋61字报、一、＋61字不报、廿一/万/两/句号都不报")
    lines, collector = run("编号边界.docx")
    suspect = [alert for alert in collector.items if alert.name == "疑似标题"]
    summary = "；".join(alert.line for alert in suspect)
    check("“疑似标题”只报了 1 条", len(suspect) == 1, summary or "（一条都没有）")
    check("报的是「1. ＋六十一字」那条", "1. " in summary, summary)
    check("「一、＋六十一字」没被报（中文一级不受长度限制）", "一、" not in summary, summary)
    for word in ("廿一", "万、", "两、", "这一句"):
        check(f"「{word}」没被报", word not in summary, summary)

    print()
    print(f"共 {checked} 条断言，失败 {len(failures)} 条")
    for item in failures:
        print(f"  ✘ {item}")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
