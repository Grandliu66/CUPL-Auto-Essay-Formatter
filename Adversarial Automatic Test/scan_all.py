# -*- coding: utf-8 -*-
r"""多夹具批跑：把单因素扫描跑遍夹具清单，汇总成"每个开关由哪些夹具触发"。

README「下一步」第 6 项。解决的事：一份夹具的扫描里，"完全没变化"大多是
**这份夹具触发不到**（民总没有图题/表题/附录/长引文/英文部件）——把同一个
开关在十几份夹具上各跑一遍，"没变化"就能压下去，同时产出一张登记表：
"这个开关要哪份文档才看得出效果"。

用法：
    .venv\Scripts\python.exe "Adversarial Automatic Test\scan_all.py"
    .venv\Scripts\python.exe "Adversarial Automatic Test\scan_all.py" --limit 3
    .venv\Scripts\python.exe "Adversarial Automatic Test\scan_all.py" --groups 结构
    .venv\Scripts\python.exe "Adversarial Automatic Test\scan_all.py" --only 图表
    .venv\Scripts\python.exe "Adversarial Automatic Test\scan_all.py" --jobs 8

多核：`--jobs N` 按**夹具**分进程（默认 4，故意不吃满——本机 16 核 32 线程，
用户定过"只要不吃满所有就可以"）。每份夹具的成本 ≈ 450 个开关 × 0.1~0.2 秒
（一次管线 0.1~0.2s），单进程全量 ≈ 20 分钟、4 进程 ≈ 5~6 分钟（2026-09-19 实测）。

迭代时用 `--limit/--only/--groups/--jobs` 收窄；产物在 `_out/scan-all/`
（每夹具一份报告 + 一份汇总登记表）；中间产物按进程隔离在 `_out/scan-work/p<pid>/`。
"""
from __future__ import annotations

import concurrent.futures
import os
import pathlib
import sys
import time

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent
for entry in (str(ROOT), str(HERE)):
    if entry not in sys.path:
        sys.path.insert(0, entry)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import aat_scan                                                 # noqa: E402

PRESET = ROOT / "presets" / "中国政法大学-本科毕业论文.json"
OUT = HERE / "_out" / "scan-all"
WORK = HERE / "_out" / "scan-work"
FIXTURE_DIRS = (ROOT / "reference", ROOT / "reference" / "用户测试稿")
DEFAULT_JOBS = 4


def fixtures(only: str | None) -> list[pathlib.Path]:
    found: list[pathlib.Path] = []
    seen: set[str] = set()
    for folder in FIXTURE_DIRS:
        if not folder.is_dir():
            continue
        for path in sorted(folder.glob("*.docx")):
            # 跨目录重名副本（用户测试稿里那几个是给 tools/ 的副本）只取一份：
            # 报告按 stem 写盘，两份会互相覆盖（2026-09-19 首轮全量实测
            # TOC.docx 出现两次、报告只留最后一份）。
            if path.stem in seen:
                continue
            seen.add(path.stem)
            found.append(path)
    if only:
        found = [path for path in found if only in path.name]
    return found


def run_fixture(payload: tuple[str, list[str] | None]) -> dict:
    """一个进程跑一份夹具：扫描 → 报告文本 + 登记行（都是普通对象，可跨进程）。

    ⚠️ 会被扔进进程池：返回值里**不能有 `Toggle`/`CaseOutcome` 这些对象**
    （`Toggle.mutate` 是函数，不能 pickle）——只取字符串。
    中间产物按进程号隔离：`scan()` 开头会清目录，共用会互删现场。
    """
    source_name, groups = payload
    source = pathlib.Path(source_name)
    try:
        outcomes = aat_scan.scan(source, PRESET, groups=groups,
                                 out_root=WORK / f"p{os.getpid()}")
    except Exception as error:  # noqa: BLE001 —— 整份夹具失败也走同一张表
        return {"name": source.name, "error": f"整份夹具跑失败：{error}"}
    report = aat_scan.render(outcomes, source, PRESET)
    rows: list[dict] = []
    tally = {"in-scope": 0, "no-effect": 0, "escaped": 0, "skip": 0}
    for outcome in outcomes:
        verdict, escaped, _coupled = outcome.verdict()
        tally[verdict] += 1
        rows.append({"group": outcome.toggle.group, "name": outcome.toggle.name,
                     "verdict": verdict, "description": outcome.description,
                     "escaped": "；".join(escaped)})
    return {"name": source.name, "report": report, "rows": rows, "tally": tally}


def main(argv: list[str]) -> int:
    limit = int(argv[argv.index("--limit") + 1]) if "--limit" in argv else None
    only = argv[argv.index("--only") + 1] if "--only" in argv else None
    groups = (argv[argv.index("--groups") + 1].split(",")
              if "--groups" in argv else None)
    jobs = int(argv[argv.index("--jobs") + 1]) if "--jobs" in argv else DEFAULT_JOBS
    sources = fixtures(only)
    if limit is not None:
        sources = sources[:limit]
    if not sources:
        print("没有找到夹具（检查 reference/ 与 reference/用户测试稿/）")
        return 1
    jobs = max(1, min(jobs, len(sources)))
    OUT.mkdir(parents=True, exist_ok=True)
    print(f"批跑 {len(sources)} 份夹具（{jobs} 进程），"
          f"分组＝{'全部' if not groups else '/'.join(groups)}", flush=True)

    # (分组, 开关名) → {"hit": [夹具…], "no-effect": […], "escaped": […], "skip": [….]}
    register: dict[tuple[str, str], dict[str, list[str]]] = {}
    escaped_total = 0
    failed: list[str] = []
    started_all = time.perf_counter()

    def handle(result: dict, done: int) -> None:
        """收一份夹具的结果（主进程里跑，进程池只负责跑扫描）。"""
        nonlocal escaped_total
        name = result["name"]
        print(f"\n===== [{done}/{len(sources)}] {name} =====", flush=True)
        if result.get("error"):
            print(f"  ✘ {result['error']}", flush=True)
            failed.append(name)
            escaped_total += 1
            return
        (OUT / f"{pathlib.Path(name).stem}.txt").write_text(
            result["report"], encoding="utf-8")
        tally = result["tally"]
        for row in result["rows"]:
            key = (row["group"], row["name"])
            entry = register.setdefault(
                key, {"hit": [], "no-effect": [], "escaped": [], "skip": []})
            entry["hit" if row["verdict"] == "in-scope" else row["verdict"]].append(name)
            if row["verdict"] == "escaped":
                escaped_total += 1
                print(f"  ✘ 越界：{row['description']} —— {row['escaped']}", flush=True)
        print(f"  有效 {tally['in-scope']}／没变化 {tally['no-effect']}／"
              f"越界 {tally['escaped']}／跳过 {tally['skip']}", flush=True)

    payloads = [(str(source), groups) for source in sources]
    if jobs > 1:
        with concurrent.futures.ProcessPoolExecutor(max_workers=jobs) as pool:
            futures = [pool.submit(run_fixture, payload) for payload in payloads]
            for done, future in enumerate(
                    concurrent.futures.as_completed(futures), start=1):
                handle(future.result(), done)
    else:
        for done, payload in enumerate(payloads, start=1):
            handle(run_fixture(payload), done)

    # ---------------- 汇总登记表 ----------------
    lines = ["=" * 78,
             "AAT 多夹具批跑汇总：每个开关由哪些夹具触发",
             f"夹具数：{len(sources)}；分组：{'全部' if not groups else '/'.join(groups)}",
             "=" * 78, ""]
    never: list[str] = []
    escaped_items: list[str] = []
    for (group, name), entry in sorted(register.items()):
        head = f"[{group}] {name}"
        if entry["escaped"]:
            # 越界要**独立列一节**：有命中又有越界的开关（如"正文页码"）如果
            # 只挂进 ✓ 分支，越界就被命中名单盖掉了（2026-09-19 实测：汇总
            # 只报了总数 7，点名的 0 个）。两边都列，互不遮盖。
            escaped_items.append(
                f"  ✘ {head} —— 越界夹具：{'、'.join(sorted(set(entry['escaped'])))}")
        if entry["hit"]:
            lines.append(f"✓ {head}")
            lines.append(f"    触发：{'、'.join(sorted(set(entry['hit'])))}")
        elif not entry["escaped"]:
            never.append(head)
    if escaped_items:
        lines.append("")
        lines.append(f"———— 有越界的开关（{len(escaped_items)} 个）————")
        lines.extend(escaped_items)
    if never:
        lines.append("")
        lines.append(f"———— 没有夹具触发它（{len(never)} 个）————")
        lines.extend(f"  ⚠️ {item}" for item in never)

    total_elapsed = time.perf_counter() - started_all
    lines.append("")
    lines.append("=" * 78)
    lines.append(f"开关总数 {len(register)}；无夹具触发 {len(never)}；越界 {escaped_total}")
    lines.append(f"总耗时 {total_elapsed:.0f} 秒")
    summary = OUT / "汇总.txt"
    summary.write_text("\n".join(lines), encoding="utf-8")
    print()
    print(f"开关总数 {len(register)}；无夹具触发 {len(never)}；越界 {escaped_total}"
          f"；总耗时 {total_elapsed:.0f} 秒")
    print(f"汇总已写出：{summary.relative_to(ROOT)}")
    return 1 if escaped_total else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
