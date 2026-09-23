r"""AAT 跑批入口：一条命令跑完所有层，**红了返回非 0 退出码**。

为什么需要：之前每层都要手敲一条命令，而且只有人眼看结论——CI 或回归批里挂不
进去（对面早就指出过 `stress_tzst.py` 不设退出码的问题）。

七层与各自的判红口径：
  1. `gen_matrix.py`     写法层（1229）：出现 `✘分歧` 或 `？未判定` 就红
  2. `gen_structure.py`  结构层（251）：出现 `✘分歧` 就红
  3. `gen_pipeline.py`   产物层（620）：出现 `✘`（不收敛/崩溃/坏产物/搬运失败）就红，
                         只有 `⚠️` 算黄
  4. `gen_invariant.py`  产物不变量（20 份）：成品里"源里认得出的部件"一个不能少、
                         先后不能乱；`问题条数 > 0` 就红（数量差异只报 ⚠️）
  5. `_rerun_all.py`     期望对拍（2026-09-19 起 25 份夹具、~10 秒）：有夹具出现“有差异”就红
  6. `aat_alerts.py`     告警对拍（文案 + 触发覆盖）：`问题条数 > 0` 就红
  7. `_check_alerts_edge.py` 告警边界（该报的报、不该报的别报）：`失败 > 0` 就红
  8. `aat_asserts.py`    断言层（R26 附录序号题名、清单域插入计划；2026-09-20 补：
                         标签前另起一页覆盖、清除分页符/空行、保留字符格式含颜色）：
                         `失败 > 0` 就红
  9. `gen_exceptions.py` 例外层（3.6 表「允许的例外」机械切分 → 168 条样本 × 对拍）：
                         与写法层同口径（`✘分歧`/`？未判定` 就红）
 10. `gen_equiv.py`      等价层（写法等价对，产物必须逐字节一致）：`问题条数 > 0` 就红

另有一条**不进本跑批**的慢线（默认 4 进程，约 6 分钟；单进程 20 分钟）：
`scan_all.py` —— 单因素扫描 × 全部夹具，产出"每个开关由哪些夹具触发"登记表。

用法：
    .venv\Scripts\python.exe "Adversarial Automatic Test\run_aat.py"
    .venv\Scripts\python.exe "Adversarial Automatic Test\run_aat.py" --quick   # 跳过产物层（最慢的 90 秒）
    .venv\Scripts\python.exe "Adversarial Automatic Test\run_aat.py" --only 告警

退出码：0 = 全绿（可以有黄）；1 = 有红。
"""
from __future__ import annotations

import pathlib
import re
import subprocess
import sys
import time

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent
OUT = HERE / "_out" / "batch"
JOBS = "4"


def _count(pattern: str, text: str) -> int:
    """`pattern` 里第 1 个捕获组是数字，求和（用于逐夹具的点数）。"""
    total = 0
    for match in re.finditer(pattern, text):
        total += int(match.group(1))
    return total


def _number(pattern: str, text: str) -> int:
    """同一个数字会出现在两处（汇总行 + “按组/按夹具”行）——取**最大值**，
    累加会把同一条分歧数成两遍（2026-09-17 实测：251 条里 1 条分歧被报成 2）。"""
    values = [int(match.group(1)) for match in re.finditer(pattern, text)]
    return max(values, default=0)


def _summary_number(pattern: str, text: str) -> int:
    """只从**汇总行**里取数字（“跑完 N 条：✓…、✘…”那一行）。

    2026-09-19 实拍伪影：产物层的**个案行**也带“✘不收敛（每遍都变）”，而有一类
    变体叫「首段拉到 5000 字」⇒ 旧 `_number` 取 max 后，真账 142 被 5000 顶掉。
    判据里的数字只该来自汇总行；万一将来汇总行认不出来了，退回全篇取 max
    （保持旧行为，不静默变绿）。
    """
    summary = "\n".join(line for line in text.splitlines()
                         if "跑完" in line or "✓一致" in line or "✓正常" in line)
    if not summary:
        return _number(pattern, text)
    return _number(pattern, summary)


def judge_matrix(text: str) -> tuple[list[str], list[str]]:
    red, yellow = [], []
    for label, pattern in (("写法分歧", r"✘分歧[ \t]+(\d+)"), ("未判定", r"？未判定[ \t]+(\d+)")):
        number = _summary_number(pattern, text)
        if number:
            red.append(f"{label} {number}")
    return red, yellow


def judge_structure(text: str) -> tuple[list[str], list[str]]:
    number = _summary_number(r"✘分歧[ \t]+(\d+)", text)
    return ([f"结构分歧 {number}"] if number else []), []


def judge_pipeline(text: str) -> tuple[list[str], list[str]]:
    red, yellow = [], []
    for label, pattern in (("不收敛", r"✘不收敛[^\d\n]*(\d+)"),
                           ("崩溃", r"✘崩溃[^\d\n]*(\d+)"),
                           ("坏产物", r"✘坏产物[^\d\n]*(\d+)"),
                           ("搬运失败", r"✘搬运失败[^\d\n]*(\d+)"),
                           ("夹具缺失", r"✘夹具缺失[^\d\n]*(\d+)")):
        number = _summary_number(pattern, text)
        if number:
            red.append(f"{label} {number}")
    if "⚠️" in text:
        yellow.append("有 ⚠️ 项")
    return red, yellow


def judge_rerun(text: str) -> tuple[list[str], list[str]]:
    """每个夹具一行结论。

    两层坑都踩过（2026-09-19）：旧版只认"结论：…"这一行，而"有差异"时那行
    叫"差异概览：…" ⇒ **有差异反而判不出来**（假绿）。所以再加一道：一个
    夹具都没读到也判红，免得输出格式再变时又静默全绿。
    """
    red = []
    pattern = r"=== (.+?)\n\s*(差异概览|结论)：([^\n]*)"
    if not re.findall(pattern, text):
        return ["没有读到任何夹具的结论（输出格式可能变了）"], []
    diffs: list[str] = []
    failed: list[str] = []
    for fixture, kind, detail in re.findall(pattern, text):
        detail = detail.strip()
        if kind == "结论" and "跑失败" in detail:
            failed.append(fixture)
        elif kind == "差异概览" and "没有差异" not in detail:
            diffs.append(fixture)
    # 夹具多了（2026-09-19 起 25 份）——逐步列名字会把汇总表撑爆，改成聚合。
    for label, items in (("有差异", diffs), ("跑失败", failed)):
        if items:
            head = "、".join(items[:6]) + ("…" if len(items) > 6 else "")
            red.append(f"{label} {len(items)} 份（{head}）")
    numbers = _count(r"未判定：(\d+) 段", text)
    if numbers:
        red.append(f"未判定 {numbers} 段")
    return red, []


def judge_generic(label: str, text: str) -> tuple[list[str], list[str]]:
    number = _number(r"问题条数：(\d+)", text) or _number(r"失败\s+(\d+)\s+条", text)
    return ([f"{label} {number}"] if number else []), []


def _crash_reason(text: str) -> str | None:
    """脚本自己没跑完（抛了异常）——必须红。

    2026-09-19 实拍：期望对拍那层抛异常退出，而它的判据只认“有差异”这个
    字眼，于是整层显示“全绿”——比红更危险（谁也不知道它根本没跑）。
    """
    if "Traceback (most recent call last)" not in text:
        return None
    tail = [line.strip() for line in text.strip().splitlines() if line.strip()]
    reason = tail[-1][:70] if tail else "?"
    return f"脚本崩溃：{reason}"


LAYERS: list[tuple[str, str, list[str], object]] = [
    ("写法层", "gen_matrix.py", ["--jobs", JOBS], judge_matrix),
    ("结构层", "gen_structure.py", ["--jobs", JOBS], judge_structure),
    ("产物层", "gen_pipeline.py", ["--all", "--jobs", JOBS, "--tag", "aat"], judge_pipeline),
    ("产物不变量", "gen_invariant.py", ["--jobs", JOBS],
     lambda text: judge_generic("问题", text)),
    ("期望对拍", "_rerun_all.py", [], judge_rerun),
    ("告警对拍", "aat_alerts.py", [], lambda text: judge_generic("问题", text)),
    ("告警边界", "_check_alerts_edge.py", [], lambda text: judge_generic("失败", text)),
    ("断言层", "aat_asserts.py", [], lambda text: judge_generic("断言失败", text)),
    ("例外层", "gen_exceptions.py", ["--jobs", JOBS], judge_matrix),
    ("等价层", "gen_equiv.py", [], lambda text: judge_generic("等价问题", text)),
]


def main(argv: list[str]) -> int:
    quick = "--quick" in argv
    only = argv[argv.index("--only") + 1] if "--only" in argv else None
    OUT.mkdir(parents=True, exist_ok=True)

    results: list[tuple[str, float, list[str], list[str]]] = []
    for name, script, args, judge in LAYERS:
        if only and only not in name:
            continue
        if quick and script == "gen_pipeline.py":
            print(f"—— {name}：--quick，跳过")
            continue
        print(f"—— {name}（{script}）跑起来了…")
        started = time.perf_counter()
        finished = subprocess.run(
            [sys.executable, str(HERE / script), *args],
            cwd=str(ROOT), capture_output=True, text=True,
            encoding="utf-8", errors="replace",
        )
        elapsed = time.perf_counter() - started
        text = finished.stdout + finished.stderr
        (OUT / f"{script.replace('.py', '')}.txt").write_text(
            text, encoding="utf-8")
        crash = _crash_reason(text)
        red, yellow = ([crash], []) if crash else judge(text)
        results.append((name, elapsed, red, yellow))

    print()
    print("=" * 62)
    total_red = 0
    for name, elapsed, red, yellow in results:
        mark = "✘" if red else ("⚠️" if yellow else "✓")
        detail = "；".join(red + yellow) if (red or yellow) else "全绿"
        print(f"  {mark} {name:<8} {elapsed:6.1f}s   {detail}")
        total_red += len(red)
    print("=" * 62)
    print(f"共 {len(results)} 层，红 {total_red} 条"
          + ("（明细在各层自己的 _out/*.txt）" if total_red else ""))
    return 1 if total_red else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
