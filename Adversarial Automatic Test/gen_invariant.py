r"""第七层：产物不变量 —— **成品里"该有的部件"一个都不能少**。

为什么单独一层：产物层只问"稳不稳"（崩没崩、四遍一不一样），期望对拍只问
"两份产物之间的差"——**两边一致地错它看不见**。"某个部件在成品里整个消失"
（封面被吃掉、参考文献整块不见、摘要只剩标签）在现有六层里**一层都抓不到**，
这层就是补它。

判据（用同一把尺子量两头：就是识别器本身）：

1. **部件一个不少**：源文档里认出来的每种部件，在成品里也认得出来；
2. **先后不变**：这些部件在源里的出现顺序，在成品里仍然保持——允许中间
   插入新生成的部件（缺件补的占位语），但相对次序不许乱；
3. 数量对不上的只报 ⚠️（先看数据再收紧；成品的部件计数本来就可能因为
   "标签+内容"两段式而变化）。

为什么要"回炉"：用户真实场景就是"导出 → 再拿产物导出一次"。成品里认不出
源里有的部件，既说明成品丢了东西，也说明**工具认不出自己写出来的东西**
（历史上"占位语下一遍认不出、每跑一遍多一条"就是这一类）。

用法：
    .venv\Scripts\python.exe "Adversarial Automatic Test\gen_invariant.py"
    .venv\Scripts\python.exe "Adversarial Automatic Test\gen_invariant.py" --only 民总
"""
from __future__ import annotations

import argparse
import concurrent.futures
import pathlib
import shutil
import sys
from collections import Counter

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent
for entry in (str(ROOT), str(HERE)):
    if entry not in sys.path:
        sys.path.insert(0, entry)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from gen_pipeline import FIXTURES, PRESET, _fixture_path         # noqa: E402

from essay_tool.config.preset_io import load_preset              # noqa: E402
from essay_tool.engine import pipeline                           # noqa: E402
from essay_tool.engine.alerts import AlertCollector              # noqa: E402

OUT = HERE / "_out" / "gen_invariant"
DEFAULT_JOBS = 4


def _parts_of(path) -> list[tuple[int, str]]:
    """一份文档认出来的部件：`[(段号, 部件名)]`，按段号排序（没有身份的不列）。

    拆内容控件的口径跟导出内部保持一致——口径不一致会让两边的段号整体错位。
    """
    from docx import Document

    from essay_tool.engine.cleanup import unwrap_section_break_sdt
    from essay_tool.engine.doc_index import body_toc_field_parts, paragraph_ordinals
    from essay_tool.engine.recognize import recognize

    document = Document(str(path))
    unwrap_section_break_sdt(document, keep_content=bool(body_toc_field_parts(document)))
    index = recognize(document)
    ordinals = paragraph_ordinals(index)
    return [(number, entry.part) for number, entry in sorted(ordinals.items())
            if entry.part]


def _squash(names: list[str]) -> list[str]:
    """相邻重名只算一个（"参考文献、参考文献、参考文献"是一个块，不是三次）。"""
    out: list[str] = []
    for name in names:
        if not out or out[-1] != name:
            out.append(name)
    return out


def _is_subsequence(needle: list[str], haystack: list[str]) -> bool:
    """needle 是不是 haystack 的子序列（成品里允许插进新生成的部件）。"""
    it = iter(haystack)
    return all(any(item == want for item in it) for want in needle)


def run_case(name: str, out_root: str) -> tuple[str, str, str]:
    """返回 `(夹具名, 结论, 明细)`。会被扔进进程池，参数与返回值都可序列化。"""
    fixture = _fixture_path(name)
    if fixture is None:
        return name, "✘夹具缺失", ""
    try:
        source = _parts_of(fixture)
    except Exception as exc:  # noqa: BLE001 —— 报出来比抛出去有用；带上出错那一行
        import traceback

        where = traceback.format_exc().strip().splitlines()[-2].strip()
        return name, "✘源文档识别失败", f"{type(exc).__name__}: {exc} @ {where}"[:130]

    work = pathlib.Path(out_root) / name.replace(".docx", "")
    shutil.rmtree(work, ignore_errors=True)
    try:
        product = pipeline.run(fixture, None, work,
                               preset_config=load_preset(PRESET),
                               alerts=AlertCollector())
    except Exception as exc:  # noqa: BLE001
        return name, "✘导出失败", f"{type(exc).__name__}: {exc}"[:90]
    try:
        made = _parts_of(product)
    except Exception as exc:  # noqa: BLE001
        return name, "✘成品识别失败", f"{type(exc).__name__}: {exc}"[:90]

    want = _squash([part for _number, part in source])
    got = _squash([part for _number, part in made])
    problems: list[str] = []
    notes: list[str] = []

    missing = sorted(set(want) - set(got))
    if missing:
        problems.append("成品里认不出这些部件：" + "、".join(missing))
    elif not _is_subsequence(want, got):
        problems.append(f"部件先后乱了（源 {len(want)} 个 → 成品 {len(got)} 个）")

    want_n, got_n = Counter(want), Counter(got)
    extra = sorted(part for part, count in got_n.items() if count > want_n.get(part, 0))
    if extra:
        notes.append("成品里多出来：" + "、".join(extra))

    return name, ("✓正常" if not problems else "✘问题"), "；".join(problems + notes)


def _run_payload(payload: tuple[str, str]) -> tuple[str, str, str]:
    """进程池入口：`Executor.map` 只吃单参数的函数。"""
    return run_case(*payload)


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="第七层：产物不变量（成品里该有的部件一个不少）")
    parser.add_argument("--jobs", type=int, default=DEFAULT_JOBS)
    parser.add_argument("--only", default=None, help="只跑名字里含这个字样的夹具")
    args = parser.parse_args(argv)

    names = [name for name in FIXTURES if not args.only or args.only in name]
    out_root = OUT if not args.only else OUT / "only"
    out_root.mkdir(parents=True, exist_ok=True)
    for stale in list(out_root.glob("*.docx")):
        stale.unlink()
    for stale in list(out_root.glob("~$*")):
        stale.unlink()

    payloads = [(name, str(out_root)) for name in names]
    if args.jobs > 1:
        with concurrent.futures.ProcessPoolExecutor(max_workers=args.jobs) as pool:
            outcomes = list(pool.map(_run_payload, payloads))
    else:
        outcomes = [run_case(*item) for item in payloads]

    ok, bad = 0, 0
    problems: list[str] = []
    notes: list[str] = []
    for name, verdict, detail in outcomes:
        if verdict == "✓正常":
            ok += 1
        else:
            bad += 1
        line = f"  {verdict}  {name}" + (f"\n        {detail}" if detail else "")
        (problems if verdict != "✓正常" else notes).append(line)

    print(f"跑完 {len(outcomes)} 份（{args.jobs} 进程）：✓正常 {ok}、✘问题 {bad}")
    for title, items in (("问题", problems), ("备注（不算红）", notes)):
        if not items:
            continue
        print()
        print("=" * 74)
        print(f"{title}：{len(items)} 条")
        print("=" * 74)
        for item in items:
            print(item)

    report = out_root / "gen_invariant.txt"
    report.write_text("\n".join(problems + notes) + f"\n\n问题条数：{bad}\n",
                      encoding="utf-8")
    print()
    print(f"问题条数：{bad}")
    print(f"明细已写出：{report.relative_to(ROOT)}")
    return 1 if bad else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
