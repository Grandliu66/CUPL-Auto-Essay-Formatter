# -*- coding: utf-8 -*-
r"""等价类（写法等价）：同一句话的**不同合法写法**，成品必须完全一致。

期望值是"两边一样"，**不需要判据**（README「下一步」第 4 项）。做法：
  1. 造两份只差"写法"的合成文档（同一份上下文）；
  2. 各跑一遍管线；
  3. 拆包逐字节对比（除 `docProps/core.xml` 的时间戳——它是唯一不确定的部件，
     2026-09-17 实测）；
  4. 不一致就报差异（用 aat_diff 给段落级现场）。

用法：
    .venv\Scripts\python.exe "Adversarial Automatic Test\gen_equiv.py"
    .venv\Scripts\python.exe "Adversarial Automatic Test\gen_equiv.py" --only 括号
"""
from __future__ import annotations

import hashlib
import pathlib
import shutil
import sys
import time
import zipfile

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent
for entry in (str(ROOT), str(HERE)):
    if entry not in sys.path:
        sys.path.insert(0, entry)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import aat_diff                                                 # noqa: E402
from aat_docx import read_docx                                  # noqa: E402
from docx import Document                                       # noqa: E402
from essay_tool.config.preset_io import load_preset             # noqa: E402
from essay_tool.engine import pipeline                          # noqa: E402
from essay_tool.engine.alerts import AlertCollector             # noqa: E402

PRESET = ROOT / "presets" / "中国政法大学-本科毕业论文.json"
WORK = HERE / "_out" / "gen_equiv"
REPORT = HERE / "_out" / "gen-equiv.txt"
# 最小上下文（同 gen_matrix）：题目段挡掉"第一段＝题目"的歧义。
CONTEXT = ["论测试文档里标题的写法", ""]

# (名字, 左写法, 右写法)——两边的上下文相同，只有"写法"不同。
CASES: list[tuple[str, list[str], list[str]]] = [
    ("编号·顿号 vs 半角点", ["1、测试标题", "段落内容。"], ["1. 测试标题", "段落内容。"]),
    ("编号·半角点 vs 全角点", ["1. 测试标题", "段落内容。"], ["1．测试标题", "段落内容。"]),
    ("编号·二级全角括号 vs 半角括号", ["（一）测试标题", "段落内容。"], ["(一)测试标题", "段落内容。"]),
    ("标签·有【】 vs 无【】", ["【摘要】", "摘要内容。"], ["摘要", "摘要内容。"]),
    ("标签·半角括号 vs 全角括号", ["[摘要]", "摘要内容。"], ["【摘要】", "摘要内容。"]),
    ("标签·大小写（ENGLISH）", ["【Abstract】", "Summary text."], ["【ABSTRACT】", "Summary text."]),
]

# 拆包对比时排除的部件（时间戳）。
SKIP_PARTS = {"docProps/core.xml"}


def package_digest(path: pathlib.Path) -> dict[str, str]:
    """整包逐部件 sha256（除 core.xml）。"""
    with zipfile.ZipFile(path) as package:
        return {name: hashlib.sha256(package.read(name)).hexdigest()
                for name in package.namelist() if name not in SKIP_PARTS}


def build_doc(path: pathlib.Path, body: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    document = Document()
    for line in CONTEXT + body:
        document.add_paragraph(line)
    document.save(str(path))


def run_one(source: pathlib.Path, target: pathlib.Path) -> pathlib.Path:
    target.mkdir(parents=True, exist_ok=True)
    return pipeline.run(source, None, target,
                        preset_config=load_preset(PRESET),
                        alerts=AlertCollector())


def main(argv: list[str]) -> int:
    only = argv[argv.index("--only") + 1] if "--only" in argv else None
    shutil.rmtree(WORK, ignore_errors=True)
    started = time.perf_counter()
    problems: list[str] = []
    lines: list[str] = []
    ran = 0
    for name, left, right in CASES:
        if only and only not in name:
            continue
        ran += 1
        left_doc = WORK / "case" / "left.docx"
        right_doc = WORK / "case" / "right.docx"
        build_doc(left_doc, left)
        build_doc(right_doc, right)
        left_out = run_one(left_doc, WORK / "case" / "out_left")
        right_out = run_one(right_doc, WORK / "case" / "out_right")
        digest_left = package_digest(left_out)
        digest_right = package_digest(right_out)
        if digest_left == digest_right:
            print(f"  ✓一致  {name}")
            continue
        differing = sorted(set(digest_left) | set(digest_right))
        differing = [part for part in differing
                     if digest_left.get(part) != digest_right.get(part)]
        print(f"  ✘不一致  {name}  差异部件：{'、'.join(differing[:6])}")
        detail = [f"✘ {name}", f"  差异部件：{'、'.join(differing)}"]
        try:
            diff = aat_diff.compare_shapes(read_docx(left_out), read_docx(right_out))
            report = aat_diff.render(diff, max_items=6)
            detail.extend("  " + line for line in report.splitlines()[:18])
        except Exception as error:                              # noqa: BLE001
            detail.append(f"  （段落级差分也失败：{error}）")
        detail.append(f"  左文：{left}")
        detail.append(f"  右文：{right}")
        problems.append("\n".join(detail))
        lines.extend(detail)

    print()
    print(f"跑完 {ran} 组等价对，耗时 {time.perf_counter() - started:.1f}s"
          f"，问题条数：{len(problems)}")
    if problems:
        print()
        print("=" * 74)
        for item in problems:
            print(item)
            print("-" * 74)
    REPORT.write_text("\n".join(lines), encoding="utf-8")
    shutil.rmtree(WORK, ignore_errors=True)
    return 1 if problems else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
