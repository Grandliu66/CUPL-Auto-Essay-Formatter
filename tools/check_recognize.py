"""识别层重写（engine/recognize.py）的初步人工核对工具，不是正式测试。

用法：.venv\\Scripts\\python.exe tools\\check_recognize.py <docx路径>

打印每一条被认出身份的段落（部件名 + meta），方便跟原文对着看识别得
对不对。大规模/穷举式的验证交给专门的测试 Agent，这里只做"跑得通、
看着大致对"的初步核对。
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from essay_tool.engine.recognize import open_and_recognize


def main() -> None:
    if len(sys.argv) < 2:
        print("用法：check_recognize.py <docx路径>")
        raise SystemExit(1)
    path = Path(sys.argv[1])
    document, index = open_and_recognize(path)

    paragraphs = document.paragraphs
    ordinal = 0
    hits = 0
    for entry in index:
        if entry.kind != "paragraph":
            continue
        i = ordinal
        ordinal += 1
        if not entry.part:
            continue
        hits += 1
        text = paragraphs[i].text.strip()
        meta = {k: v for k, v in entry.meta.items()}
        print(f"[{i:>4}] {entry.part:<8} {meta} {text[:40]!r}")

    print("-" * 96)
    print(f"总段数 {len(paragraphs)}，识别出身份的段落 {hits} 条")


if __name__ == "__main__":
    main()
