"""去掉源文件开头的 UTF-8 BOM（三个字节 EF BB BF）。

用户 2026-09-16 裁决："清掉（excel_io/alerts/apply_structure/detect）"。

**为什么按字节改**：BOM 就是文件最前面那三个字节，按字节砍掉是唯一零风险的做法
——按文本读回来再写出去会被换行符翻译（Windows 上 `open(..., "w")` 默认
newline=None，写回去时 `\\n` 会变成 `\\r\\n`），那就不是"只少三字节"了。
**内容一字不动**，只少开头三字节。

用法：

    .venv\\Scripts\\python.exe tools\\strip_bom.py            # 只列清单
    .venv\\Scripts\\python.exe tools\\strip_bom.py --apply    # 真去
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BOM = b"\xef\xbb\xbf"


def main() -> int:
    ap = argparse.ArgumentParser(description="去掉源文件开头的 UTF-8 BOM")
    ap.add_argument("--apply", action="store_true", help="真去（默认只列清单）")
    ap.add_argument("--dirs", nargs="*", default=["essay_tool", "tools"],
                    help="扫描哪些目录（默认 essay_tool 与 tools）")
    args = ap.parse_args()

    targets: list[Path] = []
    for name in args.dirs:
        base = ROOT / name
        if base.is_dir():
            targets += sorted(p for p in base.rglob("*.py") if p.is_file())

    with_bom = [p for p in targets if p.read_bytes().startswith(BOM)]
    print(f"扫描 {len(targets)} 个 .py，带 BOM 的 {len(with_bom)} 个：")
    for p in with_bom:
        print(f"  {p.relative_to(ROOT)}")
    if not with_bom:
        print("（都很干净）")
        return 0
    if not args.apply:
        print("\n（只列清单。要真去：加 --apply）")
        return 0

    for p in with_bom:
        raw = p.read_bytes()
        p.write_bytes(raw[len(BOM):])
        print(f"  已去 BOM：{p.relative_to(ROOT)}（{len(raw)} → {len(raw) - 3} 字节）")
    print(f"\n完成，处理了 {len(with_bom)} 个文件（内容一字未动）。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
