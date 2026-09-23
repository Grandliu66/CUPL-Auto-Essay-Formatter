"""重新生成 `essay_tool/config/integrity_hashes.py`（官方文件的出厂 SHA-256 清单）。

什么时候要跑：reference/ 里那几份随程序分发的官方文件换了（比如十星封面范例
换届、学校发了新版模板）之后，**打包之前**跑一次：

    .venv\\Scripts\\python.exe tools\\update_integrity.py

它按 `essay_tool` 运行时找这些文件的同一套规则去 reference/ 里找（见 `PATTERNS`），
所以清单永远只含"运行时真会读的"文件。没找到某一类会直接报错，不写半份清单。
"""

from __future__ import annotations

import hashlib
import pathlib

ROOT = pathlib.Path(__file__).resolve().parents[1]


def sha256_of(path: pathlib.Path) -> str:
    # 不 import essay_tool.config.integrity：清单文件还没生成时它自己导入会失败
    return hashlib.sha256(path.read_bytes()).hexdigest()

# 与 cover.py / spec_viewer.py 里的取文件规则保持一致
PATTERNS = (
    ("论文/学年论文封面（写作模板）", "中国政法大学本科生毕业论文写作模板.docx"),
    ("课程考核封面", "中国政法大学本科生课程考核封面.docx"),
    ("十星封面范例", "*学术十星*封面*.docx"),
    ("《规范》", "*规范*.docx"),
    ("十星格式要求 PDF", "*学术十星*.pdf"),
)

OUT = ROOT / "essay_tool" / "config" / "integrity_hashes.py"


def main() -> int:
    ref = ROOT / "reference"
    lines = []
    for label, pattern in PATTERNS:
        found = sorted(p for p in ref.glob(pattern) if not p.name.startswith("~$"))
        if not found:
            print(f"出错：reference/ 里找不到「{label}」（{pattern}）")
            return 1
        # 与运行时一致：只取排序第一个（`*规范*` 会命中不止一份时以第一份为准）
        p = found[0]
        lines.append((label, p.name, sha256_of(p)))
        print(f"{label}：{p.name}  {lines[-1][2][:12]}…")
    body = ['"""官方文件出厂 SHA-256 清单——**自动生成，别手改**。',
            "",
            "生成：`tools/update_integrity.py`；用途见 `essay_tool/config/integrity.py`。",
            '"""',
            "",
            "OFFICIAL_SHA256 = frozenset({"]
    for label, name, digest in lines:
        body.append(f'    "{digest}",  # {label}：{name}')
    body.append("})")
    OUT.write_text("\n".join(body) + "\n", encoding="utf-8")
    print(f"已写出：{OUT.relative_to(ROOT)}（{len(lines)} 条）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
