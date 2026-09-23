"""把 GUI 里的《关于使用本工具的协议》原文导出成安装向导「许可协议」页要显示的 RTF 文件。

**单一出处**：协议正文只有一份，就是 `essay_tool/gui/main_window.py` 里的 `AGREEMENT_TEXT`
（用户 2026-09-14 定稿）。这里用 `ast` 从源码里把那个常量读出来（不 import 主窗口，
不弹窗、不依赖 tkinter），转成 `build/license_zh.rtf`——安装页与程序内的协议因此
永远一致，改协议只改那一处。

原文里的排版标记（程序内弹窗同一套规则，见 `MainWindow._open_text_popup(markup=True)`）：
    行首 "@"   → 居中加粗的大标题
    **…**      → 加粗
安装页里的协议**不留空行**（用户 2026-09-20），逐行去掉空白行，其余文字逐字不动。
用 RTF 而不是纯文本，是因为 Inno Setup 的许可页只有 RTF 能显示加粗与居中。
"""

from __future__ import annotations

import ast
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC = ROOT / "essay_tool" / "gui" / "main_window.py"
OUT = ROOT / "build" / "license_zh.rtf"

_BODY_HALF_PT = 21      # 正文字号（半磅：21 = 10.5 磅）
_HEAD_HALF_PT = 24      # 大标题（12 磅）


def read_agreement() -> str:
    tree = ast.parse(SRC.read_text(encoding="utf-8"))
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(
                isinstance(t, ast.Name) and t.id == "AGREEMENT_TEXT" for t in node.targets):
            value = ast.literal_eval(node.value)
            if isinstance(value, str) and value.strip():
                return value
    raise SystemExit("找不到 main_window.py 里的 AGREEMENT_TEXT")


def rtf_escape(text: str) -> str:
    """RTF 转义：`\\ { }` 加反斜杠，非 ASCII 用 `\\uN?`（N 是有符号 16 位）。"""
    out: list[str] = []
    for ch in text:
        o = ord(ch)
        if ch in "\\{}":
            out.append("\\" + ch)
        elif o < 128:
            out.append(ch)
        else:
            out.append(f"\\u{o if o <= 32767 else o - 65536}?")
    return "".join(out)


def to_rtf(agreement: str) -> str:
    paras: list[str] = []
    for line in agreement.replace("\r\n", "\n").split("\n"):
        line = line.rstrip()
        if not line.strip():
            continue                                  # 不留空行
        if line.startswith("@"):                      # 居中加粗大标题
            paras.append(r"\pard\qc\b\fs%d %s\b0\fs%d\par" % (
                _HEAD_HALF_PT, rtf_escape(line[1:]), _BODY_HALF_PT))
            continue
        pieces = line.split("**")
        body = "".join((r"\b " + rtf_escape(seg) + r"\b0 ") if i % 2 == 1 and seg
                       else rtf_escape(seg) for i, seg in enumerate(pieces))
        paras.append(r"\pard\ql " + body + r"\par")
    return ("{\\rtf1\\ansi\\ansicpg936\\deff0\\deflang2052\\uc1\n"
            "{\\fonttbl{\\f0\\fnil\\fcharset134 Microsoft YaHei UI;}}\n"
            "\\viewkind4\\f0\\fs%d\n" % _BODY_HALF_PT
            + "\n".join(paras) + "\n}")


def main() -> int:
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(to_rtf(read_agreement()), encoding="ascii")   # 全 ASCII（中文都是 \uN?）
    print(f"已生成：{OUT.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
