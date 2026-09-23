"""生成随安装包分发的许可声明文件，并保证项目根目录的 `LICENSE` 与之一致。

**单一出处**：本工具自身的许可（`_OWN_LICENSE`）与第三方依赖库的许可原文
（`_THIRD_PARTY_NOTICES`）都只写在 `essay_tool/gui/main_window.py` 里——关于页链接、
项目根目录 `LICENSE`、安装包里的 `licenses\\*.txt` 全部从那里来，用 `ast` 读出，
不 import 主窗口（不弹窗、不依赖 tkinter）。

输出（`build/licenses/`，由 `installer.iss` 装到 `{app}\\licenses\\`）：
  LICENSE.txt                    本工具自身（MIT，Copyright (c) 2026 Grandliu66）
  THIRD-PARTY-NOTICES.txt        第三方库许可原文 + 随包的运行时组件清单
  Python-LICENSE.txt             打包所用 Python 自带的许可文件——里面含 PSF 许可，以及
                                 随 Python 分发的 OpenSSL（Apache-2.0 全文）、bzip2、libffi 等的条款
  ChineseSimplified.LICENSE.txt  安装向导简体中文翻译（MIT）的许可
同时把 `LICENSE`（项目根目录）写成与 LICENSE.txt 相同的内容。

文本一律 UTF-8 + BOM、CRLF 换行（记事本双击即可读）。
"""

from __future__ import annotations

import ast
import pathlib
import shutil
import sys

ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC = ROOT / "essay_tool" / "gui" / "main_window.py"
OUT = ROOT / "build" / "licenses"

# 随包分发、但不在 Python 库许可原文里的运行时组件（给读者一个总览）
RUNTIME_COMPONENTS = """\
Runtime components shipped in the application folder
====================================================

* Python 3.13 runtime (python313.dll and the standard library) -- Python Software
  Foundation License. Full text: Python-LICENSE.txt (in this folder). That file also
  contains the terms of components distributed with the Windows build of Python:
  the conditions for Microsoft's Distributable Code (the Visual C++ runtime),
  bzip2 (_bz2.pyd), libffi (libffi-8.dll; MIT-style license), and, at the end, the full
  text of the Apache License 2.0 -- the license of OpenSSL 3.x (libcrypto-3.dll,
  libssl-3.dll), which Python ships.
* Tcl/Tk 8.6 (tcl86t.dll, tk86t.dll, _tk_data, _tcl_data) -- Tcl/Tk license. Full text:
  _internal\\_tk_data\\license.terms (in the application folder).
* zlib (zlib1.dll) -- zlib license (permissive; no attribution required for binaries).
* liblzma / XZ Utils (_lzma.pyd) -- public domain.
* Microsoft Visual C++ Runtime (VCRUNTIME140.dll, VCRUNTIME140_1.dll) -- redistributed
  under Microsoft's Distributable Code terms (conditions reproduced in Python-LICENSE.txt,
  section "Additional Conditions for this Windows binary build").
* PyInstaller bootloader (the .exe launcher) -- GPL-2.0 with the special exception that
  allows the bootloader to be distributed as part of an application under any license.
* Installer only: Simplified Chinese translation of the Inno Setup wizard by Kira
  (https://github.com/kira-96/Inno-Setup-Chinese-Simplified-Translation) -- MIT License,
  see ChineseSimplified.LICENSE.txt.
"""


def _const(name: str):
    tree = ast.parse(SRC.read_text(encoding="utf-8"))
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(
                isinstance(t, ast.Name) and t.id == name for t in node.targets):
            return ast.literal_eval(node.value)
    raise SystemExit(f"找不到 main_window.py 里的 {name}")


def _write(path: pathlib.Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    body = text.replace("\r\n", "\n").rstrip("\n") + "\n"
    path.write_bytes(b"\xef\xbb\xbf" + body.replace("\n", "\r\n").encode("utf-8"))


def main() -> int:
    own_name, own_license, own_text = _const("_OWN_LICENSE")
    third = _const("_THIRD_PARTY_NOTICES")

    # 项目根目录 LICENSE：纯 MIT 原样（不带中文说明；带 BOM 会让 GitHub 显示异常，所以这份不加 BOM、用 LF）
    (ROOT / "LICENSE").write_text(own_text.strip("\n") + "\n", encoding="utf-8", newline="\n")

    _write(OUT / "LICENSE.txt", own_text)

    parts = [f"Third-party software notices for {own_name}\n"
             "=" * 60 + "\n\n"
             "This program includes the following third-party libraries. Their license\n"
             "texts are reproduced verbatim from the libraries' own LICENSE files.\n"]
    for name, license_name, text in third:
        parts.append(f"\n{'-' * 60}\n{name} -- {license_name}\n{'-' * 60}\n\n{text.strip()}\n")
    parts.append("\n" + "=" * 60 + "\n\n" + RUNTIME_COMPONENTS)
    _write(OUT / "THIRD-PARTY-NOTICES.txt", "".join(parts))

    py_license = pathlib.Path(sys.base_prefix) / "LICENSE.txt"
    if not py_license.is_file():
        raise SystemExit(f"找不到 Python 自带的许可文件：{py_license}")
    _write(OUT / "Python-LICENSE.txt", py_license.read_text(encoding="utf-8", errors="replace"))

    zh = ROOT / "packaging" / "languages" / "ChineseSimplified.LICENSE.txt"
    shutil.copyfile(zh, OUT / "ChineseSimplified.LICENSE.txt")

    files = sorted(p.name for p in OUT.iterdir())
    print(f"已生成：build\\licenses\\（{len(files)} 个文件：{'、'.join(files)}）；根目录 LICENSE 已同步")
    return 0


if __name__ == "__main__":
    sys.exit(main())
