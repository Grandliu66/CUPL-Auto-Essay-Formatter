"""清掉"孤儿"聊天编辑会话：对话已删除、编辑状态却还在磁盘上的那些。

**为什么需要它**（2026-09-16 事故）：
VS Code 把聊天编辑状态存在
`%APPDATA%\\Code\\User\\workspaceStorage\\<hash>\\chatEditingSessions\\<会话id>\\`，
里面 `state.json` 的 `timeline.fileBaselines` 记着每个文件"改动前的内容"，
整批撤销/丢弃/检查点回退时会把这些内容**写回磁盘**。

实测：在聊天视图里删除会话（带"不可撤销"确认）只删掉
`...\\workspaceStorage\\<hash>\\chatSessions\\<会话id>.jsonl`（也就是"对话"本身），
`chatEditingSessions\\<会话id>\\` **原封不动**。所以删完对话，回退风险一点没降。

这个脚本找出"没有对应对话文件的编辑会话目录"（孤儿），把它们整体搬到项目隔壁的
备份目录里。搬走之后，那些会话再也没有能力写回任何文件。

用法：

    python tools/purge_orphan_editing_sessions.py            # 只列出来，什么都不动
    python tools/purge_orphan_editing_sessions.py --apply    # 真的搬走（先关掉 VS Code）
    ... --apply --keep 3f8a1c...                             # 额外保留某个会话 id 前缀

⚠️ 建议先关掉 VS Code 再 `--apply`：窗口开着时，编辑器可能把手里的会话状态重新写回磁盘。
脚本默认检测到 Code.exe 在跑就只警告、不动手（要强行执行加 `--force`）。
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote, urlparse

ROOT = Path(__file__).resolve().parents[1]


def workspace_storage() -> Path | None:
    """找到本项目对应的 workspaceStorage 目录（按 workspace.json 里的文件夹路径认）。"""
    base = Path.home() / "AppData" / "Roaming" / "Code" / "User" / "workspaceStorage"
    if not base.is_dir():
        return None
    want = str(ROOT).lower()
    for d in base.iterdir():
        wj = d / "workspace.json"
        if not wj.is_file():
            continue
        try:
            info = json.loads(wj.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        for key in ("folder", "workspace"):
            uri = info.get(key)
            if not isinstance(uri, str):
                continue
            got = unquote(urlparse(uri).path.lstrip("/")).replace("/", "\\").lower()
            if got.startswith(want):
                return d
    return None


def code_running() -> bool:
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq Code.exe"],
                             capture_output=True, text=True, timeout=20)
    except (OSError, subprocess.SubprocessError):
        return False
    return "Code.exe" in (out.stdout or "")


def dir_size(d: Path) -> tuple[int, int]:
    files = [p for p in d.rglob("*") if p.is_file()]
    return len(files), sum(p.stat().st_size for p in files)


def main() -> int:
    ap = argparse.ArgumentParser(description="清掉孤儿聊天编辑会话（对话已删、状态还在）")
    ap.add_argument("--apply", action="store_true", help="真的搬走（默认只列清单）")
    ap.add_argument("--force", action="store_true", help="VS Code 在跑也照做")
    ap.add_argument("--keep", action="append", default=[],
                    help="额外保留的会话 id 前缀，可重复")
    args = ap.parse_args()

    store = workspace_storage()
    if store is None:
        print("找不到本项目对应的 workspaceStorage 目录。")
        return 1
    edit_root = store / "chatEditingSessions"
    chat_root = store / "chatSessions"
    if not edit_root.is_dir():
        print(f"没有编辑会话目录：{edit_root}")
        return 1

    chats = {p.stem for p in chat_root.glob("*.jsonl")} if chat_root.is_dir() else set()
    print(f"工作区存储：{store}")
    print(f"对话文件（chatSessions）：{len(chats)} 个")
    print(f"编辑会话（chatEditingSessions）：{len(list(edit_root.iterdir()))} 个\n")

    orphans = []
    for d in sorted(edit_root.iterdir()):
        if not d.is_dir() or d.name in chats:
            continue
        if any(d.name.startswith(k) for k in args.keep):
            print(f"  跳过（--keep）：{d.name}")
            continue
        n, size = dir_size(d)
        orphans.append((d, n, size))

    if not orphans:
        print("没有孤儿编辑会话，磁盘上很干净。")
        return 0

    orphans.sort(key=lambda x: -x[2])
    print("孤儿编辑会话（对话已删、状态还在，回退时仍会写回文件）：")
    total = 0
    for d, n, size in orphans:
        total += size
        print(f"  {d.name}  内容块 {n:5d}  合计 {size / 1024:9.1f}KB")
    print(f"  共计 {len(orphans)} 个，{total / 1024 / 1024:.1f}MB")

    if not args.apply:
        print("\n（只列清单。要真的搬走：加 --apply）")
        return 0

    if code_running() and not args.force:
        print("\n⚠️ 检测到 Code.exe 正在运行。请先关闭 VS Code 再执行 --apply，"
              "否则窗口可能把手里的会话状态重新写回磁盘。")
        print("   确定要继续就加 --force。")
        return 2

    dest = ROOT.parent / f"{ROOT.name}-备份" / f"editor_sessions_{datetime.now():%Y%m%d-%H%M%S}"
    dest.mkdir(parents=True, exist_ok=True)
    print(f"\n搬到：{dest}")
    for d, _n, _size in orphans:
        target = dest / d.name
        if target.exists():
            shutil.rmtree(target)
        shutil.move(str(d), str(target))
        print(f"  已移走 {d.name}")
    print(f"\n完成。若想还原，把 {dest} 里的目录搬回 {edit_root} 即可。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
