"""给项目打一份快照备份（本机没有装 git，所以用纯复制，只用标准库）。

备份到哪里：项目**隔壁**的 `Auto Essay Typesetting Tool-备份\\<年月日-时分秒>\\`
（放在隔壁是为了连项目文件夹一起出事时备份还在）。

备份哪些：
  essay_tool/、tools/、docs/、presets/、samples/、backup/、根目录下的 *.md 与 *.py
  reference/ 只收非 .docx 的文件（那些是测试样张，一套 9MB；要收加 --with-docx）
不备份：.venv/、work/、__pycache__/、*.pyc（可重建，或者是草稿区）

每次打完自动只保留最近的若干份（默认 30 份），更旧的删掉。

跑法：
  .venv\\Scripts\\python.exe tools\\backup_snapshot.py            打一份快照
  .venv\\Scripts\\python.exe tools\\backup_snapshot.py --list      列出已有快照
  .venv\\Scripts\\python.exe tools\\backup_snapshot.py --restore 20260916-2015 --dry-run
  .venv\\Scripts\\python.exe tools\\backup_snapshot.py --restore 20260916-2015
        （真还原；还原前会自动先给当前状态打一份快照，不会删任何多出来的文件）

也可以直接双击项目根目录下的「备份.cmd」。
"""

from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import pathlib
import shutil
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = pathlib.Path(__file__).resolve().parents[1]
SNAPSHOT_ROOT = ROOT.parent / f"{ROOT.name}-备份"

COPY_DIRS = ("essay_tool", "tools", "docs", "presets", "samples", "backup")
COPY_GLOBS = ("*.md", "*.py", "*.cmd")
SKIP_NAMES = {"__pycache__", ".venv", "work", ".git"}
SKIP_SUFFIXES = {".pyc", ".pyo"}
REFERENCE_DIR = "reference"
DEFAULT_KEEP = 30


def iter_sources(with_docx: bool = False):
    """产出 (绝对路径, 相对项目根的路径)。"""
    for name in COPY_DIRS:
        folder = ROOT / name
        if not folder.exists():
            continue
        for path in sorted(folder.rglob("*")):
            if not path.is_file():
                continue
            if any(part in SKIP_NAMES for part in path.parts):
                continue
            if path.suffix.lower() in SKIP_SUFFIXES:
                continue
            if name == REFERENCE_DIR and path.suffix.lower() == ".docx" and not with_docx:
                continue
            yield path, path.relative_to(ROOT)
    for pattern in COPY_GLOBS:
        for path in sorted(ROOT.glob(pattern)):
            if path.is_file():
                yield path, path.relative_to(ROOT)


def digest(path: pathlib.Path) -> str:
    h = hashlib.md5()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def snapshots() -> list[pathlib.Path]:
    if not SNAPSHOT_ROOT.exists():
        return []
    return sorted((p for p in SNAPSHOT_ROOT.iterdir() if p.is_dir()), reverse=True)


def make_snapshot(with_docx: bool = False, keep: int = DEFAULT_KEEP,
                  quiet: bool = False) -> pathlib.Path:
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    target = SNAPSHOT_ROOT / stamp
    target.mkdir(parents=True, exist_ok=True)

    manifest: dict[str, dict] = {}
    total = 0
    for source, relative in iter_sources(with_docx):
        dest = target / relative
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, dest)
        stat = source.stat()
        manifest[str(relative)] = {"size": stat.st_size, "mtime": stat.st_mtime,
                                   "md5": digest(source)}
        total += stat.st_size

    (target / "_manifest.json").write_text(
        json.dumps({"created": stamp, "files": manifest}, ensure_ascii=False, indent=1),
        encoding="utf-8")
    (target / "_info.txt").write_text(
        f"项目快照\n时间：{datetime.datetime.now():%Y-%m-%d %H:%M:%S}\n"
        f"来源：{ROOT}\n文件：{len(manifest)} 个\n合计：{total / 1024 / 1024:.1f} MB\n"
        f"含 reference 里的 .docx：{'是' if with_docx else '否'}\n"
        f"还原：tools\\backup_snapshot.py --restore {stamp}\n",
        encoding="utf-8")

    if not quiet:
        print(f"已备份到 {target}")
        print(f"  文件 {len(manifest)} 个，合计 {total / 1024 / 1024:.1f} MB")

    for old in snapshots()[keep:]:
        shutil.rmtree(old, ignore_errors=True)
        if not quiet:
            print(f"  清掉过旧的一份：{old.name}")
    return target


def list_snapshots() -> None:
    items = snapshots()
    if not items:
        print(f"还没有快照。备份目录：{SNAPSHOT_ROOT}")
        return
    print(f"备份目录：{SNAPSHOT_ROOT}\n共 {len(items)} 份：")
    for folder in items:
        info = folder / "_info.txt"
        head = info.read_text(encoding="utf-8").splitlines() if info.exists() else []
        summary = "  ".join(line for line in head[2:5])
        print(f"  {folder.name}   {summary}")


def restore(name: str, dry_run: bool) -> None:
    source = SNAPSHOT_ROOT / name
    if not source.exists():
        raise SystemExit(f"没有这份快照：{source}")

    changed: list[str] = []
    added: list[str] = []
    for path in sorted(source.rglob("*")):
        if not path.is_file() or path.name.startswith("_"):
            continue
        relative = path.relative_to(source)
        live = ROOT / relative
        if not live.exists():
            added.append(str(relative))
        elif digest(live) != digest(path):
            changed.append(str(relative))

    print(f"快照 {name}：需要覆盖 {len(changed)} 个文件、新建 {len(added)} 个文件")
    for item in changed:
        print(f"  改：{item}")
    for item in added:
        print(f"  新：{item}")
    if dry_run:
        print("\n（--dry-run，未改动任何文件）")
        return

    safety = make_snapshot(quiet=True)
    print(f"\n还原前先把当前状态打了一份：{safety.name}")
    for path in sorted(source.rglob("*")):
        if not path.is_file() or path.name.startswith("_"):
            continue
        live = ROOT / path.relative_to(source)
        live.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, live)
    print(f"已还原 {len(changed) + len(added)} 个文件（多出来的文件没有删）")


def main() -> None:
    parser = argparse.ArgumentParser(description="项目快照备份（不依赖 git）")
    parser.add_argument("--list", action="store_true", help="列出已有快照")
    parser.add_argument("--restore", metavar="快照名", help="从某份快照还原")
    parser.add_argument("--dry-run", action="store_true", help="还原前先看会改哪些文件")
    parser.add_argument("--with-docx", action="store_true", help="连 reference 里的 .docx 一起备份")
    parser.add_argument("--keep", type=int, default=DEFAULT_KEEP, help="保留最近几份（默认 30）")
    args = parser.parse_args()

    if args.list:
        list_snapshots()
    elif args.restore:
        restore(args.restore, args.dry_run)
    else:
        make_snapshot(args.with_docx, args.keep)


if __name__ == "__main__":
    main()
