"""回退风险体检 + 工作区状态指纹。

**这个脚本要解决的问题**（2026-09-16 事故取证结论）：
VS Code 的聊天编辑会话会把每个文件的"改动前内容（baseline）"存在
`%APPDATA%\\Code\\User\\workspaceStorage\\<hash>\\chatEditingSessions\\<会话id>\\state.json`
的 `timeline.fileBaselines` 里，`contents\\` 是内容寻址的块存储。
一旦发生整批撤销 / 丢弃 / 检查点回退（不一定要点按钮，会话重载时也会），
**这些 baseline 会被逐个写回磁盘**。会话活得越久，baseline 越旧、
波及面越大（本项目的会话已累积 1056 个请求、284 个文件）。

四个子命令：

* `report`（默认）：列出**现在还挂在旧 baseline 上的文件**（＝一次回退会退版的
  清单），按影响字节数排序，落盘 `work/_revert_risk.txt`。
* `fingerprint`：把工作区当前状态（md5 + 行数 + 字节数 + mtime）写成
  `work/_state_fingerprint.json`。**每次开工前打一份**，出事后才有比对基准。
* `check`：与最近一份指纹比对，报出"变了/没了/新出现"的文件；若某个文件的
  内容正好等于某条 baseline，就直接指出**退回的是哪个会话、哪条请求的哪一版**。
* `forensics <文件>`：把某个文件的内容（或 `--at` 指定的本地历史快照）与所有
  会话的 baseline 逐字比对，报出命中项。事故当天就是靠这个定位到真凶的。

⚠️ 只读：不修改工作区任何文件（`--fingerprint` 除外，它只写 work/ 下的指纹）。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
from pathlib import Path
from urllib.parse import unquote, urlparse

ROOT = Path(__file__).resolve().parents[1]
WORK = ROOT / "work"
FINGERPRINT = WORK / "_state_fingerprint.json"
RISK_REPORT = WORK / "_revert_risk.txt"

# 指纹只覆盖"会被回退破坏"的东西：生产代码、工具、文档、预设、根目录脚本。
INCLUDE_DIRS = ("essay_tool", "tools", "docs", "presets", "samples")
INCLUDE_ROOT = ("*.md", "*.py", "*.cmd")
SKIP_DIRS = {"work", "backup", ".venv", "__pycache__", ".git"}


def session_dirs() -> list[Path]:
    base = Path(os.environ.get("APPDATA", "")) / "Code" / "User" / "workspaceStorage"
    if not base.is_dir():
        return []
    return sorted(p for p in base.glob("*/chatEditingSessions/*") if p.is_dir())


def norm(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n")


def uri_to_rel(uri: str) -> str | None:
    """file:/// URI -> 相对工作区路径；不在工作区里返回 None。"""
    try:
        raw = urlparse(uri).path
    except ValueError:
        return None
    if not raw:
        return None
    p = Path(unquote(raw.lstrip("/")))
    try:
        return p.relative_to(ROOT).as_posix()
    except ValueError:
        return None


def load_baselines() -> list[dict]:
    """工作区内的全部 baseline 记录。⚠️ 会话状态最大 20MB+，读一次要几秒。"""
    rows: list[dict] = []
    for d in session_dirs():
        state = d / "state.json"
        if not state.exists() or state.stat().st_size < 1024:
            continue                       # 空会话（0.3KB）不必读
        try:
            data = json.loads(state.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        for key, val in ((data.get("timeline") or {}).get("fileBaselines") or []):
            if not isinstance(val, dict):
                continue
            text = val.get("content")
            if not isinstance(text, str):
                continue
            rel = uri_to_rel(str(key).split("::")[0])
            if rel is None:
                continue
            rows.append({
                "session": d.name[:8],
                "request": str(key).split("::")[-1],
                "epoch": val.get("epoch"),
                "rel": rel,
                "text": text,
            })
    return rows


def iter_files() -> list[Path]:
    out: list[Path] = []
    for sub in INCLUDE_DIRS:
        base = ROOT / sub
        if not base.is_dir():
            continue
        for p in base.rglob("*"):
            if p.is_file() and not (SKIP_DIRS & set(p.relative_to(ROOT).parts)):
                out.append(p)
    for pat in INCLUDE_ROOT:
        out += [p for p in ROOT.glob(pat) if p.is_file()]
    return sorted(out)


def digest(p: Path) -> dict:
    raw = p.read_bytes()
    return {
        "md5": hashlib.md5(raw).hexdigest(),
        "bytes": len(raw),
        "lines": raw.count(b"\n") + 1,
        "mtime": round(p.stat().st_mtime),
    }


def cmd_fingerprint() -> int:
    WORK.mkdir(exist_ok=True)
    data = {p.relative_to(ROOT).as_posix(): digest(p) for p in iter_files()}
    FINGERPRINT.write_text(
        json.dumps({"root": str(ROOT), "files": data}, ensure_ascii=False, indent=1),
        encoding="utf-8")
    print(f"已记录 {len(data)} 个文件的状态指纹 -> {FINGERPRINT}")
    return 0


def cmd_report() -> int:
    rows = load_baselines()
    if not rows:
        print("没有找到任何聊天编辑会话（或工作区内没有 baseline）。")
        return 0
    sessions = sorted({r["session"] for r in rows})
    worst: dict[str, dict] = {}
    for r in rows:
        p = ROOT / r["rel"]
        if not p.is_file():
            continue
        cur = p.read_text(encoding="utf-8", errors="replace")
        if norm(cur) == norm(r["text"]):
            continue                       # 这条 baseline 和磁盘现状一致，无所谓
        diff = len(r["text"].encode("utf-8")) - len(cur.encode("utf-8"))
        old = worst.get(r["rel"])
        if old is None or abs(diff) > abs(old["diff"]):
            worst[r["rel"]] = {"diff": diff, "epoch": r["epoch"],
                               "session": r["session"], "request": r["request"]}

    lines = [
        "＝ 回退风险清单 ＝",
        f"工作区：{ROOT}",
        f"扫描到的聊天编辑会话：{', '.join(sessions)}（共 {len(sessions)} 个，baseline 记录 {len(rows)} 条）",
        f"仍挂在旧 baseline 上的文件：{len(worst)} 个",
        "",
        "含义：这些文件一旦被整批撤销/丢弃/检查点回退，就会退回旧版；",
        "      数值是该文件会变化的字节数（负数＝会变小，即丢内容）。",
        "      最后一列＝退回的那一版来自哪个会话、哪条请求。",
        "",
    ]
    ranked = sorted(worst.items(), key=lambda kv: -abs(kv[1]["diff"]))
    for rel, info in ranked:
        lines.append(f"  {info['diff']:+9d}B  epoch={str(info['epoch']):<5} "
                     f"{info['session']}/{info['request'][-8:]}  {rel}")
    text = "\n".join(lines)
    RISK_REPORT.write_text(text, encoding="utf-8")
    print(text if len(ranked) <= 25 else "\n".join(lines[:8]) + f"\n  … 共 {len(ranked)} 个文件，"
          f"完整清单见 {RISK_REPORT}")
    print(f"\n完整清单：{RISK_REPORT}")
    return 0


def cmd_check() -> int:
    if not FINGERPRINT.exists():
        print(f"还没有指纹文件（{FINGERPRINT}）。先跑一次 --fingerprint。")
        return 1
    old = json.loads(FINGERPRINT.read_text(encoding="utf-8"))["files"]
    new = {p.relative_to(ROOT).as_posix(): digest(p) for p in iter_files()}
    rows = load_baselines()

    changed = [k for k in old if k in new and old[k]["md5"] != new[k]["md5"]]
    gone = [k for k in old if k not in new]
    added = [k for k in new if k not in old]
    print(f"与指纹相比：改动 {len(changed)}、消失 {len(gone)}、新增 {len(added)}")
    for k in changed:
        d = new[k]["lines"] - old[k]["lines"]
        print(f"  [改] {k:<46} 行数 {old[k]['lines']} -> {new[k]['lines']} ({d:+d})")

    hits = []
    for k in changed + gone:
        p = ROOT / k
        if not p.is_file():
            continue
        cur = norm(p.read_text(encoding="utf-8", errors="replace"))
        for r in rows:
            if r["rel"] == k and norm(r["text"]) == cur:
                hits.append((k, r["session"], r["request"], r["epoch"]))
    if hits:
        print("\n★ 下列文件的当前内容正好等于某条 baseline，高度疑似被回退：")
        for k, s, req, ep in hits:
            print(f"  {k}  ← 会话 {s} 请求 {req} (epoch {ep})")
        print("   → 可用 tools/restore_from_local_history.py 或备份快照还原。")
    else:
        print("\n没有发现“退回旧版”的特征（改动看着像正常编辑）。")
    return 0


def cmd_forensics(target: str, at: str | None) -> int:
    p = Path(target)
    if not p.is_absolute():
        p = ROOT / target
    if at:
        raw = history_snapshot(p, at)
        if raw is None:
            print(f"在本地历史里没找到 {p.name} 在 {at} 那一版。")
            return 1
        text = raw
    elif p.is_file():
        text = p.read_text(encoding="utf-8", errors="replace")
    else:
        print(f"文件不存在：{p}")
        return 1

    rows = load_baselines()
    hits = [r for r in rows if Path(r["rel"]).name == p.name and norm(r["text"]) == norm(text)]
    print(f"内容长度 {len(text.encode('utf-8'))}B，与 {len(rows)} 条 baseline 比对：")
    if not hits:
        print("  没有命中任何 baseline。")
        return 0
    for r in hits:
        print(f"  ★ 逐字一致：会话 {r['session']} 请求 {r['request']} epoch={r['epoch']}，"
              f"对应文件 {r['rel']}")
    print("\n结论：这份内容就是该会话在那些请求处的『改动前内容』，"
          "写回它属于整批回退，不是正常编辑。")
    return 0


def history_snapshot(target: Path, at: str) -> str | None:
    """从 VS Code 本地历史里取某个时间点（"MM-dd HH:mm:ss"）的快照内容。"""
    import datetime
    hist = Path(os.environ.get("APPDATA", "")) / "Code" / "User" / "History"
    if not hist.is_dir():
        return None
    for d in hist.iterdir():
        entries = d / "entries.json"
        if not entries.exists():
            continue
        try:
            j = json.loads(entries.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        if Path(unquote(urlparse(j.get("resource", "")).path.lstrip("/"))) != target:
            continue
        for ent in j.get("entries", []):
            t = datetime.datetime.fromtimestamp(ent["timestamp"] / 1000)
            if t.strftime("%m-%d %H:%M:%S") == at:
                f = d / ent["id"]
                if f.is_file():
                    return f.read_text(encoding="utf-8", errors="replace")
    return None


def main() -> int:
    ap = argparse.ArgumentParser(description="回退风险体检 + 状态指纹")
    ap.add_argument("command", nargs="?", default="report",
                    choices=("report", "fingerprint", "check", "forensics"))
    ap.add_argument("target", nargs="?")
    ap.add_argument("--at", help="本地历史时间点，格式 MM-dd HH:mm:ss（配合 forensics）")
    args = ap.parse_args()

    if args.command == "fingerprint":
        return cmd_fingerprint()
    if args.command == "check":
        return cmd_check()
    if args.command == "forensics":
        if not args.target:
            print("用法：check_revert_risk.py forensics <文件> [--at 09-16 19:24:12]")
            return 2
        return cmd_forensics(args.target, args.at)
    return cmd_report()


if __name__ == "__main__":
    sys.exit(main())
