"""本工具的磁盘落点：随程序安装的只读资源 vs 用户自己可写的数据目录 vs 临时工作目录。

设计前提是"打包成一个正经的 Windows 安装包"（有安装目录、系统的"添加或删除
程序"里认得到、能整个卸载）：安装目录（不管是 Program Files 还是用户自己选的
哪个目录）在标准权限下往往是**只读**的，卸载时又会被整个删掉——用户自己保存的
东西绝不能放在那里。所以三类落点各管各的：

- `resource_dir()`：随程序一起分发的**只读资源**（内置预设、`reference/` 里的官方
  文件、封面样张图、图标）的根。开发时是仓库根目录；打包成 exe 后是 PyInstaller
  的 `_internal` 目录（`sys._MEIPASS`）。`builtin_presets_dir()` 挂在它下面。
- `user_data_dir()`：每用户可写、跟安装/卸载生命周期脱钩的数据目录，也就是
  `%APPDATA%\\AutoEssayTypesettingTool`。`user_presets_dir()`（用户保存/导入的
  预设）、`import_backup_dir()`（导入论文时的自动备份）、设置文件、崩溃日志都在这里。
- `session_work_dir()`：**本次运行的临时工作目录**（`%TEMP%` 下），导入论文得到的
  工作副本和导出时的中间文件都放这里，程序退出时整个删掉、不留痕迹
  （用户 2026-09-14 定"用完就删掉，不保存"）。

开发阶段（没打包成 exe）`resource_dir()` 就是仓库根目录，行为与以前一致。
"""

from __future__ import annotations

import atexit
import os
import pathlib
import shutil
import sys
import tempfile
import time

APP_DIR_NAME = "AutoEssayTypesettingTool"

ROOT = pathlib.Path(__file__).resolve().parents[2]

# 上次异常退出（崩溃/被杀）留在临时目录里的工作目录，超过这个时长就当垃圾清掉
_STALE_WORK_SECONDS = 2 * 24 * 3600


def is_frozen() -> bool:
    """是不是已经打包成 exe（PyInstaller 等打包工具运行时会设置 sys.frozen）。"""
    return bool(getattr(sys, "frozen", False))


def install_dir() -> pathlib.Path:
    """程序自身所在目录：开发时是仓库根目录，打包后是 exe 所在目录。

    **别拿它找随程序分发的资源**——PyInstaller 6 把资源放进 `_internal`，跟 exe
    不在同一层，找资源请用 `resource_dir()`。这个函数现在只用来回答"程序装在哪"。
    """
    if is_frozen():
        return pathlib.Path(sys.executable).resolve().parent
    return ROOT


def resource_dir() -> pathlib.Path:
    """随程序分发的只读资源根目录（`presets/`、`reference/`、`essay_tool/gui/assets/`
    都在它下面，目录结构与仓库一致）。"""
    if is_frozen():
        return pathlib.Path(getattr(sys, "_MEIPASS", install_dir()))
    return ROOT


def builtin_presets_dir() -> pathlib.Path:
    """随程序安装的默认预设目录（只读）。"""
    return resource_dir() / "presets"


def user_data_dir() -> pathlib.Path:
    """每用户可写的数据目录：`%APPDATA%\\AutoEssayTypesettingTool`。

    找不到 `APPDATA`（非 Windows 环境）时退到用户主目录下的隐藏文件夹，
    不让程序直接崩掉。目录不存在就顺手建好，调用方不用自己再判断。
    """
    base = os.environ.get("APPDATA")
    root = pathlib.Path(base) if base else pathlib.Path.home() / ".config"
    d = root / APP_DIR_NAME
    d.mkdir(parents=True, exist_ok=True)
    return d


def user_presets_dir() -> pathlib.Path:
    """用户自己保存/导入的预设目录（可读可写可删）。"""
    d = user_data_dir() / "presets"
    d.mkdir(parents=True, exist_ok=True)
    return d


def import_backup_dir() -> pathlib.Path:
    """导入论文时的自动备份目录——用户手滑删了自己的原始文件时的安全网。

    2026-09-20 起放在 `user_data_dir()` 下（每用户可写、卸载安装目录时不会跟着没）。
    此前开发阶段曾暂放安装目录里（2026-09-14 用户拍板的临时做法），打包前按约定挪走。
    """
    d = user_data_dir() / "backup"
    d.mkdir(parents=True, exist_ok=True)
    return d


# ---------------------------------------------------------------- 临时工作目录
_session_work: pathlib.Path | None = None


def _work_parent() -> pathlib.Path:
    return pathlib.Path(tempfile.gettempdir()) / APP_DIR_NAME


def session_work_dir() -> pathlib.Path:
    """本次运行的临时工作目录（第一次调用时才创建；程序退出时整个删除）。

    放在 `%TEMP%\\AutoEssayTypesettingTool\\run-<进程号>`：导入论文时得到的"工作副本"、
    导出时引擎的中间文件都在这里。用户的原文件与导出结果都不在这里，删掉不损失任何东西。
    """
    global _session_work
    if _session_work is None or not _session_work.is_dir():
        d = _work_parent() / f"run-{os.getpid()}"
        d.mkdir(parents=True, exist_ok=True)
        _session_work = d
        atexit.register(cleanup_session_work_dir)
    return _session_work


def cleanup_session_work_dir() -> None:
    """删掉本次运行的临时工作目录（退出时自动调用；删不掉不算错，下次启动会补扫）。"""
    global _session_work
    d, _session_work = _session_work, None
    if d is not None:
        shutil.rmtree(d, ignore_errors=True)


def purge_stale_work_dirs() -> None:
    """启动时清掉上次崩溃/被强杀遗留的旧工作目录（只动超过两天没碰过的，
    不去猜别的进程死没死——同时开着两个本工具的窗口时不能误删对方的）。"""
    parent = _work_parent()
    try:
        entries = list(parent.glob("run-*"))
    except OSError:
        return
    now = time.time()
    for d in entries:
        try:
            if d.is_dir() and now - d.stat().st_mtime > _STALE_WORK_SECONDS:
                shutil.rmtree(d, ignore_errors=True)
        except OSError:
            continue
