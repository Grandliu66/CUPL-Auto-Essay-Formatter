"""「检查更新」的界面部分：后台联网、结果提示、新版本弹窗。

联网本身在 `essay_tool/update_check.py`（只读一个 `latest.json`，不下载、不上传）；这里只管
**不卡界面**和**怎么告诉用户**：

  * 联网放在后台线程里，主线程用 `after` 轮询结果（Tk 不是线程安全的，不能从线程里碰控件）；
  * **自动检查**（启动时）：一天最多一次、读不到就**完全静默**（离线/被墙都不打扰用户）、
    用户点过"忽略此版本"的版本不再弹；
  * **手动检查**（「关于」页的按钮）：有新版本弹窗，已是最新或读不到都明确告诉用户；
  * 下载链接**只在浏览器里打开对应页面**，本程序不下载任何文件。

颜色/字体由调用方（`MainWindow`）传进来，避免跟 `main_window` 互相 import——开发时
`python essay_tool/gui/main_window.py` 直接运行，那个模块会以 `__main__` 的名字存在，
在这里反过来 import 它会得到第二份拷贝。
"""

from __future__ import annotations

import threading
import time
import tkinter as tk
import webbrowser
from dataclasses import dataclass
from tkinter import messagebox

from essay_tool import __version__
from essay_tool import update_check
from essay_tool.config.settings import save_settings

AUTO_CHECK_INTERVAL_SECONDS = 20 * 3600          # 自动检查：一天最多一次（留点余量，别卡在整 24 小时）
_POLL_MS = 200


@dataclass(frozen=True)
class Style:
    bg: str
    fg: str
    link: str
    family: str


def _save(app) -> None:
    try:
        save_settings(app.gui_settings)
    except OSError:
        pass                                      # 存不下去只是下次还会再提醒，不是故障


def open_in_browser(app, url: str) -> None:
    """打开一个下载页面。`url` 来自 `update_check.parse_latest`（只留 https + 白名单域名）。"""
    try:
        opened = bool(webbrowser.open(url))
    except Exception:                              # noqa: BLE001
        opened = False
    if not opened:
        app.status(f"打开浏览器失败。请手动访问：{url}")


def start_check(app, style: Style, *, manual: bool) -> None:
    """起一次检查。已经有一次在跑就不再起（手动点会提示"正在检查"）。"""
    if getattr(app, "_update_running", False):
        if manual:
            app.status("正在检查更新，请稍候……")
        return
    app._update_running = True
    holder: dict = {}

    def worker() -> None:
        try:
            holder["info"] = update_check.fetch_latest(__version__)
        except Exception:                          # noqa: BLE001 —— 联网代码不许把线程带崩
            holder["info"] = None
        holder["done"] = True

    threading.Thread(target=worker, name="update-check", daemon=True).start()

    def poll() -> None:
        if not holder.get("done"):
            try:
                app.after(_POLL_MS, poll)
            except tk.TclError:                    # 窗口已经关了
                pass
            return
        app._update_running = False
        try:
            handle_result(app, style, holder.get("info"), manual=manual)
        except Exception:                          # noqa: BLE001 —— 检查更新出任何问题都不该影响主程序
            pass

    app.after(_POLL_MS, poll)


def auto_check(app, style: Style) -> None:
    """启动时调一次：开关打开、距上次成功检查超过间隔，才真的去联网。"""
    s = app.gui_settings
    if not s.update_check_enabled:
        return
    if time.time() - s.update_last_check < AUTO_CHECK_INTERVAL_SECONDS:
        return
    start_check(app, style, manual=False)


def handle_result(app, style: Style, info: update_check.LatestInfo | None, *, manual: bool) -> None:
    if info is None:
        if manual:
            messagebox.showinfo(
                "检查更新",
                "没有连上更新服务器，暂时无法确认是否有新版本。\n\n"
                "可能是当前网络访问不了 Gitee/GitHub，请稍后再试，"
                "或到你获取本工具的网盘页面看看有没有新版本。",
                parent=app)
        return
    s = app.gui_settings
    s.update_last_check = int(time.time())          # 只在**读成功**时记：读不到下次启动还会再试
    if update_check.is_newer(info.version, __version__):
        if manual or info.version != s.update_ignored_version:
            _save(app)
            show_update_popup(app, style, info)
            return
    elif manual:
        messagebox.showinfo("检查更新", f"已经是最新版本（v{__version__}）。", parent=app)
    _save(app)


def build_link_box(app, parent, style: Style, downloads) -> tk.Text:
    """下载地址：一个只读文本框，每行 `名称：链接`。链接是蓝色划线、点一下在浏览器打开；
    整个框的文字都能选中复制。链接条数由 `latest.json` 决定，所以不用固定数量的按钮。

    点击在**松开鼠标**时才触发，并且有选中文字（说明用户在拖选）就不打开——
    从链接上起手拖选复制，不会误跳浏览器。"""
    f = style.family
    box = tk.Text(parent, wrap="char", bd=0, font=(f, 10), bg=style.bg, fg=style.fg,
                  height=min(max(len(downloads) + 1, 3), 9), highlightthickness=1,
                  highlightbackground="#c8c8c8", highlightcolor="#8a8a8a", padx=8, pady=6,
                  cursor="xterm")
    box.pack(fill="x", padx=18)
    for i, link in enumerate(downloads):
        tag = f"link{i}"
        box.insert("end", f"{link.name}：")
        start = box.index("end-1c")
        box.insert("end", link.url)
        box.tag_add(tag, start, box.index("end-1c"))
        box.tag_configure(tag, foreground=style.link, underline=True)
        box.tag_bind(tag, "<Enter>", lambda _e: box.configure(cursor="hand2"))
        box.tag_bind(tag, "<Leave>", lambda _e: box.configure(cursor="xterm"))

        def release(_event, url=link.url) -> None:
            if box.tag_ranges("sel"):
                return                                  # 刚才是在拖选，不是点击
            open_in_browser(app, url)

        box.tag_bind(tag, "<ButtonRelease-1>", release)
        if i < len(downloads) - 1:
            box.insert("end", "\n")
    box.bind("<Button-1>", lambda _e: box.focus_set(), add="+")   # 点一下就能 Ctrl+C
    keep = ("Left", "Right", "Up", "Down", "Prior", "Next", "Home", "End")
    box.bind("<Key>", lambda e: None if ((e.state & 0x4 and e.keysym.lower() in ("c", "a"))
                                         or e.keysym in keep) else "break")
    return box


def show_update_popup(app, style: Style, info: update_check.LatestInfo) -> None:
    """新版本弹窗：版本变化 + 更新日志 + 每个下载链接一个按钮（点了只是打开浏览器）。"""
    existing = getattr(app, "_update_win", None)
    if existing is not None and existing.winfo_exists():
        existing.destroy()
    win = tk.Toplevel(app)
    app._update_win = win
    win.title("发现新版本")
    win.configure(bg=style.bg)
    win.geometry("640x560")
    win.minsize(520, 420)
    try:
        from essay_tool.gui.win_place import center_on
        center_on(win, app, 640, 560)
    except Exception:                              # noqa: BLE001
        pass
    f = style.family
    tk.Label(win, text=f"发现新版本 v{info.version}", font=(f, 14, "bold"),
             bg=style.bg, fg=style.fg).pack(anchor="w", padx=18, pady=(16, 2))
    when = f"（{info.display_date()} 发布）" if info.date else ""
    tk.Label(win, text=f"你当前使用的是 v{__version__}。最新版本 v{info.version}{when}",
             font=(f, 10), bg=style.bg, fg=style.fg).pack(anchor="w", padx=18)
    tk.Label(win, text="更新内容", font=(f, 10, "bold"), bg=style.bg, fg=style.fg
             ).pack(anchor="w", padx=18, pady=(12, 4))

    box = tk.Frame(win, bg=style.bg)
    box.pack(fill="both", expand=True, padx=18)
    sb = tk.Scrollbar(box)
    sb.pack(side="right", fill="y")
    text = tk.Text(box, wrap="char", bd=0, bg=style.bg, fg=style.fg, font=(f, 10),
                   highlightthickness=0, yscrollcommand=sb.set, height=10)
    text.pack(fill="both", expand=True)
    sb.config(command=text.yview)
    text.insert("1.0", info.changelog or "（这个版本没有提供更新说明）")
    text.configure(state="disabled")

    tk.Label(win, text="下载地址（点击蓝色链接在浏览器里打开，也可以选中后复制；本程序不会自动下载）：\n"
                       "下载后运行安装包即可升级。",
             font=(f, 9), bg=style.bg, fg=style.fg, justify="left"
             ).pack(anchor="w", padx=18, pady=(10, 4))
    build_link_box(app, win, style, info.downloads)

    bottom = tk.Frame(win, bg=style.bg)
    bottom.pack(fill="x", padx=18, pady=(10, 14))

    def ignore() -> None:
        app.gui_settings.update_ignored_version = info.version
        _save(app)
        win.destroy()

    tk.Button(bottom, text="忽略此版本", font=(f, 10), padx=10, pady=3, command=ignore
              ).pack(side="left")
    tk.Button(bottom, text="关闭", font=(f, 10), padx=14, pady=3, command=win.destroy
              ).pack(side="right")
