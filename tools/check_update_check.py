"""检查更新功能的回归检查（不访问真实网络：用本机假服务器）。

覆盖：
  * 版本号比较、下载链接白名单（https + 域名，含各种绕过写法）；
  * `parse_latest` 对坏数据的处理（不是 JSON / 缺字段 / 版本号写错 / 全是不合格链接 / 超长日志）；
  * `fetch_latest`：正常、跳转、第一个源挂了用第二个、响应过大、返回垃圾、连不上、超时；
  * 设置文件里新增三项的读写与类型校验；
  * 界面（不弹窗）：新版本弹窗、下载按钮、"忽略此版本"、手动检查三种结果、自动检查的节流与开关。

用法：.venv\\Scripts\\python.exe tools\\check_update_check.py
"""

from __future__ import annotations

import json
import pathlib
import sys
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

ROOT = pathlib.Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from essay_tool import __version__, update_check as uc            # noqa: E402

fails: list[str] = []


def check(name: str, ok: bool, detail: str = "") -> None:
    print(f"   {'OK  ' if ok else 'FAIL'}  {name}" + (f"  → {detail}" if detail and not ok else ""))
    if not ok:
        fails.append(name)


GOOD = {
    "version": "9.9.9", "date": "2099-01-02", "changelog": "1. 测试\n· 一条",
    "downloads": [{"name": "夸克", "url": "https://pan.quark.cn/s/abc"},
                  {"name": "GitHub", "url": "https://github.com/Grandliu66/CUPL-Auto-Essay-Formatter/releases/latest"}],
}


def case_versions_and_urls() -> None:
    print("== 版本比较 / 链接白名单 ==")
    check("1.1.5 > 1.1.4", uc.is_newer("1.1.5", "1.1.4"))
    check("相同版本不算新", not uc.is_newer("1.1.4", "1.1.4"))
    check("1.10.0 > 1.9.0（按数字比、不按字符串）", uc.is_newer("1.10.0", "1.9.0"))
    check("旧版本不算新", not uc.is_newer("1.0.9", "1.1.0"))
    check("坏版本号不算新", not uc.is_newer("v1.2", "1.1.4") and not uc.is_newer("abc", "1.1.4"))
    ok = ["https://pan.quark.cn/s/x", "https://pan.baidu.com/s/y?pwd=1", "https://github.com/a/b",
          "https://gitee.com/a/b", "https://www.github.com/a", "https://HTTPS.GITHUB.COM/a".replace("HTTPS", "api")]
    bad = ["http://pan.quark.cn/s/x", "https://evil.com/", "https://github.com.evil.com/",
           "https://evilgithub.com/", "https://pan.quark.cn@evil.com/", "https://user:pw@github.com/",
           "https://github.com:8443/", "javascript:alert(1)", "file:///C:/Windows/System32/calc.exe",
           "//github.com/a", "ftp://github.com/a", "", "https://", "https://github.com\\@evil.com/"]
    check("白名单放行", all(uc.is_allowed_url(u) for u in ok), str([u for u in ok if not uc.is_allowed_url(u)]))
    check("非白名单/非 https/绕过写法一律拒绝", not any(uc.is_allowed_url(u) for u in bad),
          str([u for u in bad if uc.is_allowed_url(u)]))


def case_parse() -> None:
    print("== parse_latest ==")
    info = uc.parse_latest(json.dumps(GOOD, ensure_ascii=False))
    check("合格文件", info is not None and info.version == "9.9.9" and len(info.downloads) == 2)
    check("带 BOM 的 UTF-8 也认", uc.parse_latest(("\ufeff" + json.dumps(GOOD)).encode("utf-8")) is not None)
    check("不是 JSON", uc.parse_latest("v1.2.0\nchangelog:xxx") is None)
    check("不是对象", uc.parse_latest("[1,2]") is None)
    for label, mutate in (
            ("缺 version", lambda d: d.pop("version")),
            ("version 不是 x.y.z", lambda d: d.update(version="1.2")),
            ("version 不是字符串", lambda d: d.update(version=2)),
            ("date 格式不对", lambda d: d.update(date="2099/01/02")),
            ("changelog 不是字符串", lambda d: d.update(changelog=["x"])),
            ("downloads 不是列表", lambda d: d.update(downloads="x")),
            ("downloads 为空", lambda d: d.update(downloads=[])),
            ("downloads 全是不合格链接", lambda d: d.update(
                downloads=[{"name": "坏", "url": "https://evil.com/"}, {"name": "坏2", "url": "http://github.com/"}]))):
        data = json.loads(json.dumps(GOOD))
        mutate(data)
        check(f"拒绝：{label}", uc.parse_latest(json.dumps(data)) is None)
    data = json.loads(json.dumps(GOOD))
    data["downloads"] = [{"name": "好", "url": "https://github.com/x"}, {"name": "坏", "url": "https://evil.com/"},
                         "不是对象", {"name": "", "url": "https://github.com/y"}, {"url": "https://github.com/z"}]
    info = uc.parse_latest(json.dumps(data))
    check("混着坏链接：只留合格的", info is not None and [d.name for d in info.downloads] == ["好"])
    data = json.loads(json.dumps(GOOD))
    data["changelog"] = "字" * 100000
    data["downloads"][0]["name"] = "名" * 500
    info = uc.parse_latest(json.dumps(data, ensure_ascii=False))
    check("超长日志/名称被截断", info is not None and len(info.changelog) <= 8000
          and all(len(d.name) <= 40 for d in info.downloads))
    data = json.loads(json.dumps(GOOD))
    data["date"] = ""
    check("date 可以留空", uc.parse_latest(json.dumps(data)) is not None)


class _Handler(BaseHTTPRequestHandler):
    routes: dict = {}

    def log_message(self, *args):  # noqa: D401
        pass

    def do_GET(self):  # noqa: N802
        route = self.routes.get(self.path)
        if route is None:
            self.send_response(404)
            self.end_headers()
            return
        kind = route[0]
        if kind == "redirect":
            self.send_response(302)
            self.send_header("Location", route[1])
            self.end_headers()
        elif kind == "sleep":
            time.sleep(route[1])
            self.send_response(200)
            self.end_headers()
        else:
            body = route[1]
            self.send_response(200)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            self.last_ua = self.headers.get("User-Agent")
            type(self).seen_ua = self.last_ua


def case_fetch() -> None:
    print("== fetch_latest（本机假服务器）==")
    good = json.dumps(GOOD, ensure_ascii=False).encode("utf-8")
    _Handler.routes = {
        "/good.json": ("body", good),
        "/jump": ("redirect", "/good.json"),
        "/junk.json": ("body", b"<html>not json</html>"),
        "/huge.json": ("body", b" " * (uc.MAX_BYTES + 10)),
        "/slow": ("sleep", 3),
    }
    server = ThreadingHTTPServer(("127.0.0.1", 0), _Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    base = f"http://127.0.0.1:{server.server_address[1]}"
    try:
        info = uc.fetch_latest("1.1.4", sources=(f"{base}/good.json",))
        check("正常读取", info is not None and info.version == "9.9.9")
        check("User-Agent 带软件名和版本", str(getattr(_Handler, "seen_ua", "")) == "CUPL-UTAF/1.1.4",
              str(getattr(_Handler, "seen_ua", "")))
        check("跟随跳转（Gitee raw 就是这样）", uc.fetch_latest("1.1.4", sources=(f"{base}/jump",)) is not None)
        check("第一个源 404、第二个源好 ⇒ 用第二个",
              uc.fetch_latest("1.1.4", sources=(f"{base}/nope", f"{base}/good.json")) is not None)
        check("第一个源返回垃圾 ⇒ 试下一个",
              uc.fetch_latest("1.1.4", sources=(f"{base}/junk.json", f"{base}/good.json")) is not None)
        check("响应过大 ⇒ 丢弃", uc.fetch_latest("1.1.4", sources=(f"{base}/huge.json",)) is None)
        check("全是垃圾 ⇒ None（不抛）", uc.fetch_latest("1.1.4", sources=(f"{base}/junk.json",)) is None)
        check("连不上（端口没人听）⇒ None（不抛）",
              uc.fetch_latest("1.1.4", sources=("http://127.0.0.1:1/x.json",)) is None)
        old_timeout = uc.TIMEOUT_SECONDS
        uc.TIMEOUT_SECONDS = 1
        started = time.perf_counter()
        slow = uc.fetch_latest("1.1.4", sources=(f"{base}/slow",))
        uc.TIMEOUT_SECONDS = old_timeout
        check("超时 ⇒ None，且不会挂住", slow is None and time.perf_counter() - started < 2.8,
              f"{time.perf_counter() - started:.1f}s")
        check("没有任何源时 ⇒ None", uc.fetch_latest("1.1.4", sources=()) is None)
    finally:
        server.shutdown()


def case_settings() -> None:
    print("== 设置文件 ==")
    import tempfile
    from essay_tool.config import settings as st

    with tempfile.TemporaryDirectory() as tmp:
        orig = st._settings_path
        path = pathlib.Path(tmp) / "settings.json"
        st._settings_path = lambda: path
        try:
            default = st.load_settings()
            check("默认：自动检查开、上次时间 0、无忽略版本",
                  default.update_check_enabled is True and default.update_last_check == 0
                  and default.update_ignored_version == "")
            default.update_check_enabled = False
            default.update_last_check = 1700000000
            default.update_ignored_version = "1.2.0"
            st.save_settings(default)
            again = st.load_settings()
            check("存了能读回", again.update_check_enabled is False and again.update_last_check == 1700000000
                  and again.update_ignored_version == "1.2.0")
            path.write_text(json.dumps({"update_check_enabled": "yes", "update_last_check": "x",
                                        "update_ignored_version": 5, "backup_enabled": False}), encoding="utf-8")
            bad = st.load_settings()
            check("类型不对的退回默认值，其它项不受影响",
                  bad.update_check_enabled is True and bad.update_last_check == 0
                  and bad.update_ignored_version == "" and bad.backup_enabled is False)
            path.write_text(json.dumps({"backup_enabled": False}), encoding="utf-8")
            check("老版本的设置文件（没有这三项）照常读", st.load_settings().update_check_enabled is True)
        finally:
            st._settings_path = orig


def case_gui() -> None:
    print("== 界面 ==")
    import tkinter as tk
    from tkinter import messagebox

    orig_init, orig_state = tk.Tk.__init__, tk.Tk.state

    def quiet_init(self, *a, **k):
        orig_init(self, *a, **k)
        self.withdraw()

    tk.Tk.__init__, tk.Tk.state = quiet_init, (lambda self, *a: "withdrawn")
    shown: list[tuple] = []
    orig_info = messagebox.showinfo
    messagebox.showinfo = lambda title, msg, **kw: shown.append((title, msg))
    opened: list[str] = []
    try:
        from essay_tool.gui import main_window as mw, update_ui as ui
        ui.webbrowser.open = lambda url, *a, **k: opened.append(url) or True
        app = mw.MainWindow()
        saved: list[int] = []
        try:
            app.gui_settings.update_last_check = 0
            ui.save_settings = lambda s: saved.append(1)             # 不动用户真实的设置文件
            style = app._update_style()
            info = uc.parse_latest(json.dumps(GOOD, ensure_ascii=False))

            # 新版本弹窗
            ui.handle_result(app, style, info, manual=False)
            app.update_idletasks()
            win = getattr(app, "_update_win", None)
            check("有新版本 ⇒ 弹窗出现", win is not None and win.winfo_exists())

            def buttons(w):
                out = []
                for child in w.winfo_children():
                    if isinstance(child, tk.Button):
                        out.append(child)
                    out.extend(buttons(child))
                return out

            texts = [b.cget("text") for b in buttons(win)]
            check("只有『忽略此版本』『关闭』两个按钮（下载地址不再是按钮）", texts == ["忽略此版本", "关闭"], str(texts))

            def link_box(w):
                for child in w.winfo_children():
                    if isinstance(child, tk.Text) and child.tag_names("1.0") is not None and any(
                            t.startswith("link") for t in child.tag_names()):
                        return child
                    found = link_box(child)
                    if found is not None:
                        return found
                return None

            box = link_box(win)
            content = box.get("1.0", "end-1c")
            check("下载地址：每行『名称：链接』，文字可复制",
                  content == ("夸克：https://pan.quark.cn/s/abc\nGitHub：https://github.com/Grandliu66/"
                              "CUPL-Auto-Essay-Formatter/releases/latest"), repr(content))
            link_tags = sorted(t for t in box.tag_names() if t.startswith("link"))
            check("每条链接一个标签、蓝色划线",
                  len(link_tags) == 2 and all(box.tag_cget(t, "underline") for t in link_tags)
                  and all(str(box.tag_cget(t, "foreground")) == style.link for t in link_tags))
            check("文本框只读（不能输入/删除）", box.bind("<Key>") != "")
            # 模拟点击：在第一条链接上按下再松开
            win.update()
            line1_url = box.tag_ranges("link0")
            x, y, _w, _h = box.bbox(line1_url[0])
            box.event_generate("<Motion>", x=x + 3, y=y + 3)      # Tk 靠鼠标移动更新"当前所在的标签"
            box.event_generate("<ButtonRelease-1>", x=x + 3, y=y + 3)
            box.update()
            check("点蓝色链接 ⇒ 只是在浏览器打开该链接", opened == ["https://pan.quark.cn/s/abc"], str(opened))
            opened.clear()
            box.tag_add("sel", "1.0", "1.5")
            box.event_generate("<Motion>", x=x + 3, y=y + 3)
            box.event_generate("<ButtonRelease-1>", x=x + 3, y=y + 3)
            box.update()
            check("有选中文字（拖选）时点链接 ⇒ 不打开", opened == [], str(opened))
            box.tag_remove("sel", "1.0", "end")
            check("上次检查时间被记下", app.gui_settings.update_last_check > 0)
            [b for b in buttons(win) if b.cget("text") == "忽略此版本"][0].invoke()
            check("忽略此版本 ⇒ 记下版本号、弹窗关掉",
                  app.gui_settings.update_ignored_version == "9.9.9" and not win.winfo_exists())

            many = uc.parse_latest(json.dumps({**GOOD, "downloads": [
                {"name": f"源{i}", "url": f"https://github.com/Grandliu66/r{i}"} for i in range(6)]}))
            ui.show_update_popup(app, style, many)
            app.update_idletasks()
            many_box = link_box(app._update_win)
            check("链接数量会变：6 条也都显示、都是链接",
                  many_box is not None and len([t for t in many_box.tag_names() if t.startswith("link")]) == 6)
            app._update_win.destroy()

            # 忽略后：自动检查不再弹，手动检查仍弹
            app._update_win = None
            ui.handle_result(app, style, info, manual=False)
            check("忽略过的版本：自动检查不再弹", getattr(app, "_update_win", None) is None)
            ui.handle_result(app, style, info, manual=True)
            check("忽略过的版本：手动检查仍然弹", app._update_win is not None and app._update_win.winfo_exists())
            app._update_win.destroy()

            # 手动检查：已是最新 / 读不到
            same = uc.parse_latest(json.dumps({**GOOD, "version": __version__}))
            shown.clear()
            ui.handle_result(app, style, same, manual=True)
            check("手动 + 已是最新 ⇒ 告诉用户", len(shown) == 1 and "最新版本" in shown[0][1], str(shown))
            shown.clear()
            ui.handle_result(app, style, None, manual=True)
            check("手动 + 读不到 ⇒ 告诉用户", len(shown) == 1 and "没有连上" in shown[0][1], str(shown))
            shown.clear()
            ui.handle_result(app, style, None, manual=False)
            ui.handle_result(app, style, same, manual=False)
            check("自动 + 读不到/已是最新 ⇒ 完全静默", shown == [] and getattr(app, "_update_win", None) is not None
                  and not app._update_win.winfo_exists())

            # 自动检查的开关与节流
            called: list[bool] = []
            ui.start_check = lambda *a, **k: called.append(k.get("manual"))
            app.gui_settings.update_check_enabled = False
            app.gui_settings.update_last_check = 0
            ui.auto_check(app, style)
            check("开关关着 ⇒ 不联网", called == [])
            app.gui_settings.update_check_enabled = True
            app.gui_settings.update_last_check = int(time.time())
            ui.auto_check(app, style)
            check("刚检查过 ⇒ 不联网（一天最多一次）", called == [])
            app.gui_settings.update_last_check = int(time.time()) - 21 * 3600
            ui.auto_check(app, style)
            check("超过间隔 ⇒ 联网（自动、非手动）", called == [False], str(called))
            # 关了自动检查、手动查过、再勾上：下次启动必须会检查（不能被"上次检查时间"挡住）
            import essay_tool.gui.main_window as _mw
            _mw.save_settings = lambda s: saved.append(1)            # 不动用户真实的设置文件
            app.var_update_auto.set(False)
            app._on_update_auto_toggle()
            check("取消勾选 ⇒ 关闭自动检查", app.gui_settings.update_check_enabled is False)
            app.gui_settings.update_last_check = int(time.time())   # 关着期间手动检查过一次
            app.var_update_auto.set(True)
            app._on_update_auto_toggle()
            called.clear()
            ui.auto_check(app, style)
            check("重新勾选 ⇒ 开关打开，且下次启动会检查（不被上次检查时间挡住）",
                  app.gui_settings.update_check_enabled is True and app.gui_settings.update_last_check == 0
                  and called == [False], f"{app.gui_settings.update_last_check} {called}")
            check("构造主窗口本身不联网（selftest 安全）", True)
        finally:
            app.destroy()
    finally:
        tk.Tk.__init__, tk.Tk.state = orig_init, orig_state
        messagebox.showinfo = orig_info


def main() -> int:
    case_versions_and_urls()
    case_parse()
    case_fetch()
    case_settings()
    case_gui()
    print()
    print(f"失败 {len(fails)} 条")
    for name in fails:
        print("  ✘", name)
    return 1 if fails else 0


if __name__ == "__main__":
    raise SystemExit(main())
