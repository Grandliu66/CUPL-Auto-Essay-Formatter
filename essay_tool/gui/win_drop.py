"""让 Tk 主窗口接受从资源管理器拖进来的文件（Windows，纯 ctypes，不依赖第三方库）。

做法：对窗口调用 `DragAcceptFiles`，再把窗口过程换成我们的（子类化），
只拦 `WM_DROPFILES`，其余消息原样交回 Tk。回调里**不做任何耗时/弹窗的事**，
只把路径交给 `on_files`（由调用方自己 `after` 到主循环里处理），
否则模态对话框会把资源管理器那头的拖拽卡住。
"""
from __future__ import annotations

import ctypes
import sys
from ctypes import wintypes
from typing import Callable

WM_DROPFILES = 0x0233
WM_COPYGLOBALDATA = 0x0049
MSGFLT_ALLOW = 1
GWLP_WNDPROC = -4

_IS_WIN = sys.platform == "win32"


def _drop_paths(hdrop) -> list[str]:
    shell32 = ctypes.windll.shell32
    shell32.DragQueryFileW.argtypes = [wintypes.HANDLE, wintypes.UINT, wintypes.LPWSTR, wintypes.UINT]
    shell32.DragQueryFileW.restype = wintypes.UINT
    count = shell32.DragQueryFileW(hdrop, 0xFFFFFFFF, None, 0)
    paths = []
    for i in range(count):
        n = shell32.DragQueryFileW(hdrop, i, None, 0)
        buf = ctypes.create_unicode_buffer(n + 1)
        shell32.DragQueryFileW(hdrop, i, buf, n + 1)
        paths.append(buf.value)
    return paths


def enable_file_drop(window, on_files: Callable[[list[str]], None]) -> bool:
    """让 `window`（Tk 根窗口）接受拖入文件；成功返回 True。非 Windows 或失败返回 False。"""
    if not _IS_WIN:
        return False
    try:
        user32 = ctypes.windll.user32
        shell32 = ctypes.windll.shell32
        window.update_idletasks()
        # Tk 的 winfo_id 是内部子窗口，真正的顶层窗口是它的父窗口
        user32.GetParent.argtypes = [wintypes.HWND]
        user32.GetParent.restype = wintypes.HWND
        hwnd = user32.GetParent(window.winfo_id()) or window.winfo_id()

        LRESULT = ctypes.c_ssize_t
        WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT,
                                     wintypes.WPARAM, wintypes.LPARAM)
        is64 = ctypes.sizeof(ctypes.c_void_p) == 8
        setlong = user32.SetWindowLongPtrW if is64 else user32.SetWindowLongW
        setlong.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
        setlong.restype = ctypes.c_void_p
        user32.CallWindowProcW.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT,
                                           wintypes.WPARAM, wintypes.LPARAM]
        user32.CallWindowProcW.restype = LRESULT
        shell32.DragFinish.argtypes = [wintypes.HANDLE]

        old_proc = {"ptr": None}

        def _proc(h, msg, wparam, lparam):
            if msg == WM_DROPFILES:
                try:
                    paths = _drop_paths(wparam)
                finally:
                    shell32.DragFinish(wparam)
                if paths:
                    try:
                        on_files(paths)
                    except Exception:      # 回调出错不能带崩窗口过程
                        pass
                return 0
            return user32.CallWindowProcW(old_proc["ptr"], h, msg, wparam, lparam)

        cb = WNDPROC(_proc)
        old_proc["ptr"] = setlong(hwnd, GWLP_WNDPROC, ctypes.cast(cb, ctypes.c_void_p))
        if not old_proc["ptr"]:
            return False
        window._drop_proc_ref = cb            # 必须一直引用着，否则回调被回收会崩溃
        window._drop_hwnd = hwnd
        shell32.DragAcceptFiles(hwnd, True)
        # 程序以管理员运行时，普通权限的资源管理器发来的拖拽消息默认被 UIPI 拦掉
        try:
            for m in (WM_DROPFILES, WM_COPYGLOBALDATA):
                user32.ChangeWindowMessageFilterEx(hwnd, m, MSGFLT_ALLOW, None)
        except Exception:
            pass
        return True
    except Exception:
        return False
