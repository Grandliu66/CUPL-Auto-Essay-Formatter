"""二级窗口摆放：一律摆在主窗口正中央（两个中心点重合）。

用户 2026-09-15 定的规矩：

    对整个项目的任何二级窗口，不允许在屏幕左上角弹出，改为在主窗口中心处弹出
    （两个中心点重合，而不是窗口左上角和中心点重合）。

Tk 的默认行为是把新窗口摆在屏幕左上角附近；项目里原来还有一处
"贴父窗口左上角 + (140, 40)"的写法。这两种都是这条规矩要禁掉的。所以每个
`tk.Toplevel` 建完、尺寸定下来之后，都要调一次 `center_on()`。
"""

from __future__ import annotations

import tkinter as tk


def center_on(
    window: tk.Misc,
    parent: tk.Misc | None = None,
    width: int | None = None,
    height: int | None = None,
) -> None:
    """把 `window` 摆到 `parent` 正中央。

    `width` / `height`：窗口尺寸——**显式设过 `geometry("WxH")` 的就照实传**，
    因为窗口还没映射时 `winfo_width()` 读出来是 1，量不到真实尺寸；不传就按
    实测/请求尺寸算。

    `parent` 没给、或者它还没映射（比如"连主窗口都建不起来"时传的是一个干净的
    `tk.Tk()`）→ 退化成屏幕居中。摆完再夹一下边界，别把窗口推到屏幕外面去。
    """
    try:
        window.update_idletasks()
        if width is None or height is None:
            measured_w, measured_h = window.winfo_width(), window.winfo_height()
            if measured_w <= 1 or measured_h <= 1:
                measured_w, measured_h = window.winfo_reqwidth(), window.winfo_reqheight()
            width = width or measured_w
            height = height or measured_h

        if parent is not None and parent.winfo_ismapped() and parent.winfo_width() > 1:
            x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
            y = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
            # ⚠️ `winfo_rootx()` 量的是**客户区**原点，而 `geometry("+x+y")` 设的是
            # **整个窗口**（含标题栏/边框）的位置，两者差着窗口装饰那一圈
            # （本机实测 8×31 像素）。不减掉这一圈的话，子窗口会整体右下偏一点，
            # 两个中心点并不真的重合（实测偏移正好是 8 和 31）。
            x -= parent.winfo_rootx() - parent.winfo_x()
            y -= parent.winfo_rooty() - parent.winfo_y()
        else:
            x = (window.winfo_screenwidth() - width) // 2
            y = (window.winfo_screenheight() - height) // 2

        x = max(0, min(x, window.winfo_screenwidth() - width))
        y = max(0, min(y, window.winfo_screenheight() - height))
        window.geometry(f"+{int(x)}+{int(y)}")
    except Exception:  # noqa: BLE001 —— 摆位置失败不影响功能，不值得往上抛
        pass
