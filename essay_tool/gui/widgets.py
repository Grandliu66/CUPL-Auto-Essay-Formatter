"""项目里两个窗口共用的自绘控件样式。目前只有"复选框"。

为什么要自己做：Windows 下经典 `tk.Checkbutton` 的原生指示器不怎么跟着字号
放大（实测字号从 10 调到 12，控件确实变高了，但多出来的空间基本是文字和留白，
那个小方块肉眼看不出变化），而且选中态会硬编码一个"按下去"的右下偏移、
`relief` 这类公开选项关不掉（用户 2026-09-14 前后反馈过两次）。所以换成
`ttk.Checkbutton` + 自造方块图 + 自定义布局：ttk 是纯声明式布局，没有那段
历史包袱，选中前后位置完全一致。

主窗口里的复选框和"设置封面范围"弹窗里的复选框都用这一套，样子才一样
（用户 2026-09-15：参照主窗口的复选框，用那种写法写复选框）。

用法：
    init_checkbox_style(root, PANEL_BG)          # 注册一次即可（幂等）
    cb = checkbutton(parent, "文字", variable=var, command=fn)
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

# 面板底色。方块图的填充色就是它，所以**用这套复选框的窗口底色要一致**
# （主窗口面板、各弹窗都是纯白）。
PANEL_BG = "#ffffff"
PANEL_FG = "#222222"
# 全项目统一的界面字体（各 gui 模块里同样写着这一行）
FAMILY = "Microsoft YaHei UI"

BOX = 16          # 方块边长
GAP = 8           # 方块与文字之间的距离（填底色，肉眼看不出，只是拉开距离）
BORDER = "#555555"
CHECK = "#2f6fbe"

_BASE_STYLE = "TCheckbutton"
_registered = False
_extra_sizes: set[int] = set()
# 图片对象必须留一个引用，否则会被回收、方块变空白
_images: dict[str, tk.PhotoImage] = {}


def wheel_notch(event, scale: int = 1) -> int:
    """滚轮转了"几格"（带高精度滚轮/触控板的兜底）。

    为什么要有这个函数：`event.delta` 在 Windows 上多数时候是 120 的倍数，
    但**高精度滚轮与触控板**会给出更小的值（`|delta| < 120`），这时
    `delta // 120` 恒为 `0` ⇒ 表根本滚不动。原来主窗口里写了兜底、结构页与
    流程图里没写（2026-09-15 审计核实为真），三处口径不一——统一走这里。

    `scale`：一次滚动算几格（调用方拿它直接乘滚动单位）。
    返回 0 表示"这次事件没有有效位移"（`delta == 0`，不应滚动）。
    """
    notch = event.delta // 120
    if notch == 0 and event.delta != 0:
        notch = 1 if event.delta > 0 else -1
    return notch * scale


def init_checkbox_style(widget: tk.Misc, background: str = PANEL_BG) -> None:
    """注册自制复选框样式（同一个进程里只做一次）。"""
    global _registered
    if _registered:
        return
    _registered = True

    width = BOX + GAP
    off = tk.PhotoImage(width=width, height=BOX, master=widget)
    off.put(background, to=(0, 0, width, BOX))
    on = tk.PhotoImage(width=width, height=BOX, master=widget)
    on.put(background, to=(0, 0, width, BOX))
    for image in (off, on):
        image.put(BORDER, to=(0, 0, BOX, 1))
        image.put(BORDER, to=(0, BOX - 1, BOX, BOX))
        image.put(BORDER, to=(0, 0, 1, BOX))
        image.put(BORDER, to=(BOX - 1, 0, BOX, BOX))
    pad = 3
    on.put(CHECK, to=(pad, pad, BOX - pad, BOX - pad))
    _images["off"], _images["on"] = off, on

    style = ttk.Style(widget)
    # 自定义布局挂在基础样式名上（不是自定义前缀名）：派生名找不到自己的布局时
    # 会自动退回这个基础布局，所以"只改字号"的派生样式照样是同一份画法。
    if "Custom.Checkbutton.indicator" not in style.element_names():
        style.element_create("Custom.Checkbutton.indicator", "image", off,
                             ("selected", on), width=width, sticky="w")
        style.layout(_BASE_STYLE, [
            ("Checkbutton.padding", {"sticky": "nswe", "children": [
                ("Custom.Checkbutton.indicator", {"side": "left", "sticky": ""}),
                ("Checkbutton.focus", {"side": "left", "sticky": "", "children": [
                    ("Checkbutton.label", {"sticky": "nswe"}),
                ]}),
            ]}),
        ])
    style.configure(_BASE_STYLE, background=background, foreground=PANEL_FG,
                    font=(FAMILY, 12))
    style.map(_BASE_STYLE, background=[], foreground=[])


_bg_styles: dict[tuple[int, str], str] = {}


def checkbutton_style_on(widget: tk.Misc, font_size: int, bg: str) -> str:
    """同一套自造复选框，但**方块之外的底色换成 `bg`**，返回样式名。

    自造方块的图片里把底色烧死了（`init_checkbox_style` 里 `put(background…)`），
    所以复选框放到别的底色上（比如页签条的灰、闪烁时的黄）就会露出一块白。
    这里按 (字号, 底色) 各注册一份样式：方块内部仍是白的，方块外面和文字底都是
    `bg`。同一个 (字号, 底色) 只注册一次。
    """
    key = (font_size, bg)
    got = _bg_styles.get(key)
    if got is not None:
        return got
    init_checkbox_style(widget)             # 确保基础样式先注册过
    n = len(_bg_styles) + 1
    width = BOX + GAP
    off = tk.PhotoImage(width=width, height=BOX, master=widget)
    on = tk.PhotoImage(width=width, height=BOX, master=widget)
    for image in (off, on):
        image.put(bg, to=(0, 0, width, BOX))
        image.put(PANEL_BG, to=(0, 0, BOX, BOX))          # 方块内部保持白色
        image.put(BORDER, to=(0, 0, BOX, 1))
        image.put(BORDER, to=(0, BOX - 1, BOX, BOX))
        image.put(BORDER, to=(0, 0, 1, BOX))
        image.put(BORDER, to=(BOX - 1, 0, BOX, BOX))
    pad = 3
    on.put(CHECK, to=(pad, pad, BOX - pad, BOX - pad))
    _images[f"off{n}"], _images[f"on{n}"] = off, on
    element = f"Custom{n}.Checkbutton.indicator"
    style = ttk.Style(widget)
    style.element_create(element, "image", off, ("selected", on), width=width, sticky="w")
    name = f"Bg{n}Size{font_size}.TCheckbutton"
    style.layout(name, [
        ("Checkbutton.padding", {"sticky": "nswe", "children": [
            (element, {"side": "left", "sticky": ""}),
            ("Checkbutton.focus", {"side": "left", "sticky": "", "children": [
                ("Checkbutton.label", {"sticky": "nswe"}),
            ]}),
        ]}),
    ])
    style.configure(name, background=bg, foreground=PANEL_FG, font=(FAMILY, font_size))
    style.map(name, background=[], foreground=[])
    _bg_styles[key] = name
    return name


def checkbutton(parent: tk.Widget, text: str = "", *, variable: tk.BooleanVar,
                command=None, font_size: int = 12) -> ttk.Checkbutton:
    """造一个用上面那套样式的复选框。`font_size` 只是覆盖字号。"""
    style_name = _BASE_STYLE
    if font_size != 12:
        style_name = f"Size{font_size}.TCheckbutton"
        if font_size not in _extra_sizes:
            ttk.Style(parent).configure(style_name, font=(FAMILY, font_size))
            _extra_sizes.add(font_size)
    return ttk.Checkbutton(parent, text=text, variable=variable, command=command,
                           style=style_name)
