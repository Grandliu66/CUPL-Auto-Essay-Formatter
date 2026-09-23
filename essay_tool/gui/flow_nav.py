"""流程图导航面板（右侧"使用向导"）：Canvas 自绘的可点选、可缩放流程图。

2026-09-19 按用户的新流程图（`reference/2.png`）重建，取代 2026-09-13 的旧版
（旧版已归档在 `work/_obsolete_流程图v1-20260919/`）。对外接口不变：
`FlowPanel(master, on_activate=…, on_zoom=…)`，点可点的节点回调
`on_activate(node_id, label)`；`MainWindow._on_flow` 按节点 id 做跳转。

图的内容与口径：
- 论文排版支线：导入你的论文 → 选择序号体系 → 想如何排版（导入 Excel 并应用 / 应用既有预设 /
  自行设置）→ 识别论文并调整识别结果 → 是否要进一步调整 → 导出 → 结果满意吗 → 是否保存预设；
  空白模板支线独立（应用既有预设〔可选〕→ 调整模板结构 → 页面/封面/格式 → 导出）；
- "结果满意吗？否"**回到"识别论文并调整识别结果"**（用户裁决：那样既不多问一遍、又能回到
  结构调整）；"内置预设"叫"既有预设"；协议勾选、序号措辞、封面提醒三处用户裁决"忽略"；
- **Ctrl＋滚轮缩放**（以鼠标所在点为中心）、普通滚轮上下滚、Shift＋滚轮左右滚、空白处双击
  复位到"适合宽度"（默认铺满面板宽度，宽度锁定：用户 2026-09-13 的口径）；
- 版面按文字大小量身排的（节点尺寸贴文字、行距收紧、参数框窄而靠左）；
- 口径沿用：只有黑、白、蓝三色，可点节点的文字蓝色＋下划线、线更粗；分支线上的文字不可点；
  没有步骤状态、没有不可用节点。

接线对照（`MainWindow._on_flow`；标签页序号＝`_goto_hint` 的第一个参数：0 文件 / 1 页面 /
2 格式 / 3 结构 / 4 封面）：
    import      → 0，闪 btn_import
    scheme      → 0，闪"原文标题序号"那一行（`row_scheme`）
    excel       → 0，闪 btn_import_preset 和各行「应用」（导入后还要点某一行的「应用」）
    preset      → 0，闪各行的「应用」（`_apply_btns`）
    tpl_preset  → 同 preset
    recog       → 3
    tpl_struct  → 3（空白模板模式的结构页）
    p_page      → 1
    p_fmt       → 2
    p_cover     → 4，闪 cover_sw_box
    export      → 0，闪 btn_exp_blank / btn_exp_doc
    save        → 0，闪 btn_save_preset

用法（单独跑）：
    .venv\\Scripts\\python.exe essay_tool\\gui\\flow_nav.py
"""

from __future__ import annotations

import pathlib
import sys
import tkinter as tk
import tkinter.font as tkfont
from collections.abc import Callable

# 直接 `python essay_tool\gui\flow_nav.py` 单独跑时，sys.path[0] 是脚本所在目录、
# 找不到 `essay_tool` 包——把项目根补进去（被主窗口 import 时这一步无害）。
_ROOT = pathlib.Path(__file__).resolve().parents[2]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from essay_tool.gui.widgets import wheel_notch  # noqa: E402

FAMILY = "Microsoft YaHei UI"
C_LINK = "#0563c1"              # 可点文字：Word 超链接同款蓝

_NODE_FILL = "#ffffff"
_NODE_LINE = "#000000"
_LW_CLICK = 1.6                 # 可点的动作节点
_LW_STATIC = 1.1                # 起止/判断这些不可点的形状

# 设计坐标（z=1 时的像素；12px 字，一个汉字宽 12）。整张图是照用户草图 `reference/2.png`
# 重新排布的（2026-09-19 用户："参数框向右突出太多导致整个歪，调窄向左收；大部分框空白
# 太大，收小一点让图更紧凑"）：节点按文字大小量身定尺寸，行距收紧，参数框缩窄并紧贴
# 判断3；走线的拓扑和草图完全一致。
AXIS_X = 226                    # 流程图主干（"开始使用"那一列）的设计横坐标
CANVAS_W = 504                  # 图内容的设计宽高（左右边距对称）
CANVAS_H = 990
# "适合宽度"的参照宽度：面板宽度 ÷ FIT_W ＝ 默认倍率。取图内容宽度本身（含左右边距），
# 也就是默认让图刚好铺满面板宽度（用户 2026-09-19："默认放大一点，现在有点浪费空间"；
# 中间一版取 650 时图在宽面板里两边留白、显得偏小）。装不下的高度靠滚动条，想一眼看全
# 用 Ctrl＋滚轮缩小。
FIT_W = CANVAS_W

# 缩放：z = 适合宽度的倍率 × 用户倍率。下限保证字号（≥9px）不至于溢出节点框。
Z_MIN, Z_MAX = 0.7, 2.6
ZOOM_STEP = 1.12                # Ctrl＋滚轮每一格
_DRAG_SLOP = 5                  # 左键按下后移动超过这么多像素才算"拖动"（否则仍是点击）

# ------------------------------------------------------------------ 图的数据
# 各列的中轴：主干 226；左下列（判断3/导出/判断4/判断5）138；Excel 列 104；空白模板列
# （右支线与参数框）372。
# id / 形状 / 文字 / 中心 cx,cy / 宽高 w,h / 是否可点 click / 点了要做什么 action
NODES = [
    dict(id="start", kind="stadium", text="开始使用", cx=226, cy=20, w=104, h=32),
    dict(id="d1", kind="diamond", text="想给论文排版\n还是导出模板?", cx=226, cy=96, w=200, h=70),
    dict(id="import", kind="io", text="导入你的论文", cx=226, cy=190, w=130, h=36,
         click=True, action="导入论文"),
    dict(id="scheme", kind="rect", text="选择论文使用\n的序号体系", cx=226, cy=258, w=110, h=48,
         click=True, action="选择原文标题序号（文件页）"),
    dict(id="d2", kind="diamond", text="想如何排版?", cx=226, cy=334, w=130, h=52),
    dict(id="excel", kind="rect", text="从已有Excel\n导入配置\n并应用", cx=104, cy=433, w=100, h=58,
         click=True, action="从 Excel 导入配置并应用"),
    dict(id="preset", kind="rect", text="应用既有预设", cx=226, cy=433, w=112, h=44,
         click=True, action="应用既有预设"),
    dict(id="recog", kind="rect", text="识别论文\n并调整识别结果", cx=226, cy=542, w=122, h=48,
         click=True, action="识别并调整识别结果（结构页）"),
    dict(id="d3", kind="diamond", text="是否要进一步调整?", cx=138, cy=640, w=168, h=56),
    dict(id="export", kind="io", text="导出", cx=138, cy=722, w=96, h=36,
         click=True, action="导出"),
    dict(id="d4", kind="diamond", text="结果满意吗?", cx=138, cy=788, w=130, h=48),
    dict(id="d5", kind="diamond", text="是否保存预设?", cx=138, cy=866, w=150, h=52),
    dict(id="save", kind="rect", text="保存预设", cx=302, cy=866, w=90, h=34,
         click=True, action="把当前设置保存为新预设"),
    dict(id="end", kind="stadium", text="结束使用", cx=226, cy=950, w=104, h=32),
    # —— 空白模板支线（右侧）
    dict(id="tpl_preset", kind="rect", text="应用既有预设\n(可选)", cx=372, cy=426, w=100, h=44,
         click=True, action="应用既有预设（可选）"),
    dict(id="tpl_struct", kind="rect", text="调整模板结构", cx=372, cy=484, w=100, h=32,
         click=True, action="调整模板结构（结构页）"),
    # —— 参数框里的三项
    dict(id="p_page", kind="rect", text="页面：纸张/页边距/页眉页脚", cx=372, cy=597, w=184, h=30,
         click=True, action="页面设置"),
    dict(id="p_cover", kind="rect", text="封面：自动封面/填写信息", cx=372, cy=635, w=184, h=30,
         click=True, action="封面设置"),
    dict(id="p_fmt", kind="rect", text="格式：字体/段落/引注/部件标签", cx=372, cy=673, w=184, h=30,
         click=True, action="格式设置"),
]

# 参数框（三个子项的容器，纯装饰）：x0,y0,x1,y1
CONTAINER = (270, 572, 474, 698)

# 连线：pts 是折线点列；arrow=False 不带箭头（并线）；label/lx/ly 是压在线上的文字
EDGES = [
    dict(pts=[(226, 36), (226, 61)]),                                            # 开始 → 判断1
    dict(pts=[(226, 131), (226, 172)], label="给论文排版", lx=226, ly=151),        # 判断1 → 导入
    dict(pts=[(326, 96), (372, 96), (372, 404)], label="导出空白模板", lx=372, ly=250),
    dict(pts=[(226, 208), (226, 234)]),                                          # 导入 → 序号体系
    dict(pts=[(226, 282), (226, 308)]),                                          # 序号体系 → 判断2
    dict(pts=[(161, 334), (104, 334), (104, 404)], label="导入Excel", lx=104, ly=385),
    dict(pts=[(226, 360), (226, 411)], label="应用预设", lx=226, ly=385),
    dict(pts=[(291, 334), (306, 334), (306, 490), (226, 490)], arrow=False,      # 自行设置 → 并入主干
         label="自行设置", lx=306, ly=385),
    dict(pts=[(104, 462), (104, 490), (226, 490)], arrow=False),                 # Excel → 并入主干
    dict(pts=[(226, 455), (226, 518)]),                                          # 预设/并线 → 识别
    dict(pts=[(226, 566), (226, 590), (138, 590), (138, 612)]),                  # 识别 → 判断3
    dict(pts=[(138, 668), (138, 704)], label="否", lx=138, ly=686),               # 判断3 否 → 导出
    dict(pts=[(222, 640), (270, 640)], label="是", lx=246, ly=640),               # 判断3 是 → 参数框
    dict(pts=[(372, 448), (372, 468)]),                                          # 空白模板：预设 → 调整结构
    dict(pts=[(372, 500), (372, 572)]),                                          # 调整结构 → 参数框
    dict(pts=[(372, 698), (372, 722), (182, 722)]),                              # 参数框 → 导出
    dict(pts=[(138, 740), (138, 764)]),                                          # 导出 → 判断4
    # 不满意：回到"识别论文并调整识别结果"（左边缘中点进）
    dict(pts=[(73, 788), (30, 788), (30, 542), (165, 542)], label="否", lx=52, ly=788),
    dict(pts=[(138, 812), (138, 840)], label="是", lx=138, ly=826),               # 判断4 是 → 判断5
    dict(pts=[(213, 866), (257, 866)], label="是", lx=235, ly=866),               # 判断5 是 → 保存
    dict(pts=[(138, 892), (138, 912), (226, 912), (226, 934)], label="否", lx=182, ly=912),
    dict(pts=[(302, 883), (302, 912), (226, 912)], arrow=False),                 # 保存 → 并入结束
]

NODE_LABEL = {nd["id"]: nd["text"].replace("\n", "") for nd in NODES}
NODE_ACTION = {nd["id"]: nd["action"] for nd in NODES if nd.get("click")}


def _round_rect(c: tk.Canvas, x0, y0, x1, y1, r, **kw):
    """圆角矩形（polygon+smooth 的经典画法）；r = 高度一半就是胶囊形。"""
    pts = [
        x0 + r, y0, x1 - r, y0, x1, y0, x1, y0 + r,
        x1, y1 - r, x1, y1, x1 - r, y1, x0 + r, y1,
        x0, y1, x0, y1 - r, x0, y0 + r, x0, y0,
    ]
    return c.create_polygon(pts, smooth=True, **kw)


class FlowPanel(tk.Frame):
    """可点选、可缩放的流程图导航面板（内嵌用；单独跑请用 `main()`）。

    - `on_activate(node_id, label)`：点了可点的节点；宿主自己决定跳哪个页签、闪哪个
      按钮（`NODE_ACTION` 给了每个节点"该做什么"的说明）；
    - `on_zoom(percent)`：缩放变了（可选，给宿主显示"缩放 120%"之类）。
    """

    def __init__(self, master: tk.Misc, on_activate: Callable[[str, str], None] | None = None,
                 on_zoom: Callable[[int], None] | None = None) -> None:
        super().__init__(master, bg="#ffffff")
        self.on_activate = on_activate
        self.on_zoom = on_zoom
        self._hits: list[tuple[float, float, float, float, str]] = []

        self.uz = 1.0               # 用户缩放倍率（Ctrl＋滚轮改它；1.0＝适合宽度）
        self.z = 1.0                # 实际倍率 = 适合宽度倍率 × uz（夹在 Z_MIN~Z_MAX）
        self.ox = 0.0               # 图左边距（比视口窄时居中）
        self._font_cache: dict[int, dict] = {}
        self._drawn: tuple | None = None

        self.canvas = tk.Canvas(self, width=CANVAS_W, height=CANVAS_H, bg="#ffffff",
                                highlightthickness=0)
        self.vsb = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.hsb = tk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.vsb.set, xscrollcommand=self._on_xscroll)
        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self._hsb_shown = False     # 横向滚动条只在图比视口宽（放大后）时才出现
        # 宿主想知道"图现在画在哪"（比如让标题对齐中轴）：重画/滚动之后回调
        self.on_layout: Callable[[], None] | None = None
        self._press: tuple[int, int] | None = None   # 左键按下的位置（视口坐标）
        self._press_hit: str | None = None           # 按下时压在哪个可点节点上
        self._dragged = False                        # 这一次按下之后是不是拖动过

        # 左键：按下记位置、拖动平移画布、松开时**没拖动过**才算点击（点节点跳转）。
        # 点击不能放在按下的瞬间——那样一按节点就跳走，没法从节点上开始拖。
        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.canvas.bind("<Double-Button-1>", self._on_double)
        self.canvas.bind("<Motion>", self._on_motion)
        self.canvas.bind("<Configure>", lambda _e: self._relayout())
        self.canvas.bind("<MouseWheel>", self._on_wheel)
        self.canvas.bind("<Shift-MouseWheel>", self._on_wheel_h)
        self.canvas.bind("<Control-MouseWheel>", self._on_wheel_zoom)

        self._draw()

    # ------------------------------------------------------------ 中轴位置（给宿主对齐用）

    def axis_view_x(self) -> float:
        """流程图主干（"开始使用"那一列）当前在**画布视口**里的横坐标（像素）。"""
        return self.ox + AXIS_X * self.z - self.canvas.canvasx(0)

    def _on_xscroll(self, first, last) -> None:
        self.hsb.set(first, last)
        if self.on_layout:
            self.on_layout()

    def _notify_layout(self) -> None:
        if self.on_layout:
            self.on_layout()

    # ------------------------------------------------------------ 缩放

    def zoom_percent(self) -> int:
        """当前显示倍率（相对设计尺寸），给状态栏用。"""
        return round(self.z * 100)

    def _fit_z(self) -> float:
        """"适合宽度"的倍率：内容刚好铺满视口宽度（夹在 Z_MIN~Z_MAX）。"""
        return max(Z_MIN, min(max(self.canvas.winfo_width(), 120) / FIT_W, Z_MAX))

    def _target_z(self) -> float:
        return max(Z_MIN, min(self._fit_z() * self.uz, Z_MAX))

    def _relayout(self) -> None:
        """视口尺寸变了（窗口缩放/分隔条拖动）：按新宽度重算倍率并重画。"""
        self._draw()

    def _reset_zoom(self) -> None:
        self.uz = 1.0
        self._draw()

    def _on_double(self, e) -> None:
        """在空白处双击 ＝ 复位到"适合宽度"。

        点在可点节点上时**仍按普通点击处理**：Tk 里连点两下（半秒内）只触发
        `<Double-Button-1>`、不再触发 `<Button-1>`，不接住的话，用户快速连点两个
        节点，第二下会被吞掉。"""
        self._begin_press(e)          # 第二下按下不会再走 <ButtonPress-1>，这里补记
        if self._press_hit is None:
            self._reset_zoom()
        # 压在节点上的那一下，交给松开时的 `_on_release` 当点击处理

    def _on_wheel_zoom(self, e):
        """Ctrl＋滚轮：以鼠标所在点为中心缩放。"""
        notch = wheel_notch(e)
        if notch == 0:
            return "break"
        c = self.canvas
        old_z, old_ox = self.z, self.ox
        # 鼠标下面那一点的设计坐标（缩放前后要钉在同一个屏幕位置）
        dx = (c.canvasx(e.x) - old_ox) / old_z
        dy = c.canvasy(e.y) / old_z
        fit = self._fit_z()
        new_uz = self.uz * (ZOOM_STEP ** notch)
        new_z = max(Z_MIN, min(fit * new_uz, Z_MAX))
        if abs(new_z - old_z) < 1e-4:                 # 已到上下限，别让 uz 越跑越远
            return "break"
        self.uz = new_z / fit
        self._draw()
        # 让 (dx, dy) 重新落回鼠标位置：目标左上角 = 该点新位置 − 鼠标在视口里的偏移
        total_w = max(CANVAS_W * self.z, c.winfo_width())
        total_h = max(CANVAS_H * self.z, c.winfo_height())
        left = dx * self.z + self.ox - e.x
        top = dy * self.z - e.y
        c.xview_moveto(max(0.0, left) / total_w)
        c.yview_moveto(max(0.0, top) / total_h)
        return "break"

    def _on_wheel(self, e):
        self.canvas.yview_scroll(-2 * wheel_notch(e), "units")
        return "break"

    def _on_wheel_h(self, e):
        self.canvas.xview_scroll(-2 * wheel_notch(e), "units")
        return "break"

    # ------------------------------------------------------------ 绘制

    def _fonts(self) -> dict:
        """按当前缩放取像素级字体（缓存）。"""
        px = max(9, round(12 * self.z))
        got = self._font_cache.get(px)
        if got is None:
            got = {
                "node": tkfont.Font(family=FAMILY, size=-px),
                "node_u": tkfont.Font(family=FAMILY, size=-px, underline=True),
                "small": tkfont.Font(family=FAMILY, size=-px),
            }
            self._font_cache[px] = got
        return got

    def _draw(self) -> None:
        c = self.canvas
        z = self._target_z()
        view_w = max(c.winfo_width(), 120)
        content_w = CANVAS_W * z
        ox = max(0.0, (view_w - content_w) / 2)        # 比视口窄就居中，宽了靠横向滚动条
        key = (round(z, 4), round(ox, 1))
        if key == self._drawn and c.find_all():
            self._notify_layout()
            return
        self._drawn = key
        self.z, self.ox = z, ox
        need_h = content_w > view_w + 1
        if need_h and not self._hsb_shown:
            self.hsb.pack(side="bottom", fill="x", before=self.canvas)
            self._hsb_shown = True
        elif not need_h and self._hsb_shown:
            self.hsb.pack_forget()
            self._hsb_shown = False
        c.delete("all")
        self._hits = []
        c.configure(scrollregion=(0, 0, max(content_w, view_w), CANVAS_H * z))

        def pt(p):
            return (p[0] * z + ox, p[1] * z)

        for e in EDGES:                                # 先画连线（节点和标签盖在上面）
            flat = [v for p in e["pts"] for v in pt(p)]
            c.create_line(*flat, fill="#000000", width=max(1.0, 1.2 * z),
                          arrow=tk.LAST if e.get("arrow", True) else tk.NONE,
                          arrowshape=(9 * z, 11 * z, 4 * z), joinstyle="round")

        x0, y0, x1, y1 = CONTAINER                     # 参数框容器（草图里是直角）
        c.create_rectangle(x0 * z + ox, y0 * z, x1 * z + ox, y1 * z,
                           fill="#ffffff", outline="#000000", width=max(1.0, 1.6 * z))

        for nd in NODES:
            self._draw_node(nd)

        fonts = self._fonts()                          # 线上的注释文字：黑字、不可点
        for e in EDGES:
            if e.get("label"):
                self._text_on_line(e["lx"] * z + ox, e["ly"] * z, e["label"], fonts["small"])

        if self.on_zoom:
            self.on_zoom(self.zoom_percent())
        self._notify_layout()

    def _text_on_line(self, x, y, text: str, font) -> None:
        """画线上注释：先写字，再在字下垫白底，让穿过的线在文字处断开。"""
        c = self.canvas
        tid = c.create_text(x, y, text=text, anchor="center", font=font,
                            fill="#000000", justify="center")
        x0, y0, x1, y1 = c.bbox(tid)
        c.create_rectangle(x0 - 3, y0 - 2, x1 + 3, y1 + 2, fill="#ffffff", outline="")
        c.tag_raise(tid)

    def _draw_node(self, nd: dict) -> None:
        c = self.canvas
        z, ox = self.z, self.ox
        cx, cy = nd["cx"] * z + ox, nd["cy"] * z
        w, h = nd["w"] * z, nd["h"] * z
        x0, y0, x1, y1 = cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2
        click = bool(nd.get("click"))
        lw = max(1.0, (_LW_CLICK if click else _LW_STATIC) * z)
        kind = nd["kind"]

        if kind == "stadium":
            _round_rect(c, x0, y0, x1, y1, h / 2, fill=_NODE_FILL, outline=_NODE_LINE, width=lw)
        elif kind == "diamond":
            c.create_polygon(cx, y0, x1, cy, cx, y1, x0, cy,
                             fill=_NODE_FILL, outline=_NODE_LINE, width=lw)
        elif kind == "io":
            s = 0.3 * h                                # 平行四边形的斜边水平偏移
            c.create_polygon(x0 + s, y0, x1, y0, x1 - s, y1, x0, y1,
                             fill=_NODE_FILL, outline=_NODE_LINE, width=lw)
        else:
            # 矩形（草图里是直角矩形，不带圆角）
            c.create_rectangle(x0, y0, x1, y1, fill=_NODE_FILL, outline=_NODE_LINE, width=lw)

        fonts = self._fonts()
        c.create_text(cx, cy, text=nd["text"], justify="center",
                      font=fonts["node_u"] if click else fonts["node"],
                      fill=C_LINK if click else "#000000")
        if click:
            self._hits.append((x0, y0, x1, y1, nd["id"]))

    # ------------------------------------------------------------ 交互

    def _hit(self, cx: float, cy: float) -> str | None:
        for x0, y0, x1, y1, nid in self._hits:
            if x0 <= cx <= x1 and y0 <= cy <= y1:
                return nid
        return None

    def _on_motion(self, e) -> None:
        nid = self._hit(self.canvas.canvasx(e.x), self.canvas.canvasy(e.y))
        self.canvas.configure(cursor="hand2" if nid else "")

    def _begin_press(self, e) -> None:
        self._press = (e.x, e.y)
        self._press_hit = self._hit(self.canvas.canvasx(e.x), self.canvas.canvasy(e.y))
        self._dragged = False
        self.canvas.scan_mark(e.x, e.y)

    def _on_press(self, e) -> None:
        self._begin_press(e)

    def _on_drag(self, e) -> None:
        """左键拖动平移画布（放大之后看不全时，抓着图拖）。移动不到几个像素当作
        手抖，仍算点击。"""
        if self._press is None:
            return
        if not self._dragged:
            if abs(e.x - self._press[0]) + abs(e.y - self._press[1]) < _DRAG_SLOP:
                return
            self._dragged = True
            self.canvas.configure(cursor="fleur")
        self.canvas.scan_dragto(e.x, e.y, gain=1)

    def _on_release(self, e) -> None:
        """没拖动过＝点击：压在可点的框上就交给宿主去跳转——这里不改任何绘制状态。"""
        nid, dragged = self._press_hit, self._dragged
        self._press, self._press_hit, self._dragged = None, None, False
        self._on_motion(e)                            # 光标还原（手形/箭头）
        if dragged or nid is None or not self.on_activate:
            return
        # 松开时还得在同一个框上（按下后又滑出去再松开，不算点它）
        if self._hit(self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)) == nid:
            self.on_activate(nid, NODE_LABEL.get(nid, nid))


# ================================================================ 演示入口


def main() -> None:
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:                                  # noqa: BLE001
        pass

    root = tk.Tk()
    root.title("流程图导航（独立演示）")
    root.configure(bg="#f4f4f2")
    root.geometry("700x1000")
    root.minsize(520, 560)

    bottom = tk.Frame(root, bg="#e9e9e6")
    bottom.pack(side="bottom", fill="x")
    lbl = tk.Label(bottom, text="点蓝色带下划线的节点试试；Ctrl＋滚轮缩放，空白处双击复位",
                   font=(FAMILY, 9), bg="#e9e9e6", fg="#333", anchor="w")
    lbl.pack(side="left", padx=10, pady=4)
    zoom_lbl = tk.Label(bottom, text="", font=(FAMILY, 9), bg="#e9e9e6", fg="#333")
    zoom_lbl.pack(side="right", padx=10)

    def on_activate(nid: str, label: str) -> None:
        lbl.configure(text=f"（演示）跳转到：{NODE_ACTION.get(nid, label)}    [节点 id = {nid}]",
                      fg="#111111")

    panel = FlowPanel(root, on_activate=on_activate,
                      on_zoom=lambda p: zoom_lbl.configure(text=f"缩放 {p}%"))
    panel.pack(side="top", fill="both", expand=True)
    root.mainloop()


if __name__ == "__main__":
    main()
