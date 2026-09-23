"""格式页那张大表格（样式 × 11 项设置）——**画布自绘版**。

为什么重写（用户 2026-09-20：展开折叠项要刷新好一会儿，拖分隔条特别卡）：
旧版是"一格一个真控件"（32 行 × 每行约 20 个下拉框/输入框 ≈ 700 个）。Windows 下
每个 Tk 控件都是一个独立的系统窗口，几百个一起改尺寸、重画就慢：展开一个 16 行
的分组 0.67 秒，拖分隔条每帧 0.75～0.93 秒（其他页签 0.1～0.15 秒）。实测把 Python
这边的重排代码全去掉也只降到一半——瓶颈是控件个数本身。

现在的做法：
* 格子只在画布上**画**（矩形 + 文字），几百个格子只是几百个画布图元；
* 你点哪一格，才在那一格上临时放**一个**真正的输入框/下拉框（整张表同一时刻最多
  一个）；它绑着同一个 tk 变量，改的就是真值，收起来后画布上的文字自动跟上；
* 每格背后的 `tk.StringVar` / `BooleanVar` **原样保留**（键名 id/name/cn_font/
  en_font/size/bold/align/fl_val/fl_unit/li_val/li_unit/sb/sa/lr/lv/pb 不变），
  导出、保存预设读取表格的方式一点没动（`MainWindow._collect_style_spec`）。

键盘：点格子进入编辑；Tab / Shift+Tab 在格子之间前后移动（会自动滚到可见处）；
Enter 或 Esc 收起编辑框；复选框格用空格或回车切换。
"""

from __future__ import annotations

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

from essay_tool.config.units import CHINESE_FONT_SIZES
from essay_tool.gui.widgets import BORDER, BOX, CHECK, wheel_notch

FAMILY = "Microsoft YaHei UI"

# ---------------------------------------------------------------- 词汇表与常量
CN_FONTS = ["宋体", "黑体", "楷体", "仿宋", "微软雅黑", "华文中宋",
            "方正小标宋简体", "隶书", "幼圆"]
EN_FONTS = ["Times New Roman", "Arial", "Calibri", "Cambria", "Georgia",
            "Verdana", "Courier New"]
PT_TO_CN = {v: k for k, v in CHINESE_FONT_SIZES.items()}
ALIGN_LABEL = {
    WD_ALIGN_PARAGRAPH.JUSTIFY: "两端对齐",
    WD_ALIGN_PARAGRAPH.CENTER: "居中",
    WD_ALIGN_PARAGRAPH.LEFT: "左对齐",
    WD_ALIGN_PARAGRAPH.RIGHT: "右对齐",
}
ALIGN_OPTIONS = ["两端对齐", "居中", "左对齐", "右对齐"]
LINE_RULE_LABEL = {
    WD_LINE_SPACING.MULTIPLE: "多倍行距（倍）",
    WD_LINE_SPACING.EXACTLY: "固定值（磅）",
    WD_LINE_SPACING.AT_LEAST: "最小值（磅）",
    WD_LINE_SPACING.SINGLE: "单倍",
    WD_LINE_SPACING.DOUBLE: "双倍",
    WD_LINE_SPACING.ONE_POINT_FIVE: "1.5倍",
}
LINE_RULE_OPTIONS = ["多倍行距（倍）", "固定值（磅）", "最小值（磅）",
                     "单倍", "双倍", "1.5倍"]
HDR_BG = "#f2f2ef"
CELL_LINE = "#d9d9d4"
GRP_BG = "#e9e9e3"          # 分组头行的底色
FMT_W0 = 132                # 首列（样式名）最小宽（像素）
FMT_ROW_H = 28              # 行高下限（像素；实际按字体与缩放算）

HEADERS = ("中文字体", "西文字体", "字号", "加粗", "对齐", "首行缩进",
           "左缩进", "段前", "段后", "行距", "段前分页")
_FIELD_BORDER = "#8a8a8a"   # 可编辑格的边线
_WHITE = "#ffffff"


# ---------------------------------------------------------------- 取值格式化
def fmt_size(size) -> str:
    try:
        pt = size.pt
    except Exception:
        return str(size)
    return PT_TO_CN.get(pt, f"{pt:g}")


def fmt_indent(ind) -> tuple[str, str]:
    kind = getattr(getattr(ind, "kind", None), "name", "")
    if kind == "CHARS":
        return f"{ind.chars:g}", "字符"
    if kind == "LENGTH" and ind.length is not None:
        return f"{ind.length.cm:g}", "厘米"
    return "", "字符"


def fmt_len_pt(v) -> str:
    if v is None:
        return ""
    try:
        return f"{v.pt:g}"
    except Exception:
        return f"{float(v):g}"


def fmt_line_value(spec) -> str:
    v = spec.line_value
    if v is None:
        return ""
    if spec.line_rule == WD_LINE_SPACING.MULTIPLE:
        return f"{float(v):g}"
    try:
        return f"{v.pt:g}"
    except Exception:
        return f"{float(v):g}"


# ---------------------------------------------------------------- 数据模型
class _Part:
    """一格里的一个"小块"：一个下拉框 / 输入框 / 复选框 / 单位文字。"""

    __slots__ = ("row", "col", "kind", "var", "values", "min_chars", "cap_chars",
                 "small", "text", "rect", "tid")

    def __init__(self, row, col, kind, var=None, values=(), min_chars=0, text="",
                 cap_chars=0, small=False):
        self.row = row
        self.col = col
        self.kind = kind          # combo（可打字）/ combo_ro（只能选）/ entry / check / label
        self.var = var
        self.values = list(values)
        self.min_chars = min_chars
        self.cap_chars = cap_chars  # 可打字下拉框的宽度上限（字符数，0＝不设）；超出的文字用…截断
        self.small = small        # 用小一号字（单位/行距下拉框，同旧版的 8 号）
        self.text = text          # label 的文字
        self.rect = (0, 0, 0, 0)  # 画布坐标（每次重画时更新）
        self.tid = None           # 文字（或复选框里的蓝方块）图元 id


class _Row:
    def __init__(self, rv: dict, name: str, group: "_Group") -> None:
        self.rv = rv
        self.name = name
        self.group = group
        self.cols: dict[int, list[_Part]] = {}
        self.y = 0


class _Group:
    def __init__(self, title: str, open_: bool) -> None:
        self.title = title
        self.open = open_
        self.rows: list[_Row] = []
        self.y = 0


# ---------------------------------------------------------------- 控件本体
class FmtGrid(tk.Frame):
    """四象限（角落 + 表头 + 首列 + 正文）的自绘表格。

    对外只有：`set_groups(groups)` 灌数据、`rows`（每行的变量字典）、
    `groups`（分组状态）、`toggle_group(i)`。
    """

    def __init__(self, master, *, theme: dict, direct_ids=frozenset(),
                 open_state: dict | None = None) -> None:
        super().__init__(master, bg=theme["panel"], highlightbackground=theme["line"],
                         highlightthickness=1)
        self.theme = theme
        self.direct_ids = direct_ids
        self.open_state: dict[str, bool] = open_state if open_state is not None else {}
        self.groups: list[_Group] = []
        self.rows: list[dict] = []
        self._rows: list[_Row] = []

        s = max(1.0, float(self.winfo_fpixels("1i")) / 96.0)
        self.scale = s
        self.f_cell = tkfont.Font(family=FAMILY, size=9)
        self.f_small = tkfont.Font(family=FAMILY, size=8)
        self.f_name = tkfont.Font(family=FAMILY, size=10)
        self.f_hdr = tkfont.Font(family=FAMILY, size=9, weight="bold")
        self.f_grp = tkfont.Font(family=FAMILY, size=10, weight="bold")
        self.cw = max(1, self.f_cell.measure("0"))
        ls = self.f_cell.metrics("linespace")
        self.row_h = max(FMT_ROW_H, ls + round(8 * s))
        self.pad = max(3, round(3 * s))          # 格子里文字离边线
        self.gap = max(2, round(3 * s))          # 同一格里两个小块的间距
        self.mv = max(1, round(2 * s))           # 小块离格边线
        self.arrow_w = max(12, round(13 * s))    # 下拉箭头区宽
        self.box = BOX                           # 复选框方块（与 widgets.py 同一尺寸）

        self._w0 = FMT_W0
        self._col_w: dict[int, int] = {c: 60 for c in range(1, 12)}
        self._part_w: dict[int, list[int]] = {}
        self._total_w = 0
        self._total_h = 0
        self._parts_flat: list[_Part] = []
        self._editor = None                      # 当前编辑框控件
        self._editing: _Part | None = None
        self._cursor: _Part | None = None        # 复选框格的键盘光标
        self._ed_item = None
        self._eds: dict[str, tk.Widget] = {}
        self._refit_job = None
        self._last_avail = -1
        self._natural_total = 0                  # 不含"均摊余量"的列宽合计
        self._extra = 0                          # 当前均摊的余量（像素）
        self._mcache: dict = {}
        self._empty_text = "（还没读到样式——先应用一个配置）"

        # ---- 四象限骨架
        self.corner = tk.Frame(self, bg=CELL_LINE)
        self.corner.grid(row=0, column=0, sticky="nsew")
        tk.Label(self.corner, text="样式名", font=self.f_hdr, bg=HDR_BG, fg=theme["text"],
                 padx=6, pady=5).pack(fill="both", expand=True, padx=(0, 1), pady=(0, 1))
        kw = dict(highlightthickness=0, bd=0, bg=CELL_LINE)
        self.canvas_hdr = tk.Canvas(self, height=self.row_h, width=FMT_W0, **kw)
        self.canvas_hdr.grid(row=0, column=1, sticky="ew")
        self.canvas_left = tk.Canvas(self, width=FMT_W0, **kw)
        self.canvas_left.grid(row=1, column=0, sticky="ns")
        self.canvas_body = tk.Canvas(self, takefocus=1, **kw)
        self.canvas_body.grid(row=1, column=1, sticky="nsew")
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas_body.yview)
        self.hsb = ttk.Scrollbar(self, orient="horizontal", command=self.canvas_body.xview)
        self.vsb.grid(row=1, column=2, sticky="ns")
        self.hsb.grid(row=2, column=0, columnspan=2, sticky="ew")
        self.grid_rowconfigure(0, minsize=self.row_h)
        self.grid_columnconfigure(0, minsize=FMT_W0)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.canvas_body.configure(xscrollcommand=self._xcmd, yscrollcommand=self._ycmd)

        for cv in (self.canvas_hdr, self.canvas_left, self.canvas_body):
            cv.bind("<MouseWheel>", self._wheel)
            cv.bind("<Shift-MouseWheel>", self._wheel_h)
        self.canvas_body.bind("<Button-1>", self._on_body_click)
        self.canvas_left.bind("<Button-1>", self._on_left_click)
        self.canvas_body.bind("<Configure>", self._on_configure)
        self.canvas_body.bind("<space>", self._on_key_toggle)
        self.canvas_body.bind("<Return>", self._on_key_toggle)
        for seq, step in (("<Tab>", 1), ("<Shift-Tab>", -1), ("<ISO_Left_Tab>", -1)):
            self.canvas_body.bind(seq, lambda e, st=step: self._tab(st))
        self.canvas_body.bind("<FocusOut>", lambda e: self._clear_cursor())

    # ------------------------------------------------------------ 灌数据
    def set_groups(self, groups: list) -> None:
        """`groups`：`[(组名, [ParagraphStyleSpec…]), …]`；整张表重建（换配置时）。"""
        self._close_editor()
        self._cursor = None
        self.groups, self._rows, self.rows = [], [], []
        for title, specs in groups:
            if not specs:
                continue
            g = _Group(title, self.open_state.get(title, title == "正文部分"))
            for spec in specs:
                rv = self._make_vars(spec)
                name = spec.name + ("＊" if spec.id in self.direct_ids else "")
                row = _Row(rv, name, g)
                row.cols = self._make_parts(row, rv)
                g.rows.append(row)
                self._rows.append(row)
                self.rows.append(rv)
            self.groups.append(g)
        self._redraw()

    def _make_vars(self, spec) -> dict:
        rv: dict = {"id": spec.id, "name": spec.name}
        rv["cn_font"] = tk.StringVar(value=spec.font_east_asian)
        rv["en_font"] = tk.StringVar(value=spec.font_ascii)
        rv["size"] = tk.StringVar(value=fmt_size(spec.size))
        rv["bold"] = tk.BooleanVar(value=bool(spec.bold))
        rv["align"] = tk.StringVar(
            value=ALIGN_LABEL.get(spec.alignment, str(spec.alignment)))
        for valkey, unitkey, ind in (("fl_val", "fl_unit", spec.first_line_indent),
                                     ("li_val", "li_unit", spec.left_indent)):
            val, unit = fmt_indent(ind)
            rv[valkey] = tk.StringVar(value=val)
            rv[unitkey] = tk.StringVar(value=unit)
        rv["sb"] = tk.StringVar(value=fmt_len_pt(spec.space_before))
        rv["sa"] = tk.StringVar(value=fmt_len_pt(spec.space_after))
        rv["lr"] = tk.StringVar(
            value=LINE_RULE_LABEL.get(spec.line_rule, str(spec.line_rule)))
        rv["lv"] = tk.StringVar(value=fmt_line_value(spec))
        rv["pb"] = tk.BooleanVar(value=bool(spec.page_break_before))
        return rv

    def _make_parts(self, row: _Row, rv: dict) -> dict[int, list[_Part]]:
        align = rv["align"].get()
        aopts = ALIGN_OPTIONS if align in ALIGN_OPTIONS else ALIGN_OPTIONS + [align]
        rule = rv["lr"].get()
        ropts = LINE_RULE_OPTIONS if rule in LINE_RULE_OPTIONS else LINE_RULE_OPTIONS + [rule]

        def P(col, kind, key=None, values=(), chars=0, text="", cap=0, small=False):
            p = _Part(row, col, kind, rv[key] if key else None, values, chars, text,
                      cap, small)
            if p.var is not None:
                p.var.trace_add("write", lambda *_a, part=p: self._on_var(part))
            return p

        units = ["字符", "厘米"]
        cols = {
            1: [P(1, "combo", "cn_font", CN_FONTS, 5, cap=5)],
            2: [P(2, "combo", "en_font", EN_FONTS, 14, cap=14)],
            3: [P(3, "combo", "size", list(CHINESE_FONT_SIZES), 3, cap=4)],
            4: [P(4, "check", "bold")],
            5: [P(5, "combo_ro", "align", aopts, 5)],
            6: [P(6, "entry", "fl_val", (), 3), P(6, "combo_ro", "fl_unit", units, 3, small=True)],
            7: [P(7, "entry", "li_val", (), 3), P(7, "combo_ro", "li_unit", units, 3, small=True)],
            8: [P(8, "entry", "sb", (), 3), P(8, "label", text="磅")],
            9: [P(9, "entry", "sa", (), 3), P(9, "label", text="磅")],
            10: [P(10, "combo_ro", "lr", ropts, 9, small=True), P(10, "entry", "lv", (), 3)],
            11: [P(11, "check", "pb")],
        }
        return cols

    # ------------------------------------------------------------ 布局
    def _font_of(self, p: _Part):
        return self.f_small if p.small else self.f_cell

    def _mw(self, font, text: str) -> int:
        """`Font.measure` 带缓存（每次都是一次 Tcl 调用，布局时要量几百次）。"""
        key = (font.name, text)
        w = self._mcache.get(key)
        if w is None:
            w = self._mcache[key] = font.measure(text)
        return w

    def _text_w(self, p: _Part) -> int:
        f = self._font_of(p)
        if p.kind == "check":
            return self.box + 2 * self.pad
        if p.kind == "label":
            return self._mw(f, p.text) + self.pad
        texts = list(p.values) if p.kind == "combo_ro" else []
        texts.append(p.var.get())
        w = max(p.min_chars * self.cw, max((self._mw(f, t) for t in texts if t), default=0))
        if p.cap_chars:
            w = min(w, max(p.min_chars, p.cap_chars) * self.cw)
        w += 2 * self.pad
        if p.kind in ("combo", "combo_ro"):
            w += self.arrow_w
        return w

    def _layout(self) -> None:
        """算列宽/行位置。视口有余量时把余量均摊给各列（表格铺满，同旧版）。"""
        rows = self._rows
        part_w: dict[int, list[int]] = {}
        col_w: dict[int, int] = {}
        for c in range(1, 12):
            n = len(rows[0].cols[c]) if rows else 1
            pw = [0] * n
            for r in rows:
                for i, p in enumerate(r.cols[c]):
                    pw[i] = max(pw[i], self._text_w(p))
            hw = self._mw(self.f_hdr, HEADERS[c - 1]) + 2 * self.pad
            base = sum(pw) + self.gap * (n - 1) + 2 * self.mv
            if base < hw:                       # 表头比内容宽：多出的给最后一块
                pw[-1] += hw - base
                base = hw
            part_w[c], col_w[c] = pw, base + 1  # +1：右侧 1px 分格线

        # 首列宽：按当前看得见的行收紧（折叠时不留空白）
        names = [self._mw(self.f_grp, f"▼ {g.title}（{len(g.rows)}）") for g in self.groups]
        for g in self.groups:
            if g.open:
                names += [self._mw(self.f_name, r.name) for r in g.rows]
        self._w0 = max([FMT_W0 // 2, *names], default=FMT_W0) + 2 * self.pad + 2

        avail = self.canvas_body.winfo_width()
        total = sum(col_w.values())
        self._natural_total = total
        self._extra = 0
        if avail > 50 and rows:
            extra = avail - total - 2
            if extra > 0:
                self._extra = extra
                add, rem = divmod(extra, 11)
                for i, c in enumerate(range(1, 12)):
                    e = add + (1 if i < rem else 0)
                    col_w[c] += e
                    # 两块的格子：余量给"下拉/单位"那块，输入框保持紧凑
                    n = len(part_w[c])
                    idx = 0 if n == 1 else next(
                        (k for k, p in enumerate(rows[0].cols[c]) if p.kind != "entry"), n - 1)
                    part_w[c][idx] += e
                total += extra
        self._part_w, self._col_w = part_w, col_w
        self._total_w = sum(col_w.values())
        self._last_avail = avail

        y = 0
        for g in self.groups:
            g.y = y
            y += self.row_h
            for r in g.rows:
                r.y = y
                if g.open:
                    y += self.row_h
        self._total_h = y

    # ------------------------------------------------------------ 绘制
    def _redraw(self) -> None:
        self._close_editor()
        keep = self._cursor
        self._cursor = None
        self._parts_flat = []
        self._layout()
        hdr, left, body = self.canvas_hdr, self.canvas_left, self.canvas_body
        for cv in (hdr, left, body):
            cv.delete("all")
        rh, s = self.row_h, self.scale
        th = self.theme

        # 表头
        x = 0
        for c in range(1, 12):
            w = self._col_w[c]
            hdr.create_rectangle(x, 0, x + w - 1, rh - 1, fill=HDR_BG, outline="")
            hdr.create_text(x + (w - 1) / 2, (rh - 1) / 2, text=HEADERS[c - 1],
                            font=self.f_hdr, fill=th["text"])
            x += w
        hdr.configure(height=rh, scrollregion=(0, 0, max(self._total_w, 1), rh))
        self.grid_rowconfigure(0, minsize=rh)
        self.left_w = self._w0
        left.configure(width=self._w0)
        self.grid_columnconfigure(0, minsize=self._w0)

        if not self._rows:
            body.create_text(10, rh / 2, anchor="w", text=self._empty_text,
                             font=self.f_cell, fill=th["hint"])
            body.configure(scrollregion=(0, 0, 1, rh))
            left.configure(scrollregion=(0, 0, 1, rh))
            self._sync_scroll()
            return

        W, H = self._total_w, self._total_h
        for g in self.groups:
            left.create_rectangle(0, g.y, self._w0 - 1, g.y + rh - 1, fill=GRP_BG,
                                  outline="", tags=("grp",))
            left.create_text(self.pad, g.y + (rh - 1) / 2, anchor="w",
                             text=("▼ " if g.open else "▶ ") + f"{g.title}（{len(g.rows)}）",
                             font=self.f_grp, fill=th["text"])
            body.create_rectangle(0, g.y, W - 1, g.y + rh - 1, fill=GRP_BG, outline="")
            if not g.open:
                continue
            for r in g.rows:
                y0 = r.y
                left.create_rectangle(0, y0, self._w0 - 1, y0 + rh - 1, fill=HDR_BG, outline="")
                left.create_text(self.pad, y0 + (rh - 1) / 2, anchor="w", text=r.name,
                                 font=self.f_name, fill=th["text"])
                x = 0
                for c in range(1, 12):
                    cw_ = self._col_w[c]
                    body.create_rectangle(x, y0, x + cw_ - 1, y0 + rh - 1, fill=_WHITE, outline="")
                    px = x + self.mv
                    for i, p in enumerate(r.cols[c]):
                        w = self._part_w[c][i]
                        p.rect = (px, y0 + self.mv, px + w, y0 + rh - 1 - self.mv)
                        self._draw_part(p)
                        px += w + self.gap
                        if p.kind != "label":
                            self._parts_flat.append(p)
                    x += cw_
        # 不可见（折叠）行的部件矩形置空，防止点击误命中
        for g in self.groups:
            if not g.open:
                for r in g.rows:
                    for ps in r.cols.values():
                        for p in ps:
                            p.rect = (0, 0, 0, 0)
                            p.tid = None
        body.create_rectangle(0, 0, W, H, outline="", fill="", tags=("bounds",))
        left.create_rectangle(0, 0, self._w0, H, outline="", fill="")
        body.configure(scrollregion=(0, 0, W, H))
        left.configure(scrollregion=(0, 0, self._w0, H))
        self._sync_scroll()
        if keep is not None and keep in self._parts_flat:
            self._set_cursor(keep, reveal=False)

    def _clip(self, text: str, avail: int, f=None) -> str:
        f = f or self.f_cell
        if avail <= 0 or f.measure(text) <= avail:
            return text
        ell = "…"
        while text and f.measure(text + ell) > avail:
            text = text[:-1]
        return text + ell

    def _draw_part(self, p: _Part) -> None:
        cv = self.canvas_body
        x0, y0, x1, y1 = p.rect
        cy = (y0 + y1) / 2
        th = self.theme
        if p.kind == "check":
            bx = (x0 + x1) // 2 - self.box // 2
            by = int(cy) - self.box // 2
            cv.create_rectangle(bx, by, bx + self.box - 1, by + self.box - 1,
                                outline=BORDER, fill=_WHITE)
            q = 3
            p.tid = cv.create_rectangle(bx + q, by + q, bx + self.box - 1 - q,
                                        by + self.box - 1 - q, outline="",
                                        fill=CHECK if p.var.get() else _WHITE)
            return
        if p.kind == "label":
            p.tid = cv.create_text(x0 + self.pad // 2, cy, anchor="w", text=p.text,
                                   font=self.f_cell, fill=th["hint"])
            return
        cv.create_rectangle(x0, y0, x1, y1, outline=_FIELD_BORDER, fill=_WHITE)
        arrow = self.arrow_w if p.kind in ("combo", "combo_ro") else 0
        avail = (x1 - x0) - 2 * self.pad - arrow
        font = self._font_of(p)
        text = self._clip(p.var.get(), avail, font)
        if p.kind == "entry":
            p.tid = cv.create_text((x0 + x1) / 2, cy, text=text, font=font,
                                   fill=th["text"])
        else:
            p.tid = cv.create_text(x0 + self.pad, cy, anchor="w", text=text,
                                   font=font, fill=th["text"])
            ax, u = x1 - arrow / 2, max(2, round(3 * self.scale))
            cv.create_line(ax - u, cy - u / 2, ax, cy + u / 2, ax + u, cy - u / 2,
                           fill="#666666", width=max(1, round(self.scale)))

    def _on_var(self, p: _Part) -> None:
        """变量变了（编辑框里打字/程序赋值）：把画布上对应的图元同步一下。"""
        cv = self.canvas_body
        if p.tid is None:
            return
        try:
            if p.kind == "check":
                cv.itemconfigure(p.tid, fill=CHECK if p.var.get() else _WHITE)
            else:
                x0, _y0, x1, _y1 = p.rect
                arrow = self.arrow_w if p.kind in ("combo", "combo_ro") else 0
                cv.itemconfigure(p.tid, text=self._clip(
                    p.var.get(), (x1 - x0) - 2 * self.pad - arrow, self._font_of(p)))
        except tk.TclError:
            pass

    # ------------------------------------------------------------ 滚动
    def _sync_scroll(self) -> None:
        self._ycmd(*self.canvas_body.yview())
        self._xcmd(*self.canvas_body.xview())

    @staticmethod
    def _sb_state(sb, first, last) -> None:
        fit = float(first) <= 0.0 and float(last) >= 1.0
        sb.state(["disabled"] if fit else ["!disabled"])

    def _ycmd(self, first, last) -> None:
        self.vsb.set(first, last)
        self.canvas_left.yview_moveto(first)
        self._sb_state(self.vsb, first, last)

    def _xcmd(self, first, last) -> None:
        self.hsb.set(first, last)
        self.canvas_hdr.xview_moveto(first)
        self._sb_state(self.hsb, first, last)

    def _scroll_by(self, units: int, vertical: bool) -> None:
        """按格滚动，夹在内容范围内（一格≈视口的 1/10，同旧版）。"""
        cv = self.canvas_body
        if not units or not self._rows:
            return
        if vertical:
            total, view, top = self._total_h, cv.winfo_height(), cv.canvasy(0)
        else:
            total, view, top = self._total_w, cv.winfo_width(), cv.canvasx(0)
        step = max(20, int(view * 0.1))
        new_top = max(0.0, min(top + units * step, max(0.0, total - view)))
        frac = new_top / total if total else 0.0
        (cv.yview_moveto if vertical else cv.xview_moveto)(frac)

    def _wheel(self, event):
        self._scroll_by(-2 * wheel_notch(event), True)
        return "break"

    def _wheel_h(self, event):
        self._scroll_by(-2 * wheel_notch(event), False)
        return "break"

    def _reveal(self, p: _Part) -> None:
        """把格子滚到可见范围内。"""
        cv = self.canvas_body
        x0, y0, x1, y1 = p.rect
        vw, vh = cv.winfo_width(), cv.winfo_height()
        vx, vy = cv.canvasx(0), cv.canvasy(0)
        if y0 < vy:
            cv.yview_moveto(max(0.0, (y0 - self.row_h) / self._total_h))
        elif y1 > vy + vh:
            cv.yview_moveto(min(1.0, (y1 - vh + self.row_h) / self._total_h))
        if x0 < vx:
            cv.xview_moveto(max(0.0, (x0 - self.pad) / self._total_w))
        elif x1 > vx + vw:
            cv.xview_moveto(min(1.0, (x1 - vw + self.pad) / self._total_w))
        self.update_idletasks()

    def _on_configure(self, _e=None) -> None:
        if self._refit_job is not None:
            self.after_cancel(self._refit_job)
        self._refit_job = self.after(120, self._refit)

    def _refit(self) -> None:
        """视口尺寸变了。只有"均摊给各列的余量"变了才需要整表重画：视口在
        变窄、列宽本来就没有余量可摊时（拖分隔条最常见的情形）什么都不用重画，
        只同步一下滚动条——这是拖动流畅的关键（实测重画每帧都要跟着重绘整块）。"""
        self._refit_job = None
        avail = self.canvas_body.winfo_width()
        extra = max(0, avail - self._natural_total - 2) if avail > 50 else 0
        self._last_avail = avail
        if self._rows and extra != self._extra:
            self._redraw()
        else:
            self._sync_scroll()

    # ------------------------------------------------------------ 折叠
    def toggle_group(self, gi: int) -> None:
        g = self.groups[gi]
        g.open = not g.open
        self.open_state[g.title] = g.open
        self._redraw()

    def _on_left_click(self, e) -> None:
        y = self.canvas_left.canvasy(e.y)
        for i, g in enumerate(self.groups):
            if g.y <= y < g.y + self.row_h:
                self.toggle_group(i)
                return

    def _group_at(self, y: float) -> int | None:
        for i, g in enumerate(self.groups):
            if g.y <= y < g.y + self.row_h:
                return i
        return None

    # ------------------------------------------------------------ 点击/键盘
    def _part_at(self, x: float, y: float) -> _Part | None:
        for g in self.groups:
            if not g.open:
                continue
            for r in g.rows:
                if r.y <= y < r.y + self.row_h:
                    for ps in r.cols.values():
                        for p in ps:
                            x0, y0, x1, y1 = p.rect
                            if p.kind != "label" and x0 <= x <= x1 and y0 <= y <= y1:
                                return p
                    return None
        return None

    def _on_body_click(self, e) -> str:
        cv = self.canvas_body
        if self._editor is not None:
            self._end_edit(focus_canvas=False)   # 先收起上一个编辑框（文字变长会重画）
        x, y = cv.canvasx(e.x), cv.canvasy(e.y)
        gi = self._group_at(y)
        if gi is not None:                      # 点分组色带 = 展开/折叠
            self.toggle_group(gi)
            return "break"
        p = self._part_at(x, y)
        if p is None:
            cv.focus_set()
            return "break"
        self._activate(p, click_x=x)
        return "break"

    def _activate(self, p: _Part, click_x: float | None = None) -> None:
        if p.kind == "check":
            self._set_cursor(p, reveal=True)
            p.var.set(not bool(p.var.get()))
            return
        self._clear_cursor()
        self._open_editor(p, click_x)

    def _on_key_toggle(self, _e=None) -> str:
        p = self._cursor
        if p is not None and p.kind == "check":
            p.var.set(not bool(p.var.get()))
        return "break"

    def _set_cursor(self, p: _Part, reveal: bool = True) -> None:
        """复选框格的键盘光标（虚线框）；画布拿键盘焦点，空格/回车切换。"""
        self._clear_cursor()
        if reveal:
            self._reveal(p)
        self._cursor = p
        x0, y0, x1, y1 = p.rect
        self.canvas_body.create_rectangle(x0, y0, x1, y1, outline=CHECK, dash=(3, 2),
                                          tags=("cursor",))
        self.canvas_body.focus_set()

    def _clear_cursor(self) -> None:
        self._cursor = None
        try:
            self.canvas_body.delete("cursor")
        except tk.TclError:
            pass

    def _tab(self, step: int) -> str:
        cur = self._editing or self._cursor
        if self._editor is not None:
            self._end_edit(focus_canvas=False)   # 文字变长会重画，之后再取 flat
        flat = self._parts_flat
        if not flat:
            return "break"
        i = flat.index(cur) if cur in flat else (-1 if step > 0 else len(flat))
        j = i + step
        if 0 <= j < len(flat):
            self._goto(flat[j])
        else:
            self._clear_cursor()
            self.canvas_body.focus_set()
        return "break"

    def _goto(self, p: _Part) -> None:
        self._reveal(p)
        if p.kind == "check":
            self._set_cursor(p, reveal=False)
        else:
            self._clear_cursor()
            self._open_editor(p)

    # ------------------------------------------------------------ 编辑框
    def _get_editor(self, kind: str) -> tk.Widget:
        w = self._eds.get(kind)
        if w is not None:
            return w
        cv = self.canvas_body
        if kind == "entry":
            w = tk.Entry(cv, font=self.f_cell, justify="center", relief="solid", bd=1,
                         highlightthickness=0)
        elif kind == "combo":
            w = ttk.Combobox(cv, font=self.f_cell)
        else:
            w = ttk.Combobox(cv, font=self.f_cell, state="readonly")
        w.bind("<MouseWheel>", lambda e: (self._wheel(e), "break")[1])
        w.bind("<Shift-MouseWheel>", lambda e: (self._wheel_h(e), "break")[1])
        for seq, step in (("<Tab>", 1), ("<Shift-Tab>", -1), ("<ISO_Left_Tab>", -1)):
            w.bind(seq, lambda e, st=step: self._tab(st))
        w.bind("<Return>", lambda e: self._end_edit())
        w.bind("<Escape>", lambda e: self._end_edit())
        w.bind("<FocusOut>", lambda e: self.after(80, self._maybe_close))
        if kind != "entry":
            w.bind("<<ComboboxSelected>>", lambda e: self.after_idle(self._end_edit))
        self._eds[kind] = w
        return w

    def _open_editor(self, p: _Part, click_x: float | None = None) -> None:
        self._close_editor()
        w = self._get_editor(p.kind)
        x0, y0, x1, y1 = p.rect
        if p.kind != "entry":
            w.configure(values=p.values)
        w.configure(textvariable=p.var, font=self._font_of(p))
        self._ed_item = self.canvas_body.create_window(
            x0, y0, anchor="nw", width=x1 - x0, height=y1 - y0, window=w, tags=("editor",))
        self._editor, self._editing = w, p
        self.update_idletasks()
        w.focus_set()
        if p.kind == "entry" or p.kind == "combo":
            try:
                w.select_range(0, "end")
                w.icursor("end")
            except tk.TclError:
                pass
        if p.kind == "combo_ro":
            self.after_idle(lambda ww=w: self._post(ww))
        elif p.kind == "combo" and click_x is not None and click_x >= x1 - self.arrow_w:
            self.after_idle(lambda ww=w: self._post(ww))

    def _post(self, w) -> None:
        if self._editor is w:
            try:
                self.tk.call("ttk::combobox::Post", str(w))
            except tk.TclError:
                pass

    def _maybe_close(self) -> None:
        """编辑框失去焦点：焦点还在它自己（含下拉列表）就保留，否则收起。"""
        if self._editor is None:
            return
        # 不能用 `focus_get()`：ttk 下拉列表（`<编辑框>.popdown…`）是 tkinter 不认识的
        # 内部窗口，会抛 KeyError——被当成"焦点走了"就会一点开下拉就把编辑框收掉。
        try:
            f = str(self.tk.call("focus"))
        except tk.TclError:
            f = ""
        if f and f.startswith(str(self._editor)):
            return
        self._end_edit(focus_canvas=False)

    def _end_edit(self, focus_canvas: bool = True) -> str:
        was_open = self._editor is not None
        self._close_editor()
        if focus_canvas:
            self.canvas_body.focus_set()
        if was_open:
            self._relayout_if_widths_changed()
        return "break"

    def _close_editor(self) -> None:
        if self._editor is None:
            return
        w, self._editor, self._editing = self._editor, None, None
        try:
            w.configure(textvariable="")
        except tk.TclError:
            pass
        if self._ed_item is not None:
            try:
                self.canvas_body.delete(self._ed_item)
            except tk.TclError:
                pass
            self._ed_item = None

    def _relayout_if_widths_changed(self) -> None:
        """编辑后文字可能变长：列宽需要跟着变时整表重画一次（很便宜）。"""
        old = (dict(self._col_w), {c: list(v) for c, v in self._part_w.items()}, self._w0)
        self._layout()
        if (self._col_w, self._part_w, self._w0) != old:
            self._redraw()
