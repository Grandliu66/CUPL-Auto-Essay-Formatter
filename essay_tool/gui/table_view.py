# -*- coding: utf-8 -*-
"""自绘表格控件（Canvas）：格子里的文字**支持字符级选中**，**每行按内容给行高**。

为什么不用 `ttk.Treeview`（2026-09-17 用户两条要求推着换掉）：

  * 它只能**整行**选中，做不到"我想复制一个字就复制一个字"（用户原话："我要的是
    选中文字，懂吗，选中文字复制"）；
  * 行高是**样式级、全表统一**的 ⇒ 只能按"全表最多行的那个格子"给所有行留一样的
    高度，短行下面空一大块（用户："行高还是太大了，空白太大。每一行的字符需要多少
    行高，就给多少行高，类似『最合适的行高』"）。

自绘之后这两条直接成立，另外还白捡三个好处：行高＝行距×该行行数（不四舍五入、
不留余量）、滚动按**像素**平滑（Treeview 滚不动半行，`moveto` 收小数会被吸到整行）、
表头固定只靠一个 `move`。

坐标口径（全部是**画布文档坐标**，滚动不影响）：

    W = 2*GRID + Σ列宽                H = 2*GRID + 表头高 + Σ行高
    第 c 列左边缘 = GRID + Σ(前面的列宽)
    第 r 行上边缘 = GRID + 表头高 + Σ(前面的行高)

⚠️ 表头是"钉"在视口顶部的：它画在文档坐标 y≈GRID 处，每次纵向滚动后整体
`canvas.move("head", 0, dy)` 跟回视口顶（`_sync_head`）。行本身坐标不动。
"""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Callable, Sequence

# 配色：表头白底，**不分色块了**（用户 2026-09-17：“浅蓝色部分的框线没渲染出来——
# 其实不用区分色块了”），行与行靠**实实在在的黑框线**分开；
# 那些色块当年是给 Treeview 用的（它没有真框线），自绘以后反而会把框线盖掉。
C_GRID = "#000000"          # 单元格框线
C_BORDER = "#000000"        # 外框
C_SEL = "#cfe1f4"           # 选中底色（浅蓝，接近系统选中色）
C_HEAD_BG = "#ffffff"
C_TEXT = "#1a1a1a"
C_HEAD_TEXT = "#000000"

GRID = 1                    # 外框线宽
PAD_X = 8                   # 单元格左右内边距（列宽里已经含 2×PAD_X）
HEAD_PAD_Y = 4              # 表头那一行上下各留的空白

# 位置＝（第几行, 第几列, 格内第几行, 行内第几个字符）
Pos = tuple[int, int, int, int]


class CanvasTable(tk.Frame):
    """一张自绘表格：逐行给高、字符级选中、像素级滚动。

    外部要用的东西：`row_count` / `cell_text(r, c)` / `head_text(c)` /
    `hit(x, y)` / `sel_text()` / `copy_*()` / `can_scroll_y/scroll_y` /
    `can_scroll_x/scroll_x`。
    """

    def __init__(self, master: tk.Misc, *, head: Sequence[str],
                 aligns: Sequence[str], body: Sequence[Sequence[str]],
                 avail: float, wrap: Callable[[str, float], list[str]],
                 font, head_font, line_h: int, max_h: int,
                 cell_min_w: int, cell_max_w: int, vscroll_w: int,
                 on_wheel: Callable[[object], str] | None = None,
                 on_menu: Callable[[object], str] | None = None,
                 bg: str = "#ffffff") -> None:
        super().__init__(master, bg=bg, highlightthickness=GRID,
                         highlightbackground=C_BORDER, highlightcolor=C_BORDER)
        self.font = font
        self.head_font = head_font
        self.line_h = int(line_h)
        self.vscroll_w = int(vscroll_w)
        self.ncols = len(head)
        self.rows_n = len(body)
        self.aligns = list(aligns)
        self.head_raw = [str(h) for h in head]
        # 数据行列数可能比表头少（作者少写一个竖线）：补空格子，后面按下标取就不会越界
        self.body_raw = [[str(cell) for cell in row]
                         + [""] * max(0, len(head) - len(row)) for row in body]
        self._cw: dict[str, float] = {}          # 单字宽度缓存
        self._line_x: dict[tuple[int, int, int], float] = {}   # 每行文字的左起点
        self._line_y: dict[tuple[int, int, int], float] = {}
        self._head_x: dict[int, float] = {}
        self.row_y: list[float] = []
        self.row_h: list[int] = []
        self._sel: tuple[Pos, Pos] = ((0, 0, 0, 0), (0, 0, 0, 0))
        self._anchor: Pos | None = None
        self._moved = False
        self._just_dbl = False

        # ---- ① 列宽：自然宽（按最长的一行量，上限 `cell_max_w`）----
        cap = cell_max_w
        if self.ncols:
            room = avail - 2 * GRID - 16.0 * self.ncols
            cap = int(max(cell_min_w, min(cap, room / self.ncols)))
        self.colw = []
        for c in range(self.ncols):
            want = self._widest(self.head_raw[c], self.head_font)
            for r in range(self.rows_n):
                want = max(want, self._widest(self.body_raw[r][c], self.font))
            self.colw.append(int(max(min(want, cap) + 2 * PAD_X, 56)))

        # ---- ② 折行：每格切"段"（作者自己写的换行）→ 段内再按列宽折行 ----
        self.wrap = wrap
        self.head_cells = [self._cell_lines(self.head_raw[c], self.head_font, c)
                           for c in range(self.ncols)]
        self.cells: list[list[list[tuple[str, str]]]] = []
        for r in range(self.rows_n):
            self.cells.append([self._cell_lines(self.body_raw[r][c], self.font, c)
                               for c in range(self.ncols)])

        # ---- ③ 剩余宽度全给最右列（用户 2026-09-17；先按"要竖滚动条"预留）----
        self._relayout(avail - 2 * GRID - self.vscroll_w)
        self.max_h = int(max_h)
        need_v = self._content_h() > self.max_h
        if self.ncols and not need_v:
            self.colw[-1] += self.vscroll_w          # 不竖滚就把那 17px 也让给最右列
            self._rewrap_last()

        # ---- ④ 画布 + 滚动条 ----
        # 滚动条**对象总是建**（`_on_yscroll` 要回调它），只是不需要时不 pack。
        inner_h = self._content_h() + 2 * GRID
        self.need_v = inner_h > self.max_h
        self.view_h = min(inner_h, self.max_h)
        self.need_h = self.content_w > avail - 2 * GRID - \
            (self.vscroll_w if self.need_v else 0)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0,
                                yscrollincrement=1, xscrollincrement=1)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        if self.need_v:
            self.vsb.pack(side="right", fill="y")
        self.hsb = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        if self.need_h:
            self.hsb.pack(side="bottom", fill="x")
        self.canvas.pack(side="top", fill="both", expand=True)
        self.canvas.configure(yscrollcommand=self._on_yscroll,
                              xscrollcommand=self.hsb.set)
        self.canvas.configure(width=int(avail) - 2 * GRID -
                              (self.vscroll_w if self.need_v else 0),
                              height=max(1, self.view_h))
        self.canvas.configure(scrollregion=(0, 0, self.content_w, inner_h))
        self._head_top = float(GRID)
        self._draw()
        self.configure(width=int(avail),
                       height=self.view_h + (self.hsb.winfo_reqheight()
                                             if self.need_h else 0))
        self.pack_propagate(False)
        self.grid_propagate(False)

        # ---- ⑤ 事件：拖选、双击选格、Ctrl+C/A、右键菜单、滚轮 ----
        self.canvas.bind("<Button-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.canvas.bind("<Double-Button-1>", self._on_double)
        self.canvas.bind("<Control-c>", self.copy_selection)
        self.canvas.bind("<Control-C>", self.copy_selection)
        self.canvas.bind("<Control-a>", self.select_all)
        self.canvas.bind("<Control-A>", self.select_all)
        if on_menu is not None:
            self.canvas.bind("<Button-3>", on_menu)
            self.bind("<Button-3>", on_menu)
        if on_wheel is not None:
            self.canvas.bind("<MouseWheel>", on_wheel)
            self.bind("<MouseWheel>", on_wheel)

    # ------------------------------------------------------------- 折行/度量
    def _cwidth(self, ch: str) -> float:
        w = self._cw.get(ch)
        if w is None:
            w = self._cw[ch] = self.font.measure(ch)
        return w

    def _widest(self, text: str, font) -> float:
        return max([font.measure(part) for part in text.split("\n")] or [0.0])

    def _cell_lines(self, text: str, font, col: int) -> list[tuple[str, str]]:
        """一格 → [(显示行, 这行后面的分隔符)]；分隔符＝段内折行给空格、作者换行给 \\n。

        这样"复制选中文字"跨行时能还原作者本意：折行是我们造出来的（粘出来给空格，
        免得粘进 Excel 一行裂成好几行），作者自己敲的回车给真换行。
        """
        inner = max(40.0, self.colw[col] - 2 * PAD_X)
        segs = text.split("\n")
        out: list[tuple[str, str]] = []
        for si, seg in enumerate(segs):
            lines = [seg] if font.measure(seg) <= inner else self.wrap(seg, inner)
            for li, line in enumerate(lines):
                last = li == len(lines) - 1
                out.append((line, ("\n" if si < len(segs) - 1 else "") if last else " "))
        return out or [("", "")]

    def _relayout(self, room: float) -> None:
        """把列宽铺满 `room`：差额全给最右列，再按最终列宽重排各行高。"""
        extra = int(room) - (2 * GRID + sum(self.colw))
        if extra > 0 and self.ncols:
            self.colw[-1] += extra
            self._rewrap_last()
        self._rows()

    def _rewrap_last(self) -> None:
        if not self.ncols:
            return
        c = self.ncols - 1
        self.head_cells[c] = self._cell_lines(self.head_raw[c], self.head_font, c)
        for r in range(self.rows_n):
            self.cells[r][c] = self._cell_lines(self.body_raw[r][c], self.font, c)

    def _rows(self) -> None:
        """**每一行按自己内容给高**：行距 × 该行最多的那个格子的行数（不留余量）。"""
        self.row_h = []
        for r in range(self.rows_n):
            lines = max([len(cell) for cell in self.cells[r]] or [1])
            self.row_h.append(self.line_h * max(1, lines))
        self.head_h = self.line_h * max([len(c) for c in self.head_cells] or [1]) \
            + 2 * HEAD_PAD_Y

    def _content_h(self) -> int:
        return self.head_h + sum(self.row_h)

    # ---------------------------------------------------------------- 命中测试
    def _col_at(self, x: float) -> int:
        left = GRID
        for c, w in enumerate(self.colw):
            if x < left + w:
                return c
            left += w
        return max(0, self.ncols - 1)

    def _char_at(self, text: str, local_x: float) -> int:
        """行内哪个字符位置离 `local_x` 最近（返回插在它前面的下标）。"""
        x = 0.0
        for i, ch in enumerate(text):
            w = self._cwidth(ch)
            if local_x < x + w / 2:
                return i
            x += w
        return len(text)

    def hit(self, x: float, y: float) -> tuple[int, int] | None:
        """画布坐标 → (行, 列)；落在表头上返回 (-1, 列)，落在表外返回 None。"""
        col = self._col_at(x)
        if y < GRID + self.head_h:
            return (-1, col) if y >= GRID else None
        top = GRID + self.head_h
        for r, h in enumerate(self.row_h):
            if y < top + h:
                return (r, col)
            top += h
        return (self.rows_n - 1, col) if self.rows_n else None

    def pos_at(self, x: float, y: float) -> Pos | None:
        """画布坐标 → 精确到"第几个字符"的位置（拖选用）。"""
        got = self.hit(x, y)
        if got is None or got[0] < 0:
            return None
        r, c = got
        cell = self.cells[r][c]
        top = self.row_y[r]
        i = int(max(0, min(len(cell) - 1, (y - top) // self.line_h)))
        text = cell[i][0]
        local = x - self._line_x.get((r, c, i), self._colx(c) + PAD_X)
        return (r, c, i, self._char_at(text, local))

    def _colx(self, c: int) -> float:
        return GRID + sum(self.colw[:c])

    # ------------------------------------------------------------------ 绘制
    def _draw(self) -> None:
        cv = self.canvas
        cv.delete("all")
        self._line_x.clear()
        self._line_y.clear()
        self._head_x.clear()
        total_w = sum(self.colw)
        bottom = GRID + self._content_h()
        # ① 白底 + 行线 + 列线 + 外框
        # ⚠️ 框线单独挂 `grid` 标签，最后再 `tag_lower("grid", "celltext")` 插到
        # “选中底色之上、文字之下”：不这么干的话，**选中区域里那几条框线会被选中
        # 底色盖掉**（用户 2026-09-17 实拍：浅色区里看不到框线）。行与行不再用色块
        # 区分，所以也不能靠斑马纹去“补”出分界。
        cv.create_rectangle(GRID, GRID, GRID + total_w, bottom,
                            fill="#ffffff", outline="", tags="bg")
        left = GRID
        for c, w in enumerate(self.colw):
            if c:
                cv.create_line(left, GRID, left, bottom, fill=C_GRID, tags="grid")
            left += w
        top = GRID + self.head_h
        self.row_y = []
        for r, h in enumerate(self.row_h):
            self.row_y.append(top)
            cv.create_line(GRID, top, GRID + total_w, top, fill=C_GRID, tags="grid")
            top += h
        cv.create_line(GRID, bottom, GRID + total_w, bottom, fill=C_GRID, tags="grid")
        cv.create_rectangle(GRID, GRID, GRID + total_w, bottom,
                            outline=C_BORDER, tags="grid")
        # ② 正文格子
        for r in range(self.rows_n):
            for c in range(self.ncols):
                self._draw_cell(cv, r, c)
        # ③ 表头（tag="head"，纵向滚动时被 `_sync_head` 拖回视口顶）
        cv.create_rectangle(GRID, GRID, GRID + total_w, GRID + self.head_h,
                            fill=C_HEAD_BG, outline=C_BORDER, tags="head")
        # 表头这块白底会把下面那些列线盖掉 ⇒ 在表头区里**重画一次**（同样挂 "head"，
        # 纵向滚动时跟表头一起走）。下边那条线是表头与第一行的分界，也要重画，
        # 否则被白底压掉半个像素、看着断断续续。
        for c in range(1, self.ncols):
            x = self._colx(c)
            cv.create_line(x, GRID, x, GRID + self.head_h, fill=C_GRID, tags="head")
        cv.create_line(GRID, GRID + self.head_h, GRID + total_w, GRID + self.head_h,
                       fill=C_GRID, tags="head")
        for c in range(self.ncols):
            for i, (line, _sep) in enumerate(self.head_cells[c]):
                x = self._text_x(c, line, self.head_font)
                y = GRID + HEAD_PAD_Y + i * self.line_h
                self._head_x[c] = self._colx(c) + PAD_X
                cv.create_text(x, y, text=line, font=self.head_font,
                               fill=C_HEAD_TEXT, anchor="nw", tags=("celltext", "head"))
        cv.tag_raise("head")
        self._draw_sel()
        cv.tag_lower("grid", "celltext")      # 框线压在选中底色之上、文字之下

    def _text_x(self, c: int, line: str, font) -> float:
        """一格文字的左起点：按列对齐方式（ttk 口径：'w' 左、'center' 中、'e' 右）。"""
        inner = self.colw[c] - 2 * PAD_X
        align = self.aligns[c] if c < len(self.aligns) else "w"
        used = font.measure(line)
        if align == "e":
            off = inner - used
        elif align == "center":
            off = (inner - used) / 2
        else:
            off = 0.0
        return self._colx(c) + PAD_X + off

    def _draw_cell(self, cv: tk.Canvas, r: int, c: int) -> None:
        for i, (line, _sep) in enumerate(self.cells[r][c]):
            y = self.row_y[r] + i * self.line_h
            x = self._text_x(c, line, self.font)
            key = (r, c, i)
            self._line_x[key] = x
            self._line_y[key] = y
            cv.create_text(x, y, text=line, font=self.font, fill=C_TEXT,
                           anchor="nw", tags="celltext")

    # ------------------------------------------------------------------ 选中
    def _sorted_sel(self) -> tuple[Pos, Pos]:
        a, b = self._sel
        return (a, b) if a <= b else (b, a)

    def sel_text(self) -> str:
        """选中的文字。段内折行 → 空格；作者换行 → \\n；跨格 → \\t；跨行 → \\n。"""
        if not self.rows_n or self._sel[0] == self._sel[1]:
            return ""
        (r0, c0, l0, k0), (r1, c1, l1, k1) = self._sorted_sel()
        out: list[str] = []
        for r in range(r0, r1 + 1):
            cs = c0 if r == r0 else 0
            ce = c1 if r == r1 else self.ncols - 1
            for c in range(cs, ce + 1):
                cell = self.cells[r][c]
                last = len(cell) - 1
                # ⚠️ 起止点落在格子中间某一折行时（`l0`/`l1` 不是这一格的第一/
                # 最后一行），选区在这一格里也该从那一行开始/到那一行为止——
                # 不能不管三七二十一把这一格所有折行都算进去，不然拖选跨行
                # 复制出来的文字会多出没选中的部分（2026-09-22 审计揪出的真
                # bug）。
                lo = l0 if (r == r0 and c == cs) else 0
                hi = l1 if (r == r1 and c == ce) else last
                for i in range(lo, hi + 1):
                    text, sep = cell[i]
                    a = k0 if i == lo and (r == r0 and c == cs) else 0
                    b = k1 if i == hi and (r == r1 and c == ce) else len(text)
                    out.append(text[a:b])
                    if i < hi:
                        out.append(sep or " ")
                if c < ce:
                    out.append("\t")
            if r < r1:
                out.append("\n")
        return "".join(out)

    def _draw_sel(self) -> None:
        cv = self.canvas
        cv.delete("selbg")
        if not self.rows_n or self._sel[0] == self._sel[1]:
            return
        (r0, c0, l0, k0), (r1, c1, l1, k1) = self._sorted_sel()
        for r in range(r0, r1 + 1):
            cs = c0 if r == r0 else 0
            ce = c1 if r == r1 else self.ncols - 1
            for c in range(cs, ce + 1):
                cell = self.cells[r][c]
                last = len(cell) - 1
                lo = l0 if (r == r0 and c == cs) else 0
                hi = l1 if (r == r1 and c == ce) else last
                for i in range(lo, hi + 1):
                    text = cell[i][0]
                    x0 = self._line_x.get((r, c, i))
                    y0 = self._line_y.get((r, c, i))
                    if x0 is None:
                        continue
                    a = k0 if i == lo and (r == r0 and c == cs) else 0
                    b = k1 if i == hi and (r == r1 and c == ce) else len(text)
                    if b < a:
                        a, b = b, a
                    xs = x0 + sum(self._cwidth(ch) for ch in text[:a])
                    xe = x0 + sum(self._cwidth(ch) for ch in text[:b])
                    if xe <= xs:                       # 空行/空选区：给个细标记
                        xe = xs + 2
                    cv.create_rectangle(xs, y0, xe, y0 + self.line_h,
                                        fill=C_SEL, outline="", tags="selbg")
        cv.tag_lower("selbg", "celltext")
        cv.tag_lower("grid", "celltext")
        cv.tag_raise("head")

    def select_all(self, _event=None) -> str:
        if self.rows_n:
            cell = self.cells[self.rows_n - 1][self.ncols - 1]
            self._sel = ((0, 0, 0, 0),
                         (self.rows_n - 1, self.ncols - 1, len(cell) - 1,
                          len(cell[-1][0])))
            self._anchor = None
            self._draw_sel()
        return "break"

    def clear_sel(self) -> None:
        self._sel = ((0, 0, 0, 0), (0, 0, 0, 0))
        self._anchor = None
        self._draw_sel()

    def copy_selection(self, _event=None) -> str:
        text = self.sel_text()
        if text:
            self.clipboard_clear()
            self.clipboard_append(text)
        return "break"

    def copy_cell(self, r: int, c: int) -> None:
        """复制一格（保留作者换行；折行仍给空格）。"""
        if not (0 <= r < self.rows_n and 0 <= c < self.ncols):
            return
        cell = self.cells[r][c]
        text = "".join(line + (sep if sep else "") for line, sep in cell)
        self.clipboard_clear()
        self.clipboard_append(text.rstrip("\n"))

    def copy_all(self) -> None:
        """整张表（含表头）按 TSV 复制——粘进 Excel 就是一张表。"""
        def flat(text: str) -> str:
            return text.replace("\n", " ")
        lines = ["\t".join(flat(self.head_raw[c]) for c in range(self.ncols))]
        lines += ["\t".join(flat(self.body_raw[r][c]) for c in range(self.ncols))
                  for r in range(self.rows_n)]
        self.clipboard_clear()
        self.clipboard_append("\n".join(lines))

    # ------------------------------------------------------------ 鼠标事件
    def _on_press(self, event) -> None:
        self.canvas.focus_set()
        self._just_dbl = False
        pos = self.pos_at(self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        if pos is None:
            return
        if event.state & 0x0001 and self._anchor is not None:   # Shift：接着选
            self._sel = (self._anchor, pos)
        else:
            self._anchor = pos
            self._sel = (pos, pos)
        self._moved = False
        self._draw_sel()

    def _on_drag(self, event) -> None:
        if self._anchor is None:
            return
        pos = self.pos_at(self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        if pos is None or pos == self._sel[1]:
            return
        self._moved = True
        self._sel = (self._anchor, pos)
        self._draw_sel()

    def _on_release(self, _event) -> None:
        if self._just_dbl:
            self._just_dbl = False
            return
        if not self._moved:                # 单击（不是拖）：跟 Word 一样清掉选区
            self.clear_sel()

    def _on_double(self, event) -> None:
        got = self.hit(self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        if got is None or got[0] < 0:
            return
        r, c = got
        cell = self.cells[r][c]
        self._sel = ((r, c, 0, 0), (r, c, len(cell) - 1, len(cell[-1][0])))
        self._anchor = None
        self._just_dbl = True
        self._draw_sel()

    # ------------------------------------------------------------ 滚动
    def _on_yscroll(self, first, last) -> None:
        self.vsb.set(first, last)
        self._sync_head()

    def _sync_head(self) -> None:
        """表头钉在视口顶部（行往上滚，表头不动）。"""
        top = float(self.canvas.canvasy(0))
        dy = top - self._head_top
        if dy:
            self.canvas.move("head", 0, dy)
            self._head_top = top

    def can_scroll_y(self, step: int) -> bool:
        first, last = self.canvas.yview()
        if last - first >= 0.999 or step < 0 and first <= 0.0 or step > 0 and last >= 1.0:
            return False
        return True

    def scroll_y(self, pixels: int) -> None:
        self.canvas.yview_scroll(pixels, "units")

    def can_scroll_x(self, step: int) -> bool:
        first, last = self.canvas.xview()
        if last - first >= 0.999 or step < 0 and first <= 0.0 or step > 0 and last >= 1.0:
            return False
        return True

    def scroll_x(self, pixels: int) -> None:
        self.canvas.xview_scroll(pixels, "units")

    # ------------------------------------------------------------ 给外部读的事实
    @property
    def row_count(self) -> int:
        return self.rows_n

    @property
    def content_w(self) -> int:
        return 2 * GRID + sum(self.colw)

    def cell_text(self, r: int, c: int) -> str:
        """一格的原文字（不含我们折行的空格）。"""
        if 0 <= r < self.rows_n and 0 <= c < self.ncols:
            return self.body_raw[r][c]
        return ""

    def head_text(self, c: int) -> str:
        return self.head_raw[c] if 0 <= c < self.ncols else ""
