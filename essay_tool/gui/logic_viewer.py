"""「工具操作逻辑」浏览窗口——左树右文，跟查阅《规范》原文（spec_viewer.py）
是同一套交互结构，直接照搬过来改的。

跟 spec_viewer 的区别只在"结构从哪来"：
- spec_viewer 读的是 docx 文档，标题识别用引擎自己的 detect_headings
  （认的是论文里那套中文数字编号：一、／（一）／1.／（1））；
- 这里的内容来源是 `essay_tool.gui.operation_logic_text` 里的
  `OPERATION_LOGIC_TEXT`（**手工维护的源头**；那条"xlsx 生成"的链子
  2026-09-16 已停用），本身就是用户自己按"2." "2.1" "2.1.1"这种阿拉伯数字十进制
  编号写的，不需要也不适合套用那一套中文数字识别，直接按编号里点号的
  个数认层级即可（0 个点＝一级，1 个点＝二级，以此类推）。

左树右文、标题分级配色、行首禁则折行这几块跟 spec_viewer 完全一致，
直接复用同一套实现思路，没有另起一套。

**表格**（2026-09-16 加）：正文里可以直接写 Markdown 风格的管道表，渲染时会变成
tk 表格控件嵌在文字中间：

    | 识别对象 | 识别规则 |
    | --- | ---: |
    | 摘要 | 整段恰好是“摘要”二字 |

第一行是表头，第二行是分隔行（同时用冒号定对齐：`---` 左对齐、`:--:` 居中、
`---:` 右对齐），其余是数据行。分隔行不可省——没有它就不当成表格，
按普通正文显示（免得正文里偶尔出现的竖线被误认为表格）。
表格不参与折行，列宽按内容算，而且**按列数自动收窄**（目标是整张表不必横向
滚）；列算完还剩宽度就**全部给最右列**（用户 2026-09-17：“把滚动表格控件整个
拉宽，多出来的空间全部给最右侧的列”），外框钉在视口内；真放不下时表格自己带
横向滚动条，滚动条**真的能滚**（2026-09-17 修的——以前外框默认按内容撑开，
右半截被视口裁掉）。

表格控件 2026-09-17 从 `ttk.Treeview` 换成**自绘**（`essay_tool/gui/table_view.py`），
因为用户要的两条 Treeview 做不到：

  * **选中文字**：“我要的是选中文字，懂吗，选中文字复制，我想复制一个字就
    复制一个字那种”——自绘后拖选任意字符，`Ctrl+C` 复制选中文字；双击选中
    整格；`Ctrl+A` 选中整表；右键另有「复制选中文字／这一格／整张表」；
  * **逐行给高**：“每一行的字符需要多少行高，就给多少行高，类似『最合适的
    行高』”——Treeview 的行高是样式级、全表统一的，只能按全表最多行的那个
    格子给所有行留一样高，短行下面空一大块。

滚动也变成像素级（一格 `_WHEEL_PIXELS` 像素，Treeview 只能一整行一整行跳）；
表头纵向滚动时固定不动（`CanvasTable._sync_head`）。
表格的高度有上限（`_TABLE_MAX_H`），超出就在表格内部滚；不封顶的话，3.6 那张
39 行的表格能撑到 12000px 高，外壳正文每滚一格要重绘这么大一块嵌入控件，实测
**122 ms 一格**（用户 2026-09-16：“表格特别特别卡”）。封顶后实测 0.1 ms。
"""

from __future__ import annotations

import re
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk

from essay_tool.gui.table_view import CanvasTable
from essay_tool.gui.win_place import center_on

FAMILY = "Microsoft YaHei UI"
C_APP = "#f4f4f2"
C_BAR = "#e9e9e6"

# 行首禁则：这些标点不许作为一行的第一个字符（放不下时吊在上一行末尾）
_NO_LINE_START = set("，。、；：？！）〉》」』】”’｝］…—～·")
_NO_LINE_START.update("，.?!:;)]}%")

# 右侧文字区折行时两侧要扣掉的宽度（padx + 左边距 + 安全余量，够标点悬挂）
PANE_INSET = 32

# 标题颜色（与《规范》浏览窗口同一套配色，层级越深颜色越"暖"）
H_FG = {1: "#1a3e6e", 2: "#1f5c99", 3: "#2e7d32", 4: "#b06a00"}
# 可点交叉引用的颜色：跟全项目的链接蓝同一份取值（`flow_nav.C_LINK`）。
C_XREF = "#0563c1"
# 表格单元格里一行最多放到多宽（像素）：超过就在格子里折行。
# 实测（2026-09-16 截图）：Treeview 的格子能显示带换行的多行文字，所以长句子
# 不用截断；而一行太长（屏幕上没法读）也不合适，所以封顶。
# 2026-09-16 从 520 收到 420：用户要求"例外那一栏也可以窄一点"——降上限是最直接
# 的杆杆（每一列都按 min(自然宽, 上限) 撑开）。哪一列还需要单独限定的话，
# 再按列名加一张小表。
_TABLE_CELL_MAX_W = 420

# 每格折行宽度的**下限**（像素）。列多的时候按可用宽度均分，但不再窄过这个值
# （再窄就一行只能放两三个字，反而没法读）。到了下限还放不下就是横向滚动条的事。
_TABLE_CELL_MIN_W = 160

# 鼠标滚轮一格滚多少像素（正文的滚动由 `_on_table_wheel` 自己换算成 yview 比例）。
# 60 跟 Word 的手感接近；太大就"一滚就飞"，太小又显得黏。
_WHEEL_PIXELS = 60

# 表格控件的可见高度上限（像素）。超出就在表格内部纵向滚动。
# 这个数字是性能与可读性的折衷：不封顶时 3.6 那张表有 12000px 高，正文
# 每滚一格要重绘这么大一块嵌入控件（实测 122 ms/格）。实测扫描（1/2/3/4/6/8 行）：
# 343px→4.5ms、959px→0.9ms、1267px→1.3ms、1883px→3.8ms、2499px→4.5ms
# ——2500px 以内都很轻快，两三千像素才是分水岭。所以取 1300px：3.6 那种
# 308px 一行的大行能露 4 行，普通的 130px 行能露 9 行，氪金才 1.3ms/格。
_TABLE_MAX_H = 1300
# 表头那一行占的高度（估的：10pt 加粗 + 上下边距），算行高上限时要扣掉。
# 竖向滚动条占的宽度（Windows 上 ttk 滚动条约 17px，实测）。列宽算"铺满"时
# 要先把这块让出来，否则最后一列的右边缘会被滚动条压住。
_VSCROLL_W = 17
# 标题行：形如"2." "2.1" "2.1.1"开头，后面跟空白再跟标题正文。
# 点号个数决定层级：0 个点＝一级（"2. 导入和自动备份"），1 个点＝二级
# （"2.1 导入"），以此类推。
_HEADING_RE = re.compile(r"^(\d+(?:\.\d+)*)\.?\s+(\S.*)$")

# 正文里的**交叉引用**（用户 2026-09-17：“允许说明的交叉引用点击跳转”）：
# 说明正文里已经写着 `参见 3.4 识别范围`、`详见 6.3` 这种写法，认“见/参见/详见”
# 后面紧跟的编号即可。**只认这一种写法**：光一个 `3.4` 不做链接——正文里编号
# 出现得太密，全标蓝反而没重点。
#
# 标成链接的是**编号**（原文若紧跟着该节标题，一并标上），不含“参见/详见”本身
# ——用户 2026-09-17：“能点击的是 3.4 识别范围，而不是参见 3.4”。
_XREF_RE = re.compile(r"(?:参见|详见|见)\s*(\d+(?:\.\d+)*)")

# 表格：整行以竖线开头、以竖线结尾（允许两侧带空白）
_TABLE_ROW_RE = re.compile(r"^\|.*\|$")
# 分隔行：|---|---|、|:--|--:|、|:--:| 都算；每个格子至少两个短横
_TABLE_SEP_RE = re.compile(r"^\|(?:\s*:?-{2,}:?\s*\|)+$")


def _split_table_row(raw: str) -> list[str]:
    """把一个管道行切成格子（去掉首尾竖线，去掉每格两端空白）。

    格子里写的字面量 `\\n` 换成**真换行**（用户 2026-09-16："所有表格里的
    换行全都丢了，带来严重可读性问题"）. 以前换成空格，是因为担心 Treeview
    不支持多行，后来实测它能排成多行（见 `_embed_table`），所以换成真换行。
    """
    text = raw.strip()
    if text.startswith("|"):
        text = text[1:]
    if text.endswith("|"):
        text = text[:-1]
    return [cell.replace("\\n", "\n").strip() for cell in text.split("|")]


def parse_table(lines: list[str], start: int, stop: int) -> tuple[list, list, int] | None:
    """从 `lines[start]` 起解析一张表，到 `stop`（不含）为止。

    返回 (表头, 数据行, 结束行号)；不构成表格时返回 None。判据：
      * 首行必须是管道行；
      * 第二行必须是分隔行（|--|--|）；
      * 遇到的第一个不是管道行的行就是表格的结束。
    分隔行里的冒号决定该列对齐：`---` 左对齐、`:--:` 居中、`---:` 右对齐。
    """
    if start >= stop or not _TABLE_ROW_RE.match(lines[start].strip()):
        return None
    rows: list[list[str]] = []
    i = start
    while i < stop and _TABLE_ROW_RE.match(lines[i].strip()):
        rows.append(_split_table_row(lines[i]))
        i += 1
    if len(rows) < 2 or not _TABLE_SEP_RE.match(lines[start + 1].strip()):
        return None
    ncols = len(rows[0])
    rows = [r + [""] * (ncols - len(r)) if len(r) < ncols else r[:ncols] for r in rows]
    return rows[0], rows[2:], i


def _column_aligns(sep_row: list[str]) -> list[str]:
    """分隔行里的冒号 → 对齐方式（'w' 左、'center' 居中、'e' 右）。"""
    out: list[str] = []
    for cell in sep_row:
        token = cell.strip()
        if token.startswith(":") and token.endswith(":"):
            out.append("center")
        elif token.endswith(":"):
            out.append("e")
        else:
            out.append("w")
    return out


def build_logic_nodes(text: str) -> tuple[list[dict], list[str]]:
    """把说明文本切成"节"清单：{level, text, start, end}，行号按 splitlines() 算。"""
    lines = text.splitlines()
    nodes: list[dict] = []
    for i, raw in enumerate(lines):
        stripped = raw.strip()
        m = _HEADING_RE.match(stripped)
        if not m:
            continue
        level = m.group(1).count(".") + 1
        nodes.append({"level": level, "text": stripped, "start": i})

    if not nodes:
        nodes = [{"level": 1, "text": "（未识别出结构）", "start": 0}]

    for k, nd in enumerate(nodes):
        end = len(lines)
        for nxt in nodes[k + 1:]:
            if nxt["level"] <= nd["level"]:
                end = nxt["start"]
                break
        nd["end"] = end
    return nodes, lines


class LogicWindow(tk.Toplevel):
    """「工具操作逻辑」浏览窗口：左树右文，树状索引跳转。

    `anchor`：打开时直接跳到哪一节（用户 2026-09-17：结构页那个按钮要直接跳
    3.6）。传的是**编号**（如 `"3.6"`），找不到就安安静静留在开头（不报错）。
    """

    def __init__(self, master: tk.Misc, text: str, anchor: str | None = None) -> None:
        super().__init__(master)
        self.title("工具操作逻辑")
        self.configure(bg=C_APP)
        try:
            self.state("zoomed")
        except Exception:
            self.geometry("1280x840")
            # 不能最大化时（比如非 Windows）也要摆在主窗口正中央，
            # 不能落在屏幕左上角（用户 2026-09-15）。
            center_on(self, master, 1280, 840)
        self.minsize(960, 600)

        self.nodes, self.lines = build_logic_nodes(text)
        self.heading_lv = {nd["start"]: nd["level"] for nd in self.nodes}

        # ---- 顶栏 ----
        top = tk.Frame(self, bg=C_BAR)
        top.pack(side="top", fill="x")
        tk.Label(top, text="工具操作逻辑", font=(FAMILY, 12, "bold"),
                 bg=C_BAR, fg="#333").pack(side="left", padx=(12, 6), pady=6)

        # ---- 主体：左树右文 ----
        paned = tk.PanedWindow(self, orient="horizontal", sashwidth=6, bg=C_APP,
                               sashrelief="flat")
        paned.pack(side="top", fill="both", expand=True)

        left = tk.Frame(paned, bg="#ffffff")
        style = ttk.Style(self)
        style.configure("Logic.Treeview", font=(FAMILY, 10), rowheight=32,
                        background="#ffffff", fieldbackground="#ffffff")
        self.tree = ttk.Treeview(left, show="tree", style="Logic.Treeview",
                                 selectmode="browse")
        vsb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(left, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)
        paned.add(left, width=380, minsize=240)

        right = tk.Frame(paned, bg="#ffffff")
        self.text = tk.Text(right, wrap="none", state="disabled", bg="#ffffff",
                            relief="flat", padx=8, pady=4, font=(FAMILY, 11))
        tvs = ttk.Scrollbar(right, orient="vertical", command=self.text.yview)
        # 横向滚动条：表格比窗口宽、以及特别长的不可断行时用得上
        ths = ttk.Scrollbar(right, orient="horizontal", command=self.text.xview)
        self.text.configure(yscrollcommand=tvs.set, xscrollcommand=ths.set)
        self.text.grid(row=0, column=0, sticky="nsew")
        tvs.grid(row=0, column=1, sticky="ns")
        ths.grid(row=1, column=0, sticky="ew")
        right.rowconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)
        paned.add(right, minsize=400)

        # 正文区样式（显示用）；字体对象同时拿来做折行测量。
        self._fonts = {
            "body": tkfont.Font(family=FAMILY, size=11),
            "h1": tkfont.Font(family=FAMILY, size=15, weight="bold"),
            "h2": tkfont.Font(family=FAMILY, size=13, weight="bold"),
            "h3": tkfont.Font(family=FAMILY, size=12, weight="bold"),
            "h4": tkfont.Font(family=FAMILY, size=12, weight="bold"),
            "table": tkfont.Font(family=FAMILY, size=10),   # 表格格子里的字（量宽用）
        }
        self._cw: dict[str, dict[str, float]] = {}     # 单字宽度缓存（按样式分）
        self._current_node: int | None = None
        self._text_job: str | None = None
        self._tables: list[tk.Widget] = []             # 当前渲染出的表格外框
        self._last_table = None                        # 最后嵌入的表格（复制的兜底句柄）
        # 表格右键菜单：**一次建好**，每次右键只更新 `_menu_target` 指到哪个表/
        # 哪一格（每弹一次新建一个 tk.Menu 会攒控件——Tk 的 Menu 不会自己回收）。
        self._menu_target: tuple = (None, None)
        self._table_menu = tk.Menu(self, tearoff=0)
        self._table_menu.add_command(label="复制选中文字", command=self._copy_menu_sel)
        self._table_menu.add_command(label="复制这一格（含换行）", command=self._copy_menu_cell)
        self._table_menu.add_command(label="复制整张表（含表头）", command=self._copy_menu_table)
        self._pixels_cache: int = 0                    # 正文总像素高（滚轮换算用，见 _text_pixels）
        self._style = style                            # ttk.Style，建表格样式用
        self._xref_tags: set[str] = set()              # 已配置过的引用标签名
        self.text.bind("<Configure>", self._on_text_resize)
        # 鼠标移到链接上变手型、点链接跳转（Tk 的 Text **没有** tag 级的 cursor
        # 选项，只能自己在事件里换控件光标）。
        self.text.bind("<Motion>", self._on_text_motion)
        self.text.bind("<Button-1>", self._on_text_click)

        self.text.tag_configure("h1", font=self._fonts["h1"], foreground=H_FG[1],
                                spacing1=12, spacing3=8)
        self.text.tag_configure("h2", font=self._fonts["h2"], foreground=H_FG[2],
                                spacing1=10, spacing3=6)
        self.text.tag_configure("h3", font=self._fonts["h3"], foreground=H_FG[3],
                                spacing1=9, spacing3=5)
        self.text.tag_configure("h4", font=self._fonts["h4"], foreground=H_FG[4],
                                spacing1=9, spacing3=5)
        self.text.tag_configure("body", font=self._fonts["body"], spacing3=7,
                                lmargin1=8, lmargin2=8)

        # 表格：跟左树同一套观感（小一号字、斑马纹、浅边框）。
        # ⚠️ 自定义样式名**必须以控件类名结尾**（这里 Treeview），否则 ttk 找不到
        # 布局，建控件时直接报 "Layout Logic.Table not found"（实测踩过）。
        style.configure("LogicTable.Treeview", font=(FAMILY, 10), rowheight=26,
                        background="#ffffff", fieldbackground="#ffffff")
        style.configure("LogicTable.Treeview.Heading", font=(FAMILY, 10, "bold"))
        self.font_table = tkfont.Font(family=FAMILY, size=10)
        self.font_table_head = tkfont.Font(family=FAMILY, size=10, weight="bold")

        self.tree.tag_configure("lv1", foreground=H_FG[1], font=(FAMILY, 10, "bold"))
        self.tree.tag_configure("lv2", foreground=H_FG[2])
        self.tree.tag_configure("lv3", foreground=H_FG[3])
        self.tree.tag_configure("lv4", foreground=H_FG[4])

        # 溢出控制：标题超过面板宽度就截断加"…"（全文在右侧可见）
        self.font_tree = tkfont.Font(family=FAMILY, size=10)
        self.font_tree_bold = tkfont.Font(family=FAMILY, size=10, weight="bold")
        self._label_job: str | None = None
        self.tree.bind("<Configure>", self._on_tree_resize)

        self._fill_tree()
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        roots = self.tree.get_children()
        if roots:
            self.tree.selection_set(roots[0])
            self.tree.focus(roots[0])
            self._show_node(int(roots[0]))
        # 调用方指定了要直接看的那一节（比如结构页的「查阅识别规则」→ 3.6）
        if anchor:
            self.goto_number(anchor)
        # 句柄必须登记进 `_label_job`：不然窗口在这 80ms 内被关掉，回调会对
        # 已销毁的树操作（同文件 _on_tree_resize 就是按这条教训写的）。
        self._label_job = self.after(80, self._refresh_tree_labels)

    # ------------------------------------------------------------ 结构树

    def _fill_tree(self) -> None:
        parents: dict[int, str] = {}   # level -> 最近一个该级节点的 iid
        for k, nd in enumerate(self.nodes):
            lv = nd["level"]
            parent = ""
            for upper in range(lv - 1, 0, -1):
                if upper in parents:
                    parent = parents[upper]
                    break
            self.tree.insert(parent, "end", iid=str(k), text=nd["text"],
                             tags=(f"lv{min(lv, 4)}",), open=True)
            parents[lv] = str(k)
            for deeper in [x for x in list(parents) if x > lv]:
                del parents[deeper]

    # ------------------------------------------------------------ 溢出控制

    def _on_tree_resize(self, _event=None) -> None:
        if self._label_job is not None:
            self.after_cancel(self._label_job)
        self._label_job = self.after(120, self._refresh_tree_labels)

    def _refresh_tree_labels(self) -> None:
        """按面板宽度把过长的标题截成"…"（树里只放得下一行）。"""
        self._label_job = None
        if not self.tree.winfo_exists():
            return                      # 窗口已经关了
        width = self.tree.winfo_width()
        if width <= 1:
            return
        for k, nd in enumerate(self.nodes):
            lv = nd["level"]
            font = self.font_tree_bold if lv == 1 else self.font_tree
            avail = max(80, width - 34 - (lv - 1) * 20)
            self.tree.item(str(k), text=self._ellipsize(nd["text"], avail, font))

    @staticmethod
    def _ellipsize(text: str, width: float, font: tkfont.Font) -> str:
        if font.measure(text) <= width:
            return text
        lo, hi = 0, len(text)
        while lo < hi:
            mid = (lo + hi + 1) // 2
            if font.measure(text[:mid] + "…") <= width:
                lo = mid
            else:
                hi = mid - 1
        return text[:lo] + "…"

    def _on_tree_select(self, _event=None) -> None:
        sel = self.tree.selection()
        if sel:
            self._show_node(int(sel[0]))

    # ------------------------------------------------------- 按编号跳转 / 交叉引用

    def _node_of_number(self, num: str) -> int | None:
        """编号（`"3.6"`）对应的那一节在 `self.nodes` 里的下标；没有返回 None。

        比的是**标题第一个词**（`"3.6 识别"` → `3.6`），不是全串包含——否则
        `3.6` 会先命中 `3.6.1` 那一行的开头。
        """
        num = (num or "").strip().rstrip(".")
        if not num:
            return None
        for k, nd in enumerate(self.nodes):
            tokens = nd["text"].split()
            if tokens and tokens[0].rstrip(".") == num:
                return k
        return None

    def _section_title(self, node: int) -> str:
        """某一节的标题正文（`"3.4 识别范围"` → `"识别范围"`）；没标题就返回空串。"""
        parts = self.nodes[node]["text"].split(None, 1)
        return parts[1].strip() if len(parts) > 1 else ""

    def goto_number(self, num: str) -> bool:
        """跳到编号为 num 的那一节（如 `"3.6"`）；跳成返回 True。

        找不到就原样待着、返回 False——按钮文案与锚点写错时不应该弹错。
        """
        k = self._node_of_number(num)
        if k is None:
            return False
        self.tree.selection_set(str(k))     # 选中会走 _on_tree_select 再渲染一次，
        self.tree.focus(str(k))             # 但显式调用 _show_node 更稳（幂等）
        self.tree.see(str(k))
        self._show_node(k)
        return True

    def _xref_at(self, event) -> str | None:
        """鼠标位置下的交叉引用标签名（`xref:3.4`）；没压在链接上就返回 None。"""
        try:
            names = self.text.tag_names(f"@{event.x},{event.y}")
        except tk.TclError:
            return None
        return next((n for n in names if n.startswith("xref:")), None)

    def _on_text_motion(self, event) -> None:
        """鼠标在正文里移动：压在交叉引用上就给个手型光标。"""
        want = "hand2" if self._xref_at(event) else "xterm"
        if str(self.text.cget("cursor")) != want:
            self.text.configure(cursor=want)

    def _on_text_click(self, event) -> str | None:
        """点交叉引用就跳到那一节去。

        ⚠️ **没有用 `tag_bind`**：正文是 `state="disabled"` 的，而实测 Tk 在
        禁用态**不派发 tag 级鼠标绑定**（控件级的 Motion 照常收得到、手型光标
        也换得动，但 `tag_bind("<Button-1>")` 一次都不触发——2026-09-17 用
        `tools/_check_logic_xref.py` 打出来的真点击验证）。所以绑定挂在控件上、
        自己按坐标认标签；标签只管“长什么样”（蓝字下划线）。
        """
        tag = self._xref_at(event)
        if not tag:
            return None
        self.goto_number(tag.split(":", 1)[1])
        return "break"

    def _tag_xrefs(self, t: tk.Text, raw: str, from_index: str) -> None:
        """把这一行正文里的交叉引用标成可点链接（蓝字 + 下划线）。

        标的是**编号**；原文若在编号后面紧跟着该节的标题，一并标上（`参见 3.4 识别范围`
        ⇒ 标“3.4 识别范围”）。**不标“参见/详见”**——用户 2026-09-17：“能点击的是
        3.4 识别范围，而不是参见 3.4”。

        ⚠️ 折行是**我们自己做的**（`_wrap`），所以不能按原串偏移算位置——那串
        文字插进控件后断在哪里不好推。改成在刚插进去的那一段里**按编号搜**：
        编号被折行切断时才搜不到（那时整段也搜不到），跳过就是。标题被折到
        下一行时，链接就只剩编号——宁可短一点，也不能把蓝字标到别的词上。
        """
        cursor = from_index
        for match in _XREF_RE.finditer(raw):
            num = match.group(1)
            node = self._node_of_number(num)
            if node is None:
                continue                    # 说明里没这一节：不做链接
            hit = t.search(num, cursor, "end", exact=True, forwards=True)
            if not hit:
                continue
            end = f"{hit}+{len(num)}c"
            title = self._section_title(node)
            if title and t.get(end, f"{end}+{len(title) + 1}c") == f" {title}":
                end = f"{end}+{len(title) + 1}c"
            tag = f"xref:{num}"
            if tag not in self._xref_tags:
                # 标签只管外观；点击由控件级 `<Button-1>` 按坐标派发
                # （禁用态不派发 tag 级鼠标绑定，见 `_on_text_click`）
                t.tag_configure(tag, foreground=C_XREF, underline=True)
                self._xref_tags.add(tag)
            t.tag_add(tag, hit, end)
            # 游标往后推：同一行重复出现的引用才不会都标到第一处
            cursor = end

    def _show_node(self, k: int) -> None:
        self._current_node = k
        self._render_current()

    # ------------------------------------------------------------ 折行（行首禁则）

    def _on_text_resize(self, _event=None) -> None:
        if self._text_job is not None:
            self.after_cancel(self._text_job)
        self._text_job = self.after(120, self._render_current)

    def _render_current(self) -> None:
        """按当前宽度手动折行后铺进右侧（宽度变了自动重排）。"""
        self._text_job = None
        k = self._current_node
        if k is None:
            return
        if not self.text.winfo_exists():
            return                      # 窗口已经关了，别再往里画
        width = self.text.winfo_width()
        if width < 80:
            # 还没排好版，稍后再来。**这个 after 也要登记进 `_text_job`**
            # （审计 2026-09-15 核实：以前它是个没人管的匿名句柄，窗口在这
            # 80ms 内被关掉就会报 "invalid command name"）。
            self._text_job = self.after(80, self._render_current)
            return
        avail = width - PANE_INSET
        nd = self.nodes[k]
        t = self.text
        # 正文要重画了：滚轮用的"总像素高"缓存跟着失效（见 `_text_pixels`）。
        self._pixels_cache = 0
        t.configure(state="normal")
        t.delete("1.0", "end")
        for widget in self._tables:                # 表格是嵌进正文的控件，得自己清
            if widget.winfo_exists():
                widget.destroy()
        self._tables.clear()
        i, stop = nd["start"], nd["end"]
        while i < stop:
            raw = self.lines[i].strip()
            if not raw:
                i += 1
                continue
            table = parse_table(self.lines, i, stop)
            if table is not None:
                header, body, next_line = table
                aligns = _column_aligns(_split_table_row(self.lines[i + 1].strip()))
                self._embed_table(t, header, aligns, body, avail)
                i = next_line
                continue
            lv = self.heading_lv.get(i)
            tag = f"h{min(lv, 4)}" if lv else "body"
            xref_from = t.index("end-1c")
            for line in self._wrap(raw, avail, tag):
                t.insert("end", line + "\n", tag)
            if not lv:                      # 只给正文行挂引用链接（标题行没有）
                self._tag_xrefs(t, raw, xref_from)
            i += 1
        t.configure(state="disabled")
        t.yview_moveto(0)
        t.xview_moveto(0)

    def _embed_table(self, t: tk.Text, header: list[str], aligns: list[str],
                     body: list[list[str]], avail: float) -> None:
        """把一张表做成**自绘控件**嵌进正文里（见 `gui/table_view.py`）。

        2026-09-17 从 `ttk.Treeview` 换成自绘，因为用户要的两条 Treeview 做不到：
          * **字符级选中**（"我要的是选中文字，懂吗，选中文字复制，我想复制一个字
            就复制一个字"）；
          * **逐行给高**（"每一行的字符需要多少行高，就给多少行高，类似『最合适的
            行高』"）——Treeview 的行高是**样式级、全表统一**的，只能按"全表最多行的
            那个格子"给所有行留一样高，短行下面空一大块。
        其余口径照旧：折行宽度按"可用宽 ÷ 列数"收窄（下限 `_TABLE_CELL_MIN_W`），
        剩余宽度全给最右列（用户 2026-09-17），外框钉在可用宽度上；高度超过
        `_TABLE_MAX_H` 时内部按**像素**滚（比 Treeview 的"一格一整行"细腻）。
        """
        table = CanvasTable(
            t, head=header, aligns=aligns, body=body, avail=avail,
            # 折行仍走正文那一套（标点不落行首、交叉引用不被切断）
            wrap=lambda text, width: self._wrap(text, width, "table"),
            font=self.font_table, head_font=self.font_table_head,
            line_h=self.font_table.metrics("linespace"),
            max_h=_TABLE_MAX_H, cell_min_w=_TABLE_CELL_MIN_W,
            cell_max_w=_TABLE_CELL_MAX_W, vscroll_w=_VSCROLL_W,
            on_wheel=self._on_table_wheel, on_menu=self._popup_table_menu)
        self._tables.append(table)
        self._last_table = table
        t.window_create("end", window=table, pady=2)
        t.insert("end", "\n\n", "body")

    # ------------------------------------------------------ 表格：复制（2026-09-17）
    def _clipboard_set(self, text: str) -> None:
        """把文字放进剪贴板（Ctrl+C 与右键菜单共用）。"""
        self.clipboard_clear()
        self.clipboard_append(text)

    def _table_of(self, event):
        """事件落在哪张表上（往上找 `CanvasTable`；找不到就用最后嵌的那张）。"""
        widget = getattr(event, "widget", None)
        while widget is not None:
            if isinstance(widget, CanvasTable):
                return widget
            widget = getattr(widget, "master", None)
        if self._last_table is not None and self._last_table.winfo_exists():
            return self._last_table
        return None

    def _popup_table_menu(self, event) -> str:
        """右键：复制选中文字 / 复制这一格 / 复制整张表。"""
        table = self._table_of(event)
        if table is None:
            return "break"
        got = table.hit(table.canvas.canvasx(event.x), table.canvas.canvasy(event.y))
        self._menu_target = (table, got)
        try:
            self._table_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self._table_menu.grab_release()
        return "break"

    def _menu_table(self):
        table = self._menu_target[0]
        if table is None or not table.winfo_exists():
            return None
        return table

    def _copy_menu_sel(self) -> None:
        """菜单「复制选中文字」：跟 Ctrl+C 同一份实现。"""
        table = self._menu_table()
        if table is not None:
            table.copy_selection()

    def _copy_menu_cell(self) -> None:
        """菜单「复制这一格」：右键点在表头上就复制表头文字。"""
        table = self._menu_table()
        got = self._menu_target[1]
        if table is None or got is None:
            return
        row, col = got
        if row < 0:
            text = table.head_text(col)
            if text:
                self._clipboard_set(text)
        else:
            table.copy_cell(row, col)

    def _copy_menu_table(self) -> None:
        """菜单「复制整张表」：带表头，粘进 Excel 就是一张表。"""
        table = self._menu_table()
        if table is not None:
            table.copy_all()

    def _on_table_wheel(self, event) -> str:
        """滚轮落在嵌入的表格上时怎么办（返回 "break" 吃掉事件）。

        三种情况：
          * **Shift+滚轮**：横向滚这张表。表格自己的横向滚动条在表格**底部**——
            一张 1300px 高的表，用户得先往下滚一屏才摸得到它；Shift+滚轮在鼠标停在
            表格上时就能横着走（全项目 Shift+滚轮都是横向，见结构页）。
          * 表格被 `_TABLE_MAX_H` 截短了、而且还没滚到头：滚**表格自己**（自绘以后
            按像素滚，一格 `_WHEEL_PIXELS` 像素；Treeview 时代只能一整行一整行跳）。
          * 其它情况：滚**外面的正文**。

        滚正文**不能按行滚**（用户 2026-09-16："滚轮滚动很卡，且会直接从开头跳到
        结尾"）：正文里嵌入的表格是整段"窗口"，Text 按行计位置时那一行就是整张表
        那么高，滚一格就飞过去一大截。而 Text 又没有 `yscrollincrement` 这种选项
        （那是 Canvas/Listbox 的，写成 Text 参数会直接报 unknown option——踩过），
        所以改成**自己把像素换算成 yview 比例**：量出正文总像素高，一格走
        `_WHEEL_PIXELS` 像素对应的那一小截（`_text_pixels` 带缓存）。
        """
        step = -1 if event.delta > 0 else 1
        table = self._table_of(event)
        if table is not None and (event.state & 0x0001):
            if table.can_scroll_x(step):
                table.scroll_x(step * _WHEEL_PIXELS)
                return "break"
        if table is not None and table.can_scroll_y(step):
            table.scroll_y(step * _WHEEL_PIXELS)
            return "break"
        total = self._text_pixels()
        if total <= 0:
            self.text.yview_scroll(step, "units")
            return "break"
        current = self.text.yview()[0]
        target = min(max(current + step * _WHEEL_PIXELS / total, 0.0), 1.0)
        self.text.yview_moveto(target)
        return "break"

    def _text_pixels(self) -> int:
        """正文的总高度（像素）——滚轮按像素走要靠它换算。

        ⚠️ **必须缓存**：`count … -ypixels` 会把整篇重新量一遍像素，正文里嵌着
        一张 14000px 高的表格时，量一次要 **128 ms**——每次滚轮都量就等于
        "表格特别特别卡"（用户 2026-09-16 反馈，实测 128 ms/格）。重画正文时
        （`_render_current`）把缓存清掉。
        """
        if self._pixels_cache:
            return self._pixels_cache
        try:
            got = self.text.count("1.0", "end", "ypixels")
        except Exception:                                             # noqa: BLE001
            return 0
        if isinstance(got, tuple):
            got = got[0] if got else 0
        self._pixels_cache = int(got or 0)
        return self._pixels_cache

    def _cwidth(self, tag: str, ch: str) -> float:
        cache = self._cw.setdefault(tag, {})
        w = cache.get(ch)
        if w is None:
            w = self._fonts[tag].measure(ch)
            cache[ch] = w
        return w

    def _wrap(self, text: str, avail: float, tag: str) -> list[str]:
        """手动折行（行首禁则版，折行只在显示上用，不改文字）：

        - 标点不许作为一行的第一个字符——放不下时吊在上一行末尾（含连续标点，
          如"。"和"——"）；
        - 英文单词尽量不从中间断开（优先退到最近的空格处）；
        - **交叉引用（`参见 3.4`）整段不许被切断**：切开了就成不了链接
          （连“搜文字标蓝”都搜不到，实测窄窗口下 `参见 3.` / `4 识别范围`）。
          宁可让那一行略微溢出（正文本来就有横向滚动条），也不切链接。
        """
        protected: set[int] = set()
        for match in _XREF_RE.finditer(text):
            # 只保护区间**内部**的位置：引用开头之前仍然可以断行（链接整个
            # 落到下一行去）
            protected.update(range(match.start() + 1, match.end()))
        lines: list[str] = []
        cur = ""
        w = 0.0
        i, n = 0, len(text)
        while i < n:
            ch = text[i]
            if not cur and ch == " ":
                i += 1
                continue                          # 行首不要空格
            cw = self._cwidth(tag, ch)
            if w + cw <= avail or not cur:
                cur += ch
                w += cw
                i += 1
                continue
            # 放不下了，该断行——
            if i in protected:
                # 断在这里会把交叉引用切开：整段一起塞进本行（允许略微溢出）
                while i < n and i in protected:
                    cur += text[i]
                    w += self._cwidth(tag, text[i])
                    i += 1
                continue
            if ch in _NO_LINE_START:
                cur += ch                         # 行首禁则：标点吊在行尾
                i += 1
                while i < n and text[i] in _NO_LINE_START:
                    cur += text[i]                # 连着的标点整串一起吊
                    i += 1
                lines.append(cur)
                cur, w = "", 0.0
                continue
            if ch.isascii() and ch.isalnum() and cur[-1].isascii() and cur[-1].isalnum():
                sp = cur.rfind(" ")
                if sp >= 0:                       # 英文单词不拦腰断
                    lines.append(cur[:sp + 1])
                    cur = cur[sp + 1:] + ch
                    w = sum(self._cwidth(tag, c) for c in cur)
                    i += 1
                    continue
            lines.append(cur)
            cur, w = ch, cw
            i += 1
        if cur:
            lines.append(cur)
        return lines or [""]

# 演示文本（`tools\_demo_logic_viewer.py` 用）：只是为了让用户看一眼表格长什么样，
# 不参与正式文档。真正写进《工具操作逻辑》里的表格，只要按同样的管道写法即可。
DEMO_TEXT = """1. 表格演示

1.1 识别规则表
下面这张表是嵌在正文里的真表格。列宽按内容实测宽度算，表头用加粗字；表格比窗口宽时，表格下方会出现它自己的横向滚动条，正文的排版不受影响。

| 识别对象 | 识别规则 | 允许的例外 | 样式应用 |
|---|---|---|---|
| 题目 | 识别范围开始第一个（非空）且（不以标点结尾）且（未被识别为其他部件）的段落 | 设了封面范围时从封面之后开始 | 直接在段落上按预设格式调整 |
| 摘要 | 整段恰好是“摘要”二字 | 允许二字外有任意的“【”“】”“[”“]”，不要求成对 | 按预设格式调整，并把内容改成【摘要】 |
| 图题 | 以“图”字后接可选的空白和数字开头，或段落中存在编号域 | 段落以标点结尾时，不识别为图题 | 套用预设样式，并收录进插图清单 |
| 脚注 | 不参与自动识别，文档中存在的脚注段落即脚注 | 文档中不存在脚注时不建立脚注样式 | 清除脚注段落已有的直接格式，套用预设样式 |

1.2 表格之外的正文
表格前后仍然是普通文字，折行规则与以前完全一致：标点不落到行首，英文单词不从中间断开。表格本身不参与折行，格子里的文字再长也由表格自己的滚动条处理。

1.3 列对齐
每一列的对齐方式单独设置：分隔行写 `---` 是左对齐，`:--:` 是居中，`---:` 是右对齐。

| 级别 | 名称 | 字数 |
|:---|:---:|---:|
| 一级 | 引言 | 2 |
| 二级 | （一）研究背景 | 6 |
| 三级 | 1. 研究对象与范围 | 9 |

1.4 分隔行不可省
没有分隔行的管道写法不会被当成表格，会照原样显示为正文（防止正文里偶尔出现的竖线被误解成表格）。
"""


def open_logic_viewer(master: tk.Misc, text: str,
                      anchor: str | None = None) -> LogicWindow:
    """主 GUI 的"查阅本工具操作逻辑说明"按钮调这个：弹出（或唤醒）说明窗口。

    `anchor`：要直接看的那一节编号（如 `"3.6"`）。窗口已经开着的时候，`anchor`
    会让它**就在当前窗口里跳过去**（用户 2026-09-17：结构页那个按钮要直接跳
    3.6，而不是从头开始翻）。
    """
    win = getattr(master, "_logic_viewer_win", None)
    if win is not None and win.winfo_exists():
        win.deiconify()
        win.lift()
        win.focus_set()
        if anchor:
            win.goto_number(anchor)
        return win
    win = LogicWindow(master, text, anchor=anchor)
    setattr(master, "_logic_viewer_win", win)
    return win
