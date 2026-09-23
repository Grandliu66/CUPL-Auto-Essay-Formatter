"""《规范》原文浏览窗口（2026-09-13；已接入主窗口「关于」页）。

用户需求：从主 GUI 点"打开规范文件"后弹窗，用树结构呈现《规范》原文
（**只整理结构、不动原文**），并允许把原版 .docx 另存出来。

设计要点：
- 结构直接用引擎的 detect_headings 认（规范文档本身就是标准四级编号：
  一、／（一）／1.／（1））；引擎的"标点结尾/过长就归正文"过滤器在浏览时
  不启用——编号模式命中就算一节（例：规范自己的"8. 图表和表达式的用法。"
  结尾带句号，不能让它从结构里消失）；
- 目录页那些"标题\\t页码"的行会被剔除——它们是目录条目，不是真标题；
- 封面/目录/说明这些"开头部分"不单独成节点（用户 2026-09-13：不要）；
- 左侧结构树（四级嵌套），右侧显示所选章节的原文，逐段照抄、一字未改；
- 树里的长标题按面板宽度截断加"…"，完整标题在右侧与底栏可见；
- 右侧原文按"行首禁则"手动折行：标点不许起行首（放不下就吊在上一行尾），
  英文单词不拦腰断；面板宽度变化自动重排（折行只是显示，不改文字）；
- 每一节的范围＝从它的标题段起，直到下一个"级别不比自己深"的标题；
- "另存原版 .docx…"＝把原文件原样复制到用户指定位置（导出的是未动过的原件）。

用法（也可单独跑）：
    .venv\\Scripts\\python.exe essay_tool\\gui\\spec_viewer.py            # 默认找 reference 里的规范
    .venv\\Scripts\\python.exe essay_tool\\gui\\spec_viewer.py --spec 路径
    .venv\\Scripts\\python.exe essay_tool\\gui\\spec_viewer.py --check    # 只打印结构树，不开窗

给主 GUI 接线：
    from essay_tool.gui.spec_viewer import open_spec_viewer, open_star10_viewer
    open_spec_viewer(self.root, spec_path)      # 按钮回调里这一行即可（重复点击会唤醒已有窗口）
    open_star10_viewer(self.root)               # 十星「格式要求」按钮（文本内置，见 star10_spec_text.py，
                                                #  开发期用 tools/make_star10_spec_text.py 从官方 PDF 生成）
"""

from __future__ import annotations

import argparse
import dataclasses
import pathlib
import re
import shutil
import sys
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

ROOT = pathlib.Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from docx import Document  # noqa: E402

from essay_tool.config.integrity import (OFFICIAL_FILE_PROBLEM, IntegrityError,  # noqa: E402
                                         verify_official)
from essay_tool.config.paths import resource_dir  # noqa: E402
from essay_tool.gui.win_place import center_on  # noqa: E402

from essay_tool.engine.detect import _try_numbered, detect_headings  # noqa: E402

FAMILY = "Microsoft YaHei UI"
REFERENCE_DIR = resource_dir() / "reference"      # 打包后在 _internal 里

C_APP = "#f4f4f2"
C_BAR = "#e9e9e6"

# 目录条目的标志：行尾是"制表符 + 页码"
_TOC_LINE = re.compile(r"\t\s*\d+\s*$")

# 十星内置文本里"整篇大块"的标题（不是编号写法）：附录一：… / 附录二：…
_APPENDIX_HEAD = re.compile(r"^附录[一二三四五六七八九十\d]+[：:、]")

# 行首禁则：这些标点不许作为一行的第一个字符（放不下时吊在上一行末尾）
_NO_LINE_START = set("，。、；：？！）〉》」』】”’｝］…—～·")
_NO_LINE_START.update("，.?!:;)]}%")

# 右侧文字区折行时两侧要扣掉的宽度（padx + 左边距 + 安全余量，够标点悬挂）
PANE_INSET = 32

# 标题颜色（与调整窗口同一套配色）
H_FG = {1: "#1a3e6e", 2: "#1f5c99", 3: "#2e7d32", 4: "#b06a00"}


def find_default_spec() -> pathlib.Path | None:
    """在 reference/ 里找默认的《规范》docx（文件名带"规范"）。"""
    if not REFERENCE_DIR.is_dir():
        return None
    candidates = sorted(p for p in REFERENCE_DIR.glob("*规范*.docx")
                        if not p.name.startswith("~$"))
    return candidates[0] if candidates else None


def find_star10_spec() -> pathlib.Path | None:
    """在 reference/ 里找学术十星的格式要求 PDF（文件名带"学术十星"）。"""
    if not REFERENCE_DIR.is_dir():
        return None
    candidates = sorted(p for p in REFERENCE_DIR.glob("*学术十星*.pdf")
                        if not p.name.startswith("~$"))
    return candidates[0] if candidates else None


def build_spec_nodes(doc: Document, texts: list[str]) -> list[dict]:
    """把规范文档切成"节"清单：{level, text, start, end}。

    只由识别出的标题成节（封面/目录/说明这些"开头部分"不单独成节点——
    用户 2026-09-13：不要）；每节的 end＝下一个"级别不比自己深"的标题段号。
    """
    matches, _demoted = detect_headings(doc)
    nodes: list[dict] = []
    for i, m in enumerate(matches):
        raw = texts[i].strip()
        if not raw or _TOC_LINE.search(raw):
            continue                     # 目录页的条目不是真标题
        if m is not None:
            level = m.level
        else:
            # 引擎有"以标点结尾/过长就归正文"的过滤器（防论文误识别）——
            # 浏览规范时不设这种过滤：编号模式命中就算一节。
            hit = _try_numbered(raw)
            if hit is None:
                continue
            level = hit[0]
        nodes.append({"level": level, "text": raw, "start": i})

    return _finalize_nodes(nodes, len(texts))


def _finalize_nodes(nodes: list[dict], total: int) -> list[dict]:
    """算每节的 end（下一个"级别不比自己深"的标题），并给"什么都没认出来"兜底。"""
    if not nodes:
        nodes = [{"level": 0, "text": "（全文：未识别出标题结构）", "start": 0}]
    for k, nd in enumerate(nodes):
        end = total
        for nxt in nodes[k + 1:]:
            if nd["level"] == 0 or nxt["level"] <= nd["level"]:
                end = nxt["start"]
                break
        nd["end"] = end
    return nodes


def build_nodes_from_texts(texts: list[str]) -> list[dict]:
    """给"只有段落文字"的稿件切节（十星内置文本用）。

    与 `build_spec_nodes` 同构，但没有 docx 可查——只按编号模式认标题；
    "附录一：…""附录二：…"这类整篇大块（不是编号写法）另有一条专门规则。
    """
    nodes: list[dict] = []
    for i, raw in enumerate(texts):
        raw = raw.strip()
        if not raw:
            continue
        if _APPENDIX_HEAD.match(raw):
            level = 1
        else:
            hit = _try_numbered(raw)
            if hit is None:
                continue
            level = hit[0]
        nodes.append({"level": level, "text": raw, "start": i})
    return _finalize_nodes(nodes, len(texts))


@dataclasses.dataclass(frozen=True)
class SpecSource:
    """浏览窗口的数据源：文本、结构、顶栏文字、另存原件的信息。"""

    window_title: str
    bar_title: str
    name: str
    texts: list[str]
    nodes: list[dict]
    save_path: pathlib.Path | None
    save_label: str
    save_dialog_title: str
    save_filetypes: list[tuple[str, str]]
    bottom_hint: str


def _docx_source(path: pathlib.Path) -> SpecSource:
    doc = Document(str(path))
    texts = [p.text for p in doc.paragraphs]
    return SpecSource(
        window_title=f"《规范》原文浏览 —— {path.name}",
        bar_title="《规范》原文浏览",
        name=path.name,
        texts=texts,
        nodes=build_spec_nodes(doc, texts),
        save_path=path,
        save_label="另存原版 .docx …",
        save_dialog_title="另存《规范》原版（docx）",
        save_filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
        # 说明性小字不显示（2026-09-19 用户要求删除）
        bottom_hint="",
    )


def _star10_source() -> SpecSource | None:
    """十星格式要求的内置文本；`star10_spec_text.py` 不在时返回 None。"""
    try:
        from essay_tool.gui.star10_spec_text import STAR10_SPEC_TEXTS
    except Exception:
        return None
    texts = list(STAR10_SPEC_TEXTS)
    pdf = find_star10_spec()
    name = pdf.name if pdf is not None else "（内置提取文本）"
    return SpecSource(
        window_title=f"学术十星格式要求原文浏览 —— {name}",
        bar_title="学术十星格式要求原文浏览",
        name=name,
        texts=texts,
        nodes=build_nodes_from_texts(texts),
        save_path=pdf,
        save_label="另存原版 .pdf …",
        save_dialog_title="另存《学术十星格式要求》原版（pdf）",
        save_filetypes=[("PDF 文档", "*.pdf"), ("所有文件", "*.*")],
        # 说明性小字不显示（2026-09-19 用户要求删除）：左树右文的用法一眼就懂
        bottom_hint="",
    )


class SpecWindow(tk.Toplevel):
    """《规范》原文浏览窗口：左树右文，只整理结构、不动原文。"""

    def __init__(self, master: tk.Misc, source: SpecSource) -> None:
        super().__init__(master)
        self._source = source
        self.title(source.window_title)
        self.configure(bg=C_APP)
        try:
            self.state("zoomed")
        except Exception:
            self.geometry("1280x840")
            # 不能最大化时也要摆在主窗口正中央（用户 2026-09-15）。
            center_on(self, master, 1280, 840)
        self.minsize(960, 600)

        # ---- 数据源给的文本与节清单 ----
        self.texts = source.texts
        self.nodes = source.nodes
        self.heading_lv = {nd["start"]: nd["level"] for nd in self.nodes if nd["level"] > 0}

        # ---- 顶栏 ----
        top = tk.Frame(self, bg=C_BAR)
        top.pack(side="top", fill="x")
        # 顶栏只留粗标题和另存按钮——文件名在窗口标题里，别在顶栏重复
        # （2026-09-19 用户要求删掉"…（只整理结构，原文一字未动）"那行小字）。
        tk.Label(top, text=source.bar_title, font=(FAMILY, 12, "bold"),
                 bg=C_BAR, fg="#333").pack(side="left", padx=(12, 6), pady=6)
        save_btn = tk.Button(top, text=source.save_label, font=(FAMILY, 10),
                             command=self.on_save_copy)
        save_btn.pack(side="right", padx=(6, 12), pady=5)
        if source.save_path is None or not source.save_path.exists():
            save_btn.config(state="disabled")

        # ---- 主体：左树右文 ----
        paned = tk.PanedWindow(self, orient="horizontal", sashwidth=6, bg=C_APP,
                               sashrelief="flat")
        paned.pack(side="top", fill="both", expand=True)

        left = tk.Frame(paned, bg="#ffffff")
        style = ttk.Style(self)
        style.configure("Spec.Treeview", font=(FAMILY, 10), rowheight=32,
                        background="#ffffff", fieldbackground="#ffffff")
        self.tree = ttk.Treeview(left, show="tree", style="Spec.Treeview",
                                 selectmode="browse")
        vsb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(left, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)
        paned.add(left, width=440, minsize=260)

        right = tk.Frame(paned, bg="#ffffff")
        self.text = tk.Text(right, wrap="word", state="disabled", bg="#ffffff",
                            relief="flat", padx=8, pady=4, font=(FAMILY, 10))
        tvs = ttk.Scrollbar(right, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=tvs.set)
        self.text.pack(side="left", fill="both", expand=True)
        tvs.pack(side="right", fill="y")
        paned.add(right, minsize=400)

        # 正文区样式（显示用；文字本身一字未改）。字体对象同时拿来做折行测量。
        self._fonts = {
            "body": tkfont.Font(family=FAMILY, size=10),
            "h1": tkfont.Font(family=FAMILY, size=13, weight="bold"),
            "h2": tkfont.Font(family=FAMILY, size=12, weight="bold"),
            "h3": tkfont.Font(family=FAMILY, size=11, weight="bold"),
            "h4": tkfont.Font(family=FAMILY, size=11, weight="bold"),
        }
        self._cw: dict[str, dict[str, float]] = {}     # 单字宽度缓存（按样式分）
        self._current_node: int | None = None
        self._text_job: str | None = None
        self.text.bind("<Configure>", self._on_text_resize)

        self.text.tag_configure("h1", font=self._fonts["h1"], foreground=H_FG[1],
                                spacing1=10, spacing3=8)
        self.text.tag_configure("h2", font=self._fonts["h2"], foreground=H_FG[2],
                                spacing1=9, spacing3=6)
        self.text.tag_configure("h3", font=self._fonts["h3"], foreground=H_FG[3],
                                spacing1=8, spacing3=5)
        self.text.tag_configure("h4", font=self._fonts["h4"], foreground=H_FG[4],
                                spacing1=8, spacing3=5)
        self.text.tag_configure("body", font=self._fonts["body"], spacing3=7,
                                lmargin1=8, lmargin2=8)

        self.tree.tag_configure("lv0", foreground="#767676")
        self.tree.tag_configure("lv1", foreground=H_FG[1], font=(FAMILY, 10, "bold"))
        self.tree.tag_configure("lv2", foreground=H_FG[2])
        self.tree.tag_configure("lv3", foreground=H_FG[3])
        self.tree.tag_configure("lv4", foreground=H_FG[4])

        # 溢出控制：标题超过面板宽度就截断加"…"（全文在右侧与底栏可见）
        self.font_tree = tkfont.Font(family=FAMILY, size=10)
        self.font_tree_bold = tkfont.Font(family=FAMILY, size=10, weight="bold")
        self._label_job: str | None = None
        self.tree.bind("<Configure>", self._on_tree_resize)

        self._fill_tree()
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        # ---- 底栏（数据源没给提示语就整条不建）----
        if source.bottom_hint:
            bottom = tk.Frame(self, bg=C_BAR)
            bottom.pack(side="bottom", fill="x")
            tk.Label(bottom, text=source.bottom_hint,
                     font=(FAMILY, 9), bg=C_BAR, fg="#8a8a8a", anchor="w").pack(
                side="left", padx=10, pady=4)

        roots = self.tree.get_children()
        if roots:
            self.tree.selection_set(roots[0])
            self.tree.focus(roots[0])
            self._show_node(int(roots[0]))
        # 句柄必须登记进 `_label_job`（同 logic_viewer 的教训：句柄没人管，
        # 窗口在这 80ms 内被关掉，回调就会对已销毁的树操作）。
        self._label_job = self.after(80, self._refresh_tree_labels)

    # ------------------------------------------------------------ 结构树

    def _fill_tree(self) -> None:
        parents: dict[int, str] = {}   # level -> 最近一个该级节点的 iid
        for k, nd in enumerate(self.nodes):
            lv = nd["level"]
            if lv == 0:
                # 兜底节点（未识别出结构时）作为独立根
                self.tree.insert("", "end", iid=str(k), text=nd["text"],
                                 tags=("lv0",), open=True)
                continue
            parent = ""
            for upper in range(lv - 1, 0, -1):
                if upper in parents:
                    parent = parents[upper]
                    break
            self.tree.insert(parent, "end", iid=str(k), text=nd["text"],
                             tags=(f"lv{lv}",), open=True)
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
        t.configure(state="normal")
        t.delete("1.0", "end")
        for i in range(nd["start"], nd["end"]):
            raw = self.texts[i].strip()
            if not raw:
                continue
            lv = self.heading_lv.get(i)
            # 只配了 h1~h4 四个 tag（见上面的 tag_configure）；识别到的层级比
            # 四个深时夹到 4，不然 `self._fonts[tag]` 直接 KeyError（审计
            # 2026-09-15 核实：logic_viewer 里早就夹了，这里漏了）。
            tag = f"h{min(lv, 4)}" if lv else "body"
            for line in self._wrap(raw, avail, tag):
                t.insert("end", line + "\n", tag)
        t.configure(state="disabled")
        t.yview_moveto(0)

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
        - 英文单词尽量不从中间断开（优先退到最近的空格处）。
        """
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

    # ------------------------------------------------------------ 导出

    def on_save_copy(self) -> None:
        """把原版文件原样另存到用户指定的位置。"""
        src = self._source.save_path
        if src is None or not src.exists():
            messagebox.showerror("另存失败", "找不到原始文件。", parent=self)
            return
        try:
            verify_official(src)
        except IntegrityError as exc:
            messagebox.showerror("另存失败", str(exc), parent=self)
            return
        target = filedialog.asksaveasfilename(
            parent=self, title=self._source.save_dialog_title,
            initialfile=src.name, defaultextension=src.suffix,
            filetypes=self._source.save_filetypes)
        if not target:
            return
        try:
            shutil.copyfile(src, target)
        except OSError as exc:
            messagebox.showerror("另存失败", str(exc), parent=self)


def open_spec_viewer(master: tk.Misc,
                     spec_path: pathlib.Path | None = None) -> SpecWindow | None:
    """主 GUI 的"打开规范文件"按钮调这个：弹出（或唤醒）规范浏览窗口。"""
    win = getattr(master, "_spec_viewer_win", None)
    if win is not None and win.winfo_exists():
        win.deiconify()
        win.lift()
        win.focus_set()
        return win
    path = spec_path or find_default_spec()
    if path is None or not path.exists():
        messagebox.showerror("找不到《规范》", OFFICIAL_FILE_PROBLEM, parent=master)
        return None
    try:
        verify_official(path)
    except IntegrityError as exc:
        messagebox.showerror("《规范》文件异常", str(exc), parent=master)
        return None
    win = SpecWindow(master, _docx_source(path))
    setattr(master, "_spec_viewer_win", win)
    return win


def open_star10_viewer(master: tk.Misc) -> "SpecWindow | None":
    """主 GUI 的十星「格式要求」按钮调这个：弹出（或唤醒）内置原文浏览窗口。"""
    win = getattr(master, "_star10_viewer_win", None)
    if win is not None and win.winfo_exists():
        win.deiconify()
        win.lift()
        win.focus_set()
        return win
    source = _star10_source()
    if source is None:
        messagebox.showerror("找不到格式要求",
                             "内置的格式要求文本缺失（star10_spec_text.py）。",
                             parent=master)
        return None
    win = SpecWindow(master, source)
    setattr(master, "_star10_viewer_win", win)
    return win


# ================================================================ 入口


def do_check(path: pathlib.Path) -> None:
    doc = Document(str(path))
    texts = [p.text for p in doc.paragraphs]
    nodes = build_spec_nodes(doc, texts)
    print(f"《{path.name}》共 {len(texts)} 段，切出 {len(nodes)} 节：")
    for nd in nodes:
        indent = "  " * nd["level"]
        print(f"{indent}[L{nd['level']}] {nd['text']}  (第 {nd['start']}–"
              f"{nd['end'] - 1} 段)")


def do_check_star10() -> None:
    source = _star10_source()
    if source is None:
        print("内置的十星文本缺失（essay_tool/gui/star10_spec_text.py）。")
        return
    print(f"〈{source.name}〉共 {len(source.texts)} 段，切出 {len(source.nodes)} 节：")
    for nd in source.nodes:
        indent = "  " * nd["level"]
        print(f"{indent}[L{nd['level']}] {nd['text']}  (第 {nd['start']}–"
              f"{nd['end'] - 1} 段)")


def main() -> None:
    parser = argparse.ArgumentParser(description="《规范》原文浏览窗口")
    parser.add_argument("--spec", type=pathlib.Path, default=None,
                        help="规范 .docx 路径（默认在 reference/ 里找）")
    parser.add_argument("--star10", action="store_true",
                        help="浏览学术十星格式要求（内置文本，不依赖 reference）")
    parser.add_argument("--check", action="store_true", help="只打印结构，不开窗")
    args = parser.parse_args()

    if args.star10:
        if args.check:
            do_check_star10()
            return
        source = _star10_source()
        if source is None:
            print("内置的十星文本缺失（essay_tool/gui/star10_spec_text.py）。")
            return
    else:
        spec = args.spec or find_default_spec()
        if spec is None or not spec.exists():
            print("找不到规范文件。用 --spec 指定路径，或把规范放进 reference/。")
            return
        if args.check:
            do_check(spec)
            return
        source = _docx_source(spec)

    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = tk.Tk()
    root.withdraw()                      # 只留 SpecWindow 这个子窗口
    win = SpecWindow(root, source)
    win.protocol("WM_DELETE_WINDOW", root.destroy)
    print("窗口已打开。", flush=True)
    root.mainloop()


if __name__ == "__main__":
    main()
