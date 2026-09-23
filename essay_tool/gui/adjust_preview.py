"""识别结果调整窗口（2026-09-13 做界面形状；2026-09-14 接上真正的导出；
2026-09-19 起 `build_rows()` 直接消费识别层（`recognize.recognize()`）产出
的索引，不再自己重新识别一遍——调整结果通过 `doc_index.build_ops(rows,
index, ...)` 换算成对同一份 `index` 的身份更新 + 一份指令表，
`pipeline.run(index=..., ops=...)` 拿它们整体取代自动识别的结果）。

设计要点（与用户对齐的"身份地图"模型）：
- 一张表分三区：正文以前 / 正文 / 正文以后，每个区只有自己用得上的列：
    * 前/后区：结构标签 ｜ 正文
    * 正文区：一级标题 ｜ 二级标题 ｜ 三级标题 ｜ 四级标题 ｜ 正文
- 每一行 = 一个段落；**文字住在哪个格子，它就是什么身份**；
- 横向拖 = 换列（改身份）；拖到虚线空槽 = 认领那个位置；
- 校验（2026-09-19 起去掉"顺序自洽"这条硬约束——用户原话"允许用户随意
  调换两段之间的位置……既可以拖动一段文字和相邻的调换位置，也可以和远隔
  重洋的另一端调换位置"，段号顺序不再是拖放合法性的判据）：
    * 目标格必须空——占用了就拒收，没有"顶替"这回事：想换人，先在别处
      插一个空位把占用者挪走，腾出来再放（右键任意一行＝插入空位，不再
      局限于摘要/参考文献这几个原本就有"＋行"的块，正文区也能插）；
    * 结构未启用拒收（左侧勾选=启用，没勾整块灰）；
    * 支持多选（Ctrl/Shift 点选）一起拖——落位时依次对应到目标格开始
      往后数的若干个空位，选中行之间的先后顺序必须保持不变；
- 正文行的结构格 = ☐引用(长引文) ☐图 ☐表 三个小勾选（互斥）；**题注行勾了
  图/表之后还会多出一个 ☐续**（2026-09-15 晚）——续图/续表：仍算题注，
  但不挂样式 ⇒ 不进清单（**不涉及编号**：2026-09-16 起工具不写任何域）；
  用户 2026-09-17 统一了口径：**只要勾了图/表就画这个框**（英文对照行也画），
  它就是个"叠加开关"——勾上＝续，取消＝回到识别时的角色（`Row.caption_role_base`）；
- 引言/结论不"实体化"：结构页里画成正文里的一级标题，身份来自识别索引
  （`Entry.part=="引言"/"结论"`）；
- 文字之前的六圆点＝拖动抓手（点文字任意位置都能拖）；空位提示词只有三种（见 `_slot_hint` / `_slot_note`）；
- 不直接改原文档的段落 XML——GUI 编辑期间只重排 `self.rows` 这个 Python
  列表，真正的段落搬运在导出时才由 `doc_index.build_ops()`/`rebuild_body()`
  一次性做完；"这一段现在算哪个部件"这件事本身是**现场按当前排在哪个标签
  后面算的**（见 `doc_index.compute_span_ownership()`），不是一个存在
  `Row.block` 上、需要维护的持久标签——标签一挪，跟着它的内容归属立刻跟
  着变（2026-09-19 用户拍板："块"没有比"避免识别结果错乱"更深的含义）。

用法：
    .venv\\Scripts\\python.exe essay_tool\\gui\\adjust_preview.py            # 打开窗口
    .venv\\Scripts\\python.exe essay_tool\\gui\\adjust_preview.py --check    # 只打印模型，不开窗

已接上：调整后的身份地图会在导出时真正生效（标题级别、前置标签、图表题注、
长引文——长引文套预设里的 S18 样式，见 apply_structure.apply_long_quotes；
这个没有对应的自动探测器，只有手动勾选这一条路径）。
窗口关闭后调整不会丢：由调用方（主窗口）在重新打开时传回上一次的 `rows`
**和** `index`（两者配对，缺一个都不行），见 `AdjustWindow.__init__` 的
`initial_rows`/`initial_index` 参数。
"""

from __future__ import annotations

import argparse
import copy
import pathlib
import sys
import tkinter as tk
import tkinter.font as tkfont
from dataclasses import dataclass, replace
from tkinter import messagebox

ROOT = pathlib.Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from docx import Document  # noqa: E402

from essay_tool.engine.cleanup import unwrap_section_break_sdt  # noqa: E402
from essay_tool.engine.captions import CaptionMatch, captions_from_index  # noqa: E402
from essay_tool.engine.detect import (  # noqa: E402
    HeadingMatch,
    detect_appendix_entries_from_index,
)
from essay_tool.engine.doc_index import (  # noqa: E402
    BACK_ORDER,
    FRONT_ORDER,
    Entry,
    body_toc_field_parts,
    compute_span_ownership,
    has_toc_field,
    is_toc_paragraph,
    paragraph_ordinals,
    toc_field_part,
    toc_field_span_indices,
)
from essay_tool.engine.recognize import LEVEL_PART_NAMES, grammar_for, recognize  # noqa: E402


def _list_shape_conflict(key: str, enabling: bool) -> str | None:
    """三张清单的启用组合是否合法；不合法时返回一句说明，合法返回 None。

    用户 2026-09-15 定的四种合法形态：
    `插图清单`／`附表清单`／`插图和附表清单`／`插图清单 + 附表清单`。
    也就是说**合并形态与任一独立清单互斥**——两者同时开的话，同一张图会在
    成品里列两遍。引擎那边不再报"清单形态冲突"，改成在这里直接拦住。
    """
    if not enabling or key not in _LIST_SHAPE_PARTS:
        return None
    merged = "插图和附表清单"
    # 两个方向（开合并时已有独立的／开独立时已有合并的）说同一句话（用户 2026-09-20）
    conflict = "插图和附表清单不能和插图清单或附表清单一同启用！"
    if key == merged and (ENABLED.get("插图清单") or ENABLED.get("附表清单")):
        return conflict
    if key != merged and ENABLED.get(merged):
        return conflict
    return None


def _toc_field_row_text(paragraph) -> str:
    """目录/清单域在识别界面里显示成什么字（用户 2026-09-15）。

    这些域是 Word 自己生成的，不是用户写的正文，所以不把里面的条目铺成
    "正文"行，而是**一整行**标成它是什么。行文字同样进 `row.text`，只是
    字形和颜色不同（见 `_draw_cell_content`），而且不给拖。
    """
    return "word自动" + toc_field_part(paragraph)

DEFAULT_DOC = ROOT / "reference" / "民总.docx"

# 部件启用状态。
# 用户 2026-09-14 定"不套用任何预设时的默认值"：只有题目/副标题/摘要/目录
# （＋不在本表里的引言/结论）默认开，其余一律默认关；实际建表时再由
# `seed_enabled_from` 按"预设 ∪ 识别到的部件"重算，所以文档里真有的部件
# 不会因为没有默认值而被丢掉。
ENABLED: dict[str, bool] = {
    "题目": True, "副标题": True, "摘要": True, "关键词": False,
    "英文题目": False, "英文副标题": False, "英文摘要": False, "英文关键词": False,
    "目录": True, "插图清单": False, "附表清单": False, "插图和附表清单": False,
    "参考文献": False, "附录": False, "致谢": False,
}

BLOCK_NAME = {
    "题目": "题　目", "副标题": "副标题", "摘要": "摘　要", "关键词": "关键词",
    "英文题目": "英文题目", "英文副标题": "英文副标题", "英文摘要": "英文摘要",
    "英文关键词": "英文关键词",
    "目录": "目　录",
    "插图清单": "插图清单", "附表清单": "附表清单", "插图和附表清单": "插图和附表清单",
    "参考文献": "参考文献", "附录": "附　录", "致谢": "致　谢",
}


def detected_parts(rows: list) -> set[str]:
    """rows 里哪些结构块真的从文档里认出了内容——不是虚线空位，是真段落。"""
    return {r.block for r in rows if r.block and r.kind != "slot"}


def seed_enabled_from(preset_structure: dict | None, rows: list) -> None:
    """定"部件启用"的初始值 = 预设设定 ∪ 实际识别到的部件（用户 2026-09-14
    定的规则）：预设没决定、这份文档里也没识别到的部件，一律按未启用处理，
    不再用写死的默认值兜底——识别到的东西不会因为预设的默认假设被静默
    丢弃，但也不会平白无故多出预设压根没预期、文档里也没有的部件。

    只应该在**对着一份新文档建表时**调用（`main_window._rebuild_struct_adjust()`
    每次重建都是这种情况）；调用之后 `ENABLED` 完全交给用户自己改，不会被这个
    函数覆盖第二次。已经在调整中的文档又换了个预设，走的是下面的
    `merge_enabled_from`，不是这个。
    """
    baseline = dict(preset_structure or {})
    found = detected_parts(rows)
    ENABLED.clear()
    for part in BLOCK_NAME:
        ENABLED[part] = bool(baseline.get(part, False)) or part in found


def merge_enabled_from(preset_structure: dict | None) -> None:
    """应用预设时，把预设的部件要求"追加"进当前的启用状态——只会把还没
    启用的部件打开，不会关掉用户已经启用的（不管是识别自动认出来的还是
    手动点的）。

    跟 `seed_enabled_from` 的区别：那个是"重新定初始值"，会清空重来，只在
    换文档的时候用；这个是"在已有基础上再合并一次"，专门给"文档已经在
    调整中了，这时候又换了个预设"这种场景——很多用户是先导入文档、再去
    选预设，这时候文档早就建好表了，预设的结构设定不该被无视，但也不该
    把用户已经做的选择冲掉，所以只做"点亮"，不做"熄灭"。
    """
    for part, on in (preset_structure or {}).items():
        if on and part in ENABLED:
            ENABLED[part] = True


# 前置区/后置区里"整块可调换顺序"的部件，规范顺序（2026-09-18，配合"检测
# 顺序颠倒/自动重排"功能新增）——直接从引擎那份顺序知识
# （`doc_index.FRONT_ORDER`/`BACK_ORDER`）派生，去掉"引言"/"结论"：结构页
# 里这两个仍然画在正文区的一级标题栏（`build_rows()` 的
# `anchor_heading_view` 那条注释），不是前置/后置区的"块"。
FRONT_ROLES = tuple(p for p in FRONT_ORDER if p != "引言")
BACK_ROLES = tuple(p for p in BACK_ORDER if p != "结论")


def order_by_actual_position(parts: tuple[str, ...], first_pos: dict[str, int]) -> list[str]:
    """前置区各部件在结构页里从上到下的排法：**按它们在原文里的实际顺序**
    （用户 2026-09-19："按照实际识别出来的顺序安排 GUI 表格"），不替用户
    悄悄排成规范顺序——要不要按规范重排由用户自己点「一键重排」。

    `parts`：按规范顺序排好的部件名；`first_pos`：识别到的部件 → 它在原文里
    的首段号（没识别到的部件不在里面）。
      * 识别到的部件：按首段号从小到大排；
      * 没识别到的部件（只有一个虚线空位）：按规范顺序，各自插到**规范顺序
        里排在它前面的那些部件中、目前位置最靠后的那一个**后面（不夹在
        乱序的两个已有部件中间）；前面一个都没有就放最前。这样空位总跟着
        它规范上的前辈走，不会被甩到表尾。
    """
    order = sorted((p for p in parts if p in first_pos), key=lambda p: first_pos[p])
    for i, part in enumerate(parts):
        if part in first_pos:
            continue
        earlier = [order.index(p) for p in parts[:i] if p in order]
        order.insert(max(earlier) + 1 if earlier else 0, part)
    return order
# 用户手敲的封面范围在结构页里的块名（2026-09-15）：那几段是用户自己排的封面，
# 工具不识别、不加工，也**不允许拖动**——单独成一个结构块，仅按正文显示。
COVER_BLOCK = "已有封面"
# "设置封面范围"弹窗的宽度与左右留白（用户 2026-09-15：整个窗口调大、间距
# 都调大一点，别显得臃肿局促）。宽度写在 `minsize` 上：只定下限、不限高，
# 高度仍随内容走。
DIALOG_MIN_W = 880
PAD = 26
# 列表里每行显示多少个字（超过就截断加省略号）——按上面那个宽度量出来的，
# 保证一行装得下、不会被右边框切掉半截。
DIALOG_ROW_CHARS = 44
# 引言/结论不实体化：结构页里仍然画成正文区的一级标题——用户 2026-09-13
# 拍板："我们真的有必要实体化引言和结论吗？只把它当正常一级标题"。2026-09-19
# 起身份来自索引（`Entry.part=="引言"/"结论"`），不是靠标题文字现场判断，
# 见 `build_rows()` 的 `anchor_heading_view`。
# "标签+正文"型结构：至少常驻 标签位 + 一个正文位（2026-09-13 用户要求）
CONTENT_STRUCTS = ("摘要", "关键词", "英文摘要", "英文关键词", "参考文献", "附录", "致谢")
# 不区分"结构标签/正文"的整行结构：文字直接占满整行（题目/副标题/英文题目）。
# 英文题目原来混在"英文摘要"块的正文里，不能单独开关、也没有自己的空位——
# 用户 2026-09-14："应该跟中文标题一样单加一个部件，可以打开关闭"，改成
# 独立部件，待遇跟题目/副标题完全对齐。
FULL_WIDTH = ("题目", "副标题", "英文题目", "英文副标题")
# FULL_WIDTH 部件认领空位时，身份该定成哪个 kind——不能像前/后区那样按
# "是不是标签位"猜，这几个部件压根没有标签栏，只有一种身份。
_FULL_WIDTH_KIND = {"题目": "title", "副标题": "subtitle", "英文题目": "entitle",
                    "英文副标题": "subtitle"}
# 空位提示词一共这几类（用户 2026-09-14 定稿）：
#   * 未启用 → "未启用"（两种模式一样）；空白模板模式 → "将生成占位字符"；
#   * 导入后、启用但空着 → "可拖放文字到此" + 换行 "留白将生成占位字符"；
#   * 下面这几个部件特殊：它们的内容是 **Word 域自动生成**的（目录域 / 清单域），
#     既不是等文字来填的坑、也不能靠拖文字来生成——只显示一行"将自动生成"，
#     拖放一律拒收（用户原话："识别不到也不允许拖放，只允许启用后自动生成。
#     不然很迷惑"）。
# （原来这里有一张 PLACEHOLDER_NOTE_PARTS 名单，作用类似，已在这次统一提示词
#   时删掉；引擎侧"缺什么补什么"的行为一直没变。）
_AUTO_GENERATED_PARTS = ("目录", "插图清单", "附表清单", "插图和附表清单")
# 三张清单部件：它们之间只能有四种合法组合（用户 2026-09-15 定），见
# `_list_shape_conflict`。
_LIST_SHAPE_PARTS = ("插图清单", "附表清单", "插图和附表清单")
from essay_tool.gui.widgets import PANEL_BG, checkbutton, init_checkbox_style, wheel_notch
from essay_tool.gui.win_place import center_on
from essay_tool.engine.normalize import heading_style_text, label_style_text  # noqa: E402
from essay_tool.engine.apply_structure import LIST_PART_KINDS, reference_display_text

# 六圆点"抓手"占的宽度——画在文字之前，文字从格左缘 pad + GRIP_ROOM 处开始
GRIP_ROOM = 16

# 配色
C_APP = "#f4f4f2"
ZONE_BAND = {"front": "#ccd9ea", "body": "#d3e0cb", "back": "#e8dac4"}
ZONE_HDR = {"front": "#e3eaf4", "body": "#e8efe4", "back": "#f2e9da"}
ZONE_TEXT = {"front": "正文以前", "body": "正　文", "back": "正文以后"}
# 区块右侧那句"两列：结构标签 ｜ 正文"／"五列：一级~四级标题 ｜ 正文"已删
# （用户 2026-09-17：“删除图2-4展示的提示”）。
C_SEP = "#d5d4cc"
C_ROW_LINE = "#c9c8c0"
C_STRUCT_BG = "#f5f2ea"
C_STRUCT_BG_OFF = "#e8e8e8"
C_STRUCT_LINE = "#ddd6c6"
C_NUM = "#8b8b8b"
C_LABEL = "#8a5a00"
C_TITLE = "#1a3e6e"
C_HEADING = {"l1": "#1f5c99", "l2": "#2e7d32", "l3": "#b06a00", "l4": "#7b1fa2"}
C_SLOT_LINE = "#b8b8ae"
C_SLOT_TEXT = "#8f8f85"
# 目录/清单域那一行（"word自动目录"）：灰蓝斜体，跟正文一眼区分开
C_FIELD = "#5b6b7c"
C_CHIP = {"图": "#2e7d32", "表": "#00695c", "引用": "#6a1b9a", "续": "#b35c00"}
C_GHOST_BG = "#fff3bf"
C_GHOST_LINE = "#d4a017"

W_NUM = 50
# W_STRUCT 和 COL_LAYOUTS 里"结构标签"列的 170 都是**最小宽**。实际宽度在
# AdjustWindow.__init__ 里按当前 DPI 的实测字体宽重算（self.w_struct /
# self.w_label）——写死的像素在高 DPI 屏上不够用：同一段文字 100% 缩放和
# 150% 缩放下像素宽能差一半，"插图和附表清单"100% 只要 91px、150% 要
# 137px，原来按 172 固定给、名字区只剩 94px，必截成"插图和…"，用户
# 永远看不全（2026-09-14 反馈，定性为纯 UI 问题）。
W_STRUCT = 172
# 多选复选框列（2026-09-19 新增，用户明确要求单开一列，放在"结构"表头和
# 区特定列（"结构标签"/正文……）之间，列头写"多选"）——固定宽度，装得下
# 一个 ☑/☐ 加一点留白，不用像 w_struct 那样按内容实测。
W_SELECT = 54

# 行内多行内容（提示词 / 说明 / "自动重新编号参考文献"复选框）的行距。
# 9 号字实测行高 17px，留 2px 余量；行高、居中排布、点击区全都用这一份数字。
_SWITCH_LINE_PITCH = 19

COL_LAYOUTS = {
    "front": [("label", "结构标签", 170), ("text", "正文", None)],
    "body": [("l1", "一级", 0), ("l2", "二级", 0), ("l3", "三级", 0),
             ("l4", "四级", 0), ("text", "正文", None)],
    "back": [("label", "结构标签", 170), ("text", "正文", None)],
}
# 一~四级标题这四列的表头以前写"一级标题"…四个四字全称，四列并排就是 336px
# （每列写死 84），把正文挤得没地方（用户 2026-09-14："那四个字太长了，挤满了"）。
# 现在缩成"一级"…两字，宽度值写 0 表示"由 AdjustWindow.w_head 按字号实测算"，
# 别在这儿填死数字。完整说法由区带注释"五列：一级~四级标题 ｜ 正文"和拖拽
# 提示（COL_NAME）负责，信息不丢。
_HEAD_KEYS = ("l1", "l2", "l3", "l4")
COL_NAME = {"label": "结构标签", "text": "正文", "l1": "一级标题", "l2": "二级标题",
            "l3": "三级标题", "l4": "四级标题"}

FAMILY = "Microsoft YaHei UI"


# ================================================================ 数据模型


@dataclass(eq=False)
class Row:
    rid: int = 0
    kind: str = "body"          # zone/title/subtitle/label/content/heading/body/caption/slot
    text: str = ""
    col: str = "text"           # label/text/l1..l4
    doc_index: int | None = None
    block: str | None = None
    special: str | None = None  # "引用"/"图"/"表"（引自或用勾选切换）
    slot_for: str | None = None
    slot_col: str = "label"
    zone: str = "body"
    block_first: bool = False
    block_last: bool = False
    added: bool = False        # 手加的空位（＋行）：带 ✕、可删；默认空位不许删
    regenerate: bool = False   # 已有真域的目录/清单行：用户勾了"重新为文档生成"
    # 目录/清单域行专用：折进这一行的**缓存条目段**（域起始段之后、域结束之前
    # 那些没有域代码的段落，民总的目录就是这样：只有第一行带域代码，其余 19 行
    # 是纯 Normal 文字）。导出时 Word 会刷新出真目录，这些行不是用户写的正文，
    # 所以不单独铺行；`doc_index.build_ops` 让它们跟着这一行一起走。
    attached: tuple[int, ...] = ()
    # 结构页"显示成导出时的写法"要用的原材料（用户 2026-09-15）：
    #   * `heading`＝这一段被识别成的标题对象（装着编号/标记/标题内容，
    #     `normalize.heading_style_text()` 就是靠它们算出规范写法的）；
    #   * `label`＝它作为标签时的部件名（"摘要"/"Abstract"/"引言"……）。
    # 两者都可能为 None（普通正文、封面里的段落、空位、题注）——那种行照旧
    # 显示原文。**存原材料而不是当场算成字符串**：用户随时能把这一行拖到
    # 别的级别上，级别一变写法就得跟着变（见 `shown_text()`）。
    heading: HeadingMatch | None = None
    label: str | None = None
    # 题注角色（`special` 是"图"/"表"时才有效）——**只决定挂不挂样式**，
    # 跟编号无关（2026-09-16 起工具不写任何域、不改任何编号）：
    #   "main"＝普通题注（挂图题/表题真样式 ⇒ 进清单；**分图也算这一档**，
    #           `图 1（a）…`、`图 1a…` 不再区分主分）
    #   "cont"＝续图/续表（**不挂样式** ⇒ 不进清单）——结构页那个
    #           “☑续”勾的就是它（用户 2026-09-15 晚：“可以让用户加『续』，
    #           结果是即使识别、设置为图题表题也不进清单”）。**种类不换**：
    #           “图+续”＝续图、“表+续”＝续表，直接格式跟着种类的标签走。
    #   "en"  ＝英文对照行（不挂样式、不进清单）
    # 识别层把题注写进 `Entry.meta["caption"]`（`captions.py`），这一行的
    # `special`/`caption_role` 由它翻译而来；导出时 `doc_index.build_ops()`
    # 把这里（含手动勾选/取消）写回同一个键，告警与套样式都读索引，
    # 不再现场重新识别（2026-09-19 接通）。
    caption_role: str | None = None
    # **识别时的角色**（`caption_role` 的“原值”），只用于一件事：取消"续"
    # 时回退到哪里。用户 2026-09-17：「一旦图表被勾选，就应该显示续……按理来
    # 讲这个就应该可用」⇒ "续"现在在**所有**勾了图/表的行上都画（含英文对照行），
    # 它成了一个纯粹的"叠加开关"：勾上＝cont，取消＝回到这一格里的原角色。
    # 没有它的话，英文对照行点两下"续"就会被推成普通题注（挂样式⇒清单多一条，
    # 2026-09-16 审计抓到过的坑）；分开存之后，“手动把图/表取消再勾上”仍然是
    # 唯一一条"我要把它当正题注"的明确路径。
    caption_role_base: str | None = None
    # 附录序号题名（`附录A …`，规范 [226]"另起一行……标注附录序号和题名"）：
    # 这类段落不算标题、不进目录，光看文字跟普通正文没区别，所以结构页里
    # 给它一个标记（用户 2026-09-16 定："放到标签那一栏，但靠右对齐"）。
    # **只影响显示**：套 S20 是引擎按 `detect.detect_appendix_entries` 做的
    # （`apply_structure.apply_appendix_entries`），跟这一页有没有标记无关。
    appendix_entry: bool = False
    # 这一段是 Word 自动编号列表的一项*时，它会被**写成手打文字的编号**长什么
    # 样（例：`（1） `，全角、带一个空格）。**只在识别那一刻算一次**：
    # 用户 2026-09-17：“要替用户写成规范写法！比如改为全角什么的。只是不要让它
    # 随着用户拖动层级更新层级，比如把（1）改成 1. 是不许的。”
    # 所以这一行被拖到别的标题层级、或拖回“正文”时，这里的文字一个字不变；
    # 导出也拿同一份（见 `build_manual_structure` 里的 `auto_prefix`/`auto_bodies`）。
    auto_number: str | None = None


def build_rows(path: pathlib.Path, cover_end: int | None = None,
               heading_scheme: str | None = None) -> tuple[list[Row], list[Entry]]:
    """读一份真实文档，把**识别索引**（`recognize.recognize()` 的产出）直接
    翻译成"身份地图"的行——2026-09-19 起不再自己重新判断"这段是什么"，
    那件事识别层已经做完了，这里只做 `Entry.part`/`Entry.meta` → `Row` 字段
    的搬运（用户原话："索引里已经加入了这一段是什么，按照索引来重建 GUI
    表格，为什么还要判断这段文字是什么？"）。

    返回 `(rows, index)`：`index` 是这次识别产出的 `list[doc_index.Entry]`，
    调用方要留着（`AdjustWindow.index`）——导出时 `doc_index.build_ops(rows,
    index, ...)` 要用同一份对象：它会把面板当前的行序换算成对这份 `index`
    的身份更新（原地改 `entry.part`/`.meta`），`pipeline.run()` 的
    `normalize_document`/`apply_from_index` 等各步再读这份被更新过的
    `index`——这是"调整结果真的能影响导出"这件事成立的唯一途径，`index`
    和 `rows` 必须配对着传，不能只留 `rows`。

    `cover_end`：用户手敲的封面范围（第 0 段到第 N 段，**含**）。那一片不参与
    识别（`recognize()` 自己会跳过），不然封面上的"课程名称：____"会被当成
    题目、盖过正文里真正的题目。被忽略的段落仍会铺成普通正文行（"已有封面"
    那个结构块），只是不再冒充任何部件。

    `heading_scheme`：原文的标题编号方案（"chinese"/"dotted"，见
    `recognize.grammar_for`），None 走默认的中文方案。

    ⚠️ 先 `unwrap_section_break_sdt`——跟 `pipeline.run()` 第一步一模一样，
    `keep_content` 也按同一条规矩算（源文件里有目录/清单域就拆包保留、
    没有就可以扔掉旧内容），两边不一致的话，这里存下的段号到了管线里就会
    对不上（`Entry.top_index` 是按元素在 body 里的位置算的）。
    """
    doc = Document(str(path))
    unwrap_section_break_sdt(doc, keep_content=bool(body_toc_field_parts(doc)))
    paras = doc.paragraphs

    index = recognize(doc, grammar=grammar_for(heading_scheme), cover_end=cover_end)
    ordinals = paragraph_ordinals(index)

    # 每个部件名 → 它在文档里的段号列表（按位置排序）：标签 + 识别层已经
    # 扫好的内容段落全在里面，不用自己再判一遍"到哪为止"。
    positions_of: dict[str, list[int]] = {}
    for i, entry in ordinals.items():
        if entry.part:
            positions_of.setdefault(entry.part, []).append(i)
    for lst in positions_of.values():
        lst.sort()

    def first_of(part: str) -> int | None:
        lst = positions_of.get(part)
        return lst[0] if lst else None

    caption_at = {c.paragraph_index: c for c in captions_from_index(index)}
    # 附录序号题名（`附录A …`）：不是标题、不进目录，只在结构页上标一行
    # 小字。判据用引擎那一份（`detect.detect_appendix_entries_from_index`），
    # 别在这儿再写一套。
    appendix_at = {i for i, _letter in detect_appendix_entries_from_index(doc, index)}
    # 目录/清单域落在哪个部件名下：域本身不是识别层的活（它不认段落文字，
    # 认的是 Word 域代码），得单独扫一遍——跟识别层是不是找到了对应的字面
    # 标签无关。
    field_part_at = {i: toc_field_part(p) for i, p in enumerate(paras) if has_toc_field(p)}
    # 域区间里的缓存条目段：折进域行，不单独铺（见 `Row.attached`）。按部件
    # 分别算区间，免得相邻的两张清单域被当成一个域。
    attached_of: dict[int, tuple[int, ...]] = {}
    for part_name in set(field_part_at.values()):
        span = sorted(toc_field_span_indices(paras, {part_name}))
        run: list[int] = []
        for k in [*span, None]:
            if run and (k is None or k != run[-1] + 1):
                heads = [m for m in run if field_part_at.get(m) == part_name]
                for head in heads:
                    nxt_heads = [m for m in heads if m > head]
                    stop = nxt_heads[0] if nxt_heads else run[-1] + 1
                    attached_of[head] = tuple(
                        m for m in run if head < m < stop and m not in field_part_at)
                run = []
            if k is not None:
                run.append(k)
    folded = {m for rest in attached_of.values() for m in rest}

    rows: list[Row] = []
    counter = [0]

    def add(**kw) -> Row:
        counter[0] += 1
        row = Row(rid=counter[0], **kw)
        rows.append(row)
        return row

    def txt(i: int) -> str:
        return " ".join(paras[i].text.split())

    def auto_number_of(i: int) -> str | None:
        """段落 i 是 Word 自动编号的标题时，它会被写成的那串手打编号文字
        （识别层在 `Entry.meta["canonical"]` 里已经算好，直接用，不重算）。
        用户 2026-09-17："不要让它随着用户拖动层级更新层级"——这串文字只在
        识别那一刻定一次，后面不管拖到哪一级都不变。
        """
        entry = ordinals.get(i)
        if entry is None or not entry.meta.get("auto_numbered"):
            return None
        return entry.meta.get("canonical")

    def heading_view(i: int) -> HeadingMatch:
        """数字编号标题的"原材料"对象——鸭子类型，字段名跟旧
        `detect.HeadingMatch` 一致，`normalize.heading_style_text()` 认的
        就是这几个属性。`canonical` 是识别层在识别那一刻按当时层级定死的
        规范写法（None ＝ 不规范化，原文原样保留），显示和导出读的是同一份，
        用户之后拖动层级也不变。
        """
        meta = ordinals[i].meta
        return HeadingMatch(
            i, meta["level"], meta.get("number"), meta.get("marker", ""),
            meta.get("title", ""), role=None,
            auto_numbered=bool(meta.get("auto_numbered")),
            canonical=meta.get("canonical"),
        )

    def anchor_heading_view(i: int, role: str) -> HeadingMatch:
        """引言/结论：结构页仍然画在正文区的一级标题栏里（纯 UI 布局延续，
        见模块顶部 `CONTENT_STRUCTS` 上面那条注释），但身份不是"数字编号
        标题"——是 `role` 本身。`role` 字段让 `heading_style_text()` 走"固定
        名称锚点"那条分支，显示成规范写法（引　言／结　论）；`doc_index.
        build_ops()` 看到 `row.heading.role` 也会直接拿它当 `entry.part`，
        不按列算"一级标题"（见该函数注释）。点分体系里带编号的（`1 引言`，
        `keep_text`）显示用户原文，不改写。
        """
        return HeadingMatch(i, 1, None, "", txt(i), role=role,
                            keep_text=bool(ordinals[i].meta.get("keep_text")))

    def add_body_row(j: int, *, kind: str, block: str | None, zone: str) -> None:
        """给段落 j 铺一行（前置内容/正文/正文以后共用）。

        目录/清单域特殊：**一整行**标成 "word自动目录"（它的条目是 Word 生成的，
        铺成正文行会让人以为那是用户写的字），并且归到它自己的部件名下；
        域生成的条目（`toc N` 样式）直接不铺。

        题注也要在这里认出来：前后置部件区间里的图题/表题要带上
        `special`/`caption_role`，否则打开过结构页再导出，那些题注会丢掉
        "图题/表题"身份——不再套图题/表题样式、不进清单。
        """
        paragraph = paras[j]
        part = field_part_at.get(j)
        if part is not None:
            add(kind="toc_field", text=_toc_field_row_text(paragraph), col="text",
                doc_index=j, block=part if part in _AUTO_GENERATED_PARTS else block,
                zone=zone, attached=attached_of.get(j, ()))
            return
        if is_toc_paragraph(paragraph) or j in folded:
            return
        if not paragraph.text.strip():
            return
        cap = caption_at.get(j)
        add(kind="caption" if cap is not None else kind, text=txt(j), col="text",
            doc_index=j, block=block, zone=zone,
            special=(cap.kind if cap is not None else None),
            caption_role=(_caption_role(cap) if cap is not None else None),
            caption_role_base=(_caption_role(cap) if cap is not None else None),
            appendix_entry=j in appendix_at,
                      auto_number=(None if cap is not None else auto_number_of(j)))

    # ---------------------------------------------------------- 正文以前
    add(kind="zone", zone="front")
    # 用户手敲的封面范围：单独铺成一个结构块「已有封面」。里面的东西**不允许
    # 拖动**（见 `_draw_cell_content`），仅按正文显示；导出时也不作任何加工
    # （识别/套样式/清格式/改图全跳过，见 `pipeline.run()`）。空段不铺（跟别处
    # 一致），行里的"段号"仍是它在原文里的真实段号。
    if cover_end is not None:
        add(kind="label", text=COVER_BLOCK, col="label", block=COVER_BLOCK, zone="front")
        for j in range(0, min(cover_end + 1, len(paras))):
            if paras[j].text.strip():
                add(kind="content", text=txt(j), col="text", doc_index=j,
                    block=COVER_BLOCK, zone="front")
    # 前置区各部件从这里开始，**按原文里的实际顺序**一个个铺（见文件顶部的
    # `order_by_actual_position`）；下面每个 `xxx_block` 只管铺自己那一块。

    def single_block(key: str) -> None:
        """题目/副标题/英文题目/英文副标题：整行一格，文字放"正文"列，身份由
        左侧结构块的勾选标识；缺失时有自己的空位。英文题目是独立部件，待遇
        跟中文题目/副标题对齐（用户 2026-09-14："应该跟中文标题一样单加一个
        部件，可以打开关闭"）。行 kind：英文副标题沿用 "subtitle"（渲染、
        拖放规则跟中文副标题一模一样），靠 `block` 区分是哪一个。"""
        idx = first_of(key)
        if idx is not None:
            add(kind=_FULL_WIDTH_KIND[key], text=txt(idx), col="text", doc_index=idx,
                block=key, zone="front")
        else:
            add(kind="slot", col="text", block=key, slot_for=key, slot_col="text", zone="front")

    def label_block(key: str) -> None:
        """"标签+正文"两段式部件（摘要/关键词/英文摘要/英文关键词/参考
        文献/附录/致谢）通用铺法——识别层的"标签+内容两段式"扫描已经把这
        个部件所有段落（标签 + 内容）算进了 `positions_of[key]`，这里只管
        照抄，不用再自己找"到哪为止"。目录/插图清单/附表清单/插图和附表
        清单不是"标签+内容"式扫描的对象（识别层不给它们扫内容），走各自的
        分支，不调用这个。
        """
        positions = positions_of.get(key, [])
        if not positions:
            add(kind="slot", col="label", block=key, slot_for=key, slot_col="label", zone="front")
            if key in CONTENT_STRUCTS:
                add(kind="slot", col="text", block=key, slot_for=key, slot_col="text",
                    zone="front")
            return
        label_idx = positions[0]
        add(kind="label", text=txt(label_idx), col="label", doc_index=label_idx, block=key,
            zone="front", label=key)
        for j in positions[1:]:
            add_body_row(j, kind="content", block=key, zone="front")

    existing_toc = bool(body_toc_field_parts(doc))

    def toc_block() -> None:
        """目录块：标签行后面**紧跟着一行"将生成Word自动目录"**（除非源文件里
        已经有真目录域——那种情况不在这里画任何内容行）。用户 2026-09-15：
        "我还是没在这个界面看到 Word 自动目录什么的。"

        ⚠️ 目录（以及三张清单）**不做"标签之后扫到下一个已识别部件"这种
        内容行扫描**（2026-09-19 修，`民总.docx` 实测炸出来的真事故）：
        已有的目录域往往是"只有第一段带真正域代码、其余条目段落纯 Normal
        样式"的多段渲染结果——`is_toc_paragraph`/`has_toc_field` 逐段查都
        认不出那些中间段落，扫描会把它们当成"目录的正文"一并收进来，
        `doc_index.build_ops()` 据此把 `entry.part` 全部改写成"目录"，
        这些段落本来是文档正文里的一级/二级标题，身份被整个吃掉，成品里
        变成一串重复的"【目录】"占位段。目录/清单的"内容"本来就该是 Word
        域自己生成的，不需要在这个表格里铺行——真正的域段落（`field_part_at`
        能认出来的那一段）会被下面"前置区里没有身份的段落"那个通用扫描
        顺带捞出来、标成 "word自动目录"，不用在这里另开一条路。
        """
        toc_label = first_of("目录")
        if toc_label is None:
            add(kind="slot", col="label", block="目录", slot_for="目录", slot_col="label",
                zone="front")
            return
        add(kind="label", text=txt(toc_label), col="label", doc_index=toc_label,
            block="目录", zone="front", label="目录")
        if not existing_toc:
            add(kind="toc_field", text="将生成Word自动目录", col="text",
                block="目录", zone="front")

    def list_block(key: str) -> None:
        """三张清单同样只画标签行（理由见 `toc_block`），分三种情况（用户
        2026-09-15）：
          * 源文件里已经有标签 → 标签行（清单域会被下面的通用扫描铺成
            "word自动插图清单"那一行）；
          * 只有真域、没有标签 → 那个域行会在下面"没有身份的段落"那一段
            被找到，不再画空位；
          * 两样都没有 → 画空位，等着工具在导出时生成。
        """
        if any(r.block == key and r.kind in ("label", "content", "toc_field") for r in rows):
            return
        idx = first_of(key)
        if idx is None:
            add(kind="slot", col="label", block=key, slot_for=key, slot_col="label", zone="front")
            return
        add(kind="label", text=txt(idx), col="label", doc_index=idx, block=key,
            zone="front", label=key)

    def emit_front_part(key: str) -> None:
        if key in _FULL_WIDTH_KIND:
            single_block(key)
        elif key == "目录":
            toc_block()
        elif key in _LIST_SHAPE_PARTS:
            list_block(key)
        else:
            label_block(key)

    first_positions = {p: first_of(p) for p in FRONT_ROLES if first_of(p) is not None}
    for part in order_by_actual_position(FRONT_ROLES, first_positions):
        emit_front_part(part)

    # -------------------------- 正文/正文以后从哪开始（下面几段都要用）
    # 数字编号标题 + 引言/结论：正文区的起点；参考文献/附录/致谢：正文以后
    # 区的起点。两组合起来才是"前置区结束的地方"——跟旧口径一致：正文区
    # 从第一个已识别标题（含引言/结论）开始，正文以后从第一个
    # 参考文献/附录/致谢开始，取更靠前的那个当前置区终点。
    body_like = sorted(i for i, e in ordinals.items()
                       if e.part in ("引言", "结论") or e.part in LEVEL_PART_NAMES.values())
    back_like = sorted(i for i, e in ordinals.items() if e.part in BACK_ROLES)
    front_end = min([*body_like, *back_like], default=len(paras))
    body_start = body_like[0] if body_like else None
    back_start = back_like[0] if back_like else len(paras)

    # --------------------------------- 前置区里"没有身份"的段落（用户 2026-09-15）
    # 前置区是**按部件**铺的（各部件的先后按原文实际顺序，见上面），没有
    # "题目之前"这个位置；正文区又是从**第一个已识别标题**才开始的。两头一夹，凡是
    # "落在前置区、又没被认成任何部件"的段落就**一行都不铺**——而导出时它们
    # 还在文档里（套正文格式），成品上就挂在题目上方或某个部件旁边，用户在
    # 这个界面里却看不见、也没处改（用户实测抓到的：题目上面一句话、没有
    # 【摘要】标签时的那段摘要文字，都显示不出来）。
    #
    # 现在把它们按**在原文里的真实位置**补铺成普通正文行：位置＝"最后一个
    # 文档下标比它小的已铺行"之后，所以题目之前的就排在题目行之前，两个
    # 部件之间的就落在中间——顺序跟源文件一致。已经在上面铺过的段落
    # （题目/副标题/英文题目/标签/域/正文位）不会重复铺。
    emitted = {r.doc_index for r in rows if r.doc_index is not None}
    # 插入位置不能越过区块抬头行（`zone`），否则正文行会跑到"前置区"那条
    # 横带上面去，看着像另一个区的东西。
    floor = 1 if rows and rows[0].kind == "zone" else 0
    for i in range(front_end):
        if i in emitted:
            continue
        paragraph = paras[i]
        part = field_part_at.get(i)
        if part is not None:
            # 目录/清单域：一整行说明它是什么（跟正文区的待遇一致），而不是
            # 把域里的条目铺成正文。没有标签的真域就是靠这一行被看见的。
            row = Row(rid=0, kind="toc_field", text=_toc_field_row_text(paragraph),
                      col="text", doc_index=i,
                      block=part if part in _AUTO_GENERATED_PARTS else None, zone="front",
                      attached=attached_of.get(i, ()))
        elif not paragraph.text.strip() or is_toc_paragraph(paragraph) or i in folded:
            continue
        else:
            cap = caption_at.get(i)
            row = Row(rid=0, kind="caption" if cap is not None else "body",
                      text=txt(i), col="text", doc_index=i, zone="front",
                      special=(cap.kind if cap is not None else None),
                      caption_role=(_caption_role(cap) if cap is not None else None),
                      caption_role_base=(_caption_role(cap) if cap is not None else None),
                      auto_number=(None if cap is not None else auto_number_of(i)))
        position = floor
        for k, existing in enumerate(rows):
            if existing.doc_index is not None and existing.doc_index < i:
                position = k + 1
        rows.insert(max(position, floor), row)

    # -------------------------------------------------------------- 正文
    add(kind="zone", zone="body")
    if body_start is not None:
        for i in range(body_start, back_start):
            part = field_part_at.get(i)
            if part is not None:
                add(kind="toc_field", text=_toc_field_row_text(paras[i]), col="text",
                    doc_index=i, block=part if part in _AUTO_GENERATED_PARTS else None,
                    zone="body", attached=attached_of.get(i, ()))
                continue
            if is_toc_paragraph(paras[i]) or not paras[i].text.strip() or i in folded:
                continue
            entry = ordinals.get(i)
            # 引言/结论是"标签+内容两段式"部件（识别层把它们的正文内容也
            # 扫成了 `part="引言"/"结论"`，`slot="content"`）——只有标签本身
            # 那一段才画进一级标题栏，内容段落仍按普通正文处理，否则整段
            # 引言正文会被误铺成一串"标题"。
            if (entry is not None and entry.part in ("引言", "结论")
                    and entry.meta.get("slot") != "content"):
                add(kind="heading", text=txt(i), col="l1", doc_index=i, zone="body",
                    heading=anchor_heading_view(i, entry.part),
                    auto_number=auto_number_of(i))
                continue
            if entry is not None and entry.part in LEVEL_PART_NAMES.values():
                level = entry.meta["level"]
                add(kind="heading", text=txt(i), col=f"l{level}", doc_index=i,
                    zone="body", heading=heading_view(i), auto_number=auto_number_of(i))
                continue
            add_body_row(i, kind="body", block=None, zone="body")

    # ---------------------------------------------------------- 正文以后
    add(kind="zone", zone="back")
    back_start_row = len(rows)      # 区带行之后的位置：末尾只重排这一段
    for key in BACK_ROLES:
        label_idx = first_of(key)
        if label_idx is None:
            add(kind="slot", col="label", block=key, slot_for=key, slot_col="label", zone="back")
            add(kind="slot", col="text", block=key, slot_for=key, slot_col="text", zone="back")
            continue
        add(kind="label", text=txt(label_idx), col="label", doc_index=label_idx, block=key,
            zone="back", label=key)
        for j in positions_of[key][1:]:
            add_body_row(j, kind="content", block=key, zone="back")

    # 后置区这些行**按文档顺序重排一遍**（没有对应段落的空位排在最后）：
    # 规范顺序 ≠ 文档顺序时（比如"附录"写在"致谢"后面），按角色顺序硬排会让
    # 表格顺序跟原文对不上——初始铺表就该如实显示识别到的真实顺序，要不要
    # 按规范重排交给用户自己点（2026-09-19 用户口径："按照实际的识别出来的
    # 顺序安排 GUI 表格……允许用户在调整各块归属以后手动拖动部件名称调整
    # 部件顺序，也支持按照预设一键重排"）。
    rows[back_start_row:] = sorted(
        rows[back_start_row:],
        key=lambda r: (r.doc_index is None, r.doc_index or 0))

    ensure_min_content_rows(rows)
    for i, row in enumerate(rows):
        row.rid = i + 1
    annotate_blocks(rows)
    return rows, index


def _caption_role(caption: CaptionMatch) -> str:
    """题注的角色字串（存到 `Row.caption_role` 上，跟引擎的标志一一对应）。

    只有两种：**不挂样式的两种**（英文对照行 `en`、续行 `cont`）与其余（`main`）。
    2026-09-16 去掉了 `sub`（分图）——“只要能认出来是题注，区分主分是没有
    意义的”，分图跟普通题注走完全一样的路（挂样式、进清单）。
    """
    if caption.is_english:
        return "en"
    if caption.is_continuation:
        return "cont"
    return "main"


def _caption_chip(row: Row) -> str | None:
    r"""题注行在「结构」列里的小标签：图题／表题；标了“续”写「图题（续）」、
    英文对照行写「图题（英）」。

    “续”是**角色**、不是种类（用户 2026-09-15 晚）：勾了“图”+续＝续图、
    “表”+续＝续表——仍然算题注，但**不挂样式** ⇒ **不进清单**（清单域是按
    样式名收条目的）。标签里带上“（续）”，一眼能看出它不是普通那条。

    英文对照行（`en`）**必须带标记**：它也**不挂样式 ⇒ 不进清单**，但面板
    上光看文字跟普通题注没区别，用户根本看不出这行为什么没进清单（审计
    2026-09-16 报的）。顺便把“不挂样式”的两种情形用两种标签分开。
    """
    if row.kind != "caption" or row.special not in ("图", "表"):
        return None
    chip = "图题" if row.special == "图" else "表题"
    role = row.caption_role or "main"
    if role == "cont":
        return chip + "（续）"
    if role == "en":
        return chip + "（英）"
    return chip


def shown_text(row: Row, insert_text: dict[str, str] | None) -> str:
    """这一行在结构页里**显示什么文字**＝它导出时会写成什么样。

    用户 2026-09-15："在结构界面就显示识别后的更正。例如，通过 1、识别进入
    标题的，直接显示为导出时的 1. ，该改全角的半角改为全角，该加【】的加【】"。
    算规则全部复用 `engine/normalize.py` 的那两个函数（显示和导出不会各说
    各话）。

    ⚠️ **编号写法在识别那一刻按当时的层级定死，拖动层级绝不重编**（用户
    2026-09-19："拖动绝对不能编号"，翻转了 2026-09-15"拖到别的级别就按新
    级别重编"的老口径；2026-09-17 对 Word 自动编号早就是这个口径：
      「要替用户写成规范写法！比如改为全角什么的。只是不要让它随着用户拖动
        层级更新层级，比如把（1）改成 1. 是不许的。」）
    现在手打编号也一样：显示用的 `row.heading.canonical`、导出用的
    `entry.meta["canonical"]`（`doc_index.build_ops()` 从前者写进后者）是
    同一串，跟这一行现在挂在哪一列无关。Word 自动编号的项编号不在文字里，
    `row.auto_number` 那一支放在最前面，直接把它接上去。

    题注行**显示原文**（2026-09-16 起我们不碰编号了，所以“导出时会写成什么样”
    就是原文：以前那套“把编号换成推算出来的号”的显示（`caption_style_text`）
    随题注化一起删掉了）。

    没有原材料的行（普通正文/封面里的段落/空位/题注）或没拿到预设的“插入
    文字”时，一律显示原文——宁可少显示，也不展示一份猜出来的写法。
    """
    if not row.text:
        return row.text
    # Word 自动编号的项：编号原样跟上（跟预设无关，所以放在"没拿到预设"那条
    # 分支之前）。标题行和拖回正文的行走同一套——不然拖一下编号就变了样。
    if row.auto_number:
        return f"{row.auto_number}{row.text}"
    if insert_text is None:
        return row.text
    if row.heading is not None:
        return heading_style_text(row.heading, insert_text, raw=row.text)
    if row.label is not None:
        return label_style_text(row.label, insert_text, raw=row.text)
    return row.text


def build_blank_rows() -> list[Row]:
    """还没导入文件时，"论文结构"页签要显示的骨架——每个前/后置部件一个
    空位，只用来给用户先勾选"生成空白模板时要哪些部件"（用户 2026-09-14：
    "导入文件前就打开结构页面的，在表格里显示空白模板的信息……禁用右边
    拖动，只允许用户调整部件开启关闭"）。

    没有正文区：正文的识别结果来自真实段落，这时候压根没有文档、没有段落
    可言，画出来也没有意义。跟 `build_rows()` 的各个"else: 空位"分支完全
    对齐（同一套 BLOCK_NAME/CONTENT_STRUCTS/FULL_WIDTH/BACK_ROLES），只是
    每个块永远走"没识别到"那一条。

    拖拽本身不用额外禁用：空位没有抓手（`_draw_cell_content` 的 "slot"
    分支从不把自己加进 `self.draggables`），全表都是空位，天然就拖不动，
    只有左侧的启用勾选框能点。
    """
    rows: list[Row] = []
    counter = [0]

    def add(**kw) -> Row:
        counter[0] += 1
        row = Row(rid=counter[0], **kw)
        rows.append(row)
        return row

    add(kind="zone", zone="front")
    # 顺序跟 `build_rows()` 里真实文档的部件顺序一致（也是规范要求的页序）：
    # 题目/副标题 → 摘要/关键词 → 英文题目 → 英文摘要 → 英文关键词 → 目录。
    # 英文题目 2026-09-14 起是独立部件，以前在空骨架里被排在副标题后面
    # （跟真实文档的顺序对不上），一并对齐。
    for key in ("题目", "副标题"):
        add(kind="slot", col="text", block=key, slot_for=key, slot_col="text", zone="front")
    for key in ("摘要", "关键词"):
        add(kind="slot", col="label", block=key, slot_for=key, slot_col="label", zone="front")
        add(kind="slot", col="text", block=key, slot_for=key, slot_col="text", zone="front")
    add(kind="slot", col="text", block="英文题目", slot_for="英文题目", slot_col="text", zone="front")
    add(kind="slot", col="text", block="英文副标题", slot_for="英文副标题", slot_col="text", zone="front")
    for key in ("英文摘要", "英文关键词"):
        add(kind="slot", col="label", block=key, slot_for=key, slot_col="label", zone="front")
        add(kind="slot", col="text", block=key, slot_for=key, slot_col="text", zone="front")
    add(kind="slot", col="label", block="目录", slot_for="目录", slot_col="label", zone="front")
    for key in ("插图清单", "附表清单", "插图和附表清单"):
        add(kind="slot", col="label", block=key, slot_for=key, slot_col="label", zone="front")

    add(kind="zone", zone="back")
    for key in BACK_ROLES:
        add(kind="slot", col="label", block=key, slot_for=key, slot_col="label", zone="back")
        add(kind="slot", col="text", block=key, slot_for=key, slot_col="text", zone="back")

    for i, row in enumerate(rows):
        row.rid = i + 1
    annotate_blocks(rows)
    return rows


def annotate_blocks(rows: list[Row]) -> None:
    prev: Row | None = None
    for row in rows:
        row.block_first = bool(row.block) and (prev is None or prev.block != row.block)
        prev = row
    nxt: Row | None = None
    for row in reversed(rows):
        row.block_last = bool(row.block) and (nxt is None or nxt.block != row.block)
        nxt = row


def _role_positions(rows: list[Row], roles: tuple[str, ...]) -> dict[str, int]:
    """`roles` 里每个部件在当前行序里第一次出现的位置（`rows` 的下标）。"""
    pos: dict[str, int] = {}
    for i, r in enumerate(rows):
        # 空位（`kind=="slot"`）不算"这个部件存在"：只是等着被填的占位，
        # 位置排在最后是布局约定，拿它跟真实内容比先后会报出假的"顺序颠倒"
        # （2026-09-19 实测：有致谢、没有附录的文档，附录的空位排在致谢
        # 后面，被报成「附录应在致谢之前」，而且无法自动修正）。
        if r.kind == "slot":
            continue
        if r.block in roles and r.block not in pos:
            pos[r.block] = i
    return pos


def find_order_violations(rows: list[Row], roles: tuple[str, ...]) -> list[tuple[str, str]]:
    """检查 `roles` 规定的规范顺序，在**当前存在**的部件之间有没有颠倒。

    只比较相邻（按规范顺序）的两个部件——足够定位"谁跟谁反了"，不产生
    冗余的传递性结论。缺失的部件（没有位置）不参与比较。返回
    `[(应该更靠前却排在后面的部件, 它前面那个反而排在它后面的部件), ...]`
    ——即 `(a, b)` 表示 a 按规范该排在 b 前面，但现在 a 排在 b 后面。
    """
    pos = _role_positions(rows, roles)
    present = [r for r in roles if r in pos]
    violations: list[tuple[str, str]] = []
    for i in range(len(present) - 1):
        a, b = present[i], present[i + 1]
        if pos[a] > pos[b]:
            violations.append((a, b))
    return violations


def _find_block_span(rows: list[Row], block: str) -> tuple[int, int] | None:
    """`block` 在 `rows` **当前顺序**下的那一段区间 `[start, end)`（不含）。

    2026-09-19 改：不再自己按 `r.block == block` 挨个比对——那是相信一个
    可能过期的标签；现在借 `compute_span_ownership()` 现场算的 `spans`
    （"块"＝当前排在哪个标签行后面，见该函数说明），只取第一段匹配的
    （正常情况下一个部件只会有一段，除非用户手动把内容拆得东一块西一块，
    这种边界情况下按"第一段"处理，跟旧版本口径一致）。
    """
    _owner, spans = compute_span_ownership(rows)
    found = next(((s, e) for s, e, p in spans if p == block), None)
    if found is None:
        # 整个部件都是空位（没识别到内容，比如只有标签位＋正文位两个虚线
        # 框）：`compute_span_ownership` 不把空位算进任何块，这里退回"挂着
        # 这个部件名的那一段连续的行"。不然只会拖走其中一个空位，把一个部件
        # 拆成上下两个（2026-09-19 用户截图：出现两个"附录"）。
        start = end = None
        for i, r in enumerate(rows):
            if r.block == block and r.kind != "zone":
                if start is None:
                    start = i
                end = i + 1
            elif start is not None:
                break
        return (start, end) if start is not None else None
    start, end = found
    # 紧挨着的、属于同一个部件的空位（"＋行"加的正文位、没填的标签位）
    # 也是这个块的一部分，整块移动时必须一起走。
    while start > 0 and rows[start - 1].kind == "slot" and rows[start - 1].block == block:
        start -= 1
    while end < len(rows) and rows[end].kind == "slot" and rows[end].block == block:
        end += 1
    return start, end


def _swap_blocks(rows: list[Row], block_a: str, block_b: str) -> bool:
    """把 `block_a`、`block_b` 各自的整段（标签+内容，连续区间）互换位置，
    块内部顺序不变；夹在两块之间的其它内容原样保留在"中间"。返回是否真的
    换了（找不到任一块时不动，返回 False）。"""
    span_a = _find_block_span(rows, block_a)
    span_b = _find_block_span(rows, block_b)
    if span_a is None or span_b is None:
        return False
    (a0, a1), (b0, b1) = span_a, span_b
    if a0 == b0:
        return False
    if a0 > b0:
        (a0, a1), (b0, b1) = (b0, b1), (a0, a1)
    seg_first, middle, seg_second = rows[a0:a1], rows[a1:b0], rows[b0:b1]
    rows[a0:b1] = seg_second + middle + seg_first
    return True


def _swap_block_group(rows: list[Row], src_blocks: list[str], dst_block: str) -> bool:
    """把 `src_blocks`（可以是好几块，比如长按拖动时多选跨了几个部件）作为
    一个整体，和 `dst_block` 互换位置——`src_blocks` 内部各块的相对顺序不变，
    两组之间夹着的其它内容原样留在"中间"。`dst_block` 在 `src_blocks` 范围
    内部时不动（没有意义）。返回是否真的换了。"""
    spans = []
    for b in src_blocks:
        span = _find_block_span(rows, b)
        if span is None:
            return False
        spans.append(span)
    spans.sort()
    src_start, src_end = spans[0][0], spans[-1][1]
    dst_span = _find_block_span(rows, dst_block)
    if dst_span is None:
        return False
    d0, d1 = dst_span
    if d0 >= src_start and d1 <= src_end:
        return False
    # ⚠️ 2026-09-19 修：两块紧挨着（中间没有夹别的内容，比如"参考文献"内容
    # 刚结束、"附录"标签紧接着）时 `d0 == src_end`——原来这两条分支都用严格
    # 不等号，谁都接不住这个最常见的"正好挨着"情形，调换直接静默失败（用户
    # 原话："哪都放不进去"）。改成 `>=`/`<=` 都含等号，`middle` 自然会算成
    # 空列表，效果等价于"两块直接互换"。
    if d0 >= src_end:
        middle, seg_src, seg_dst = rows[src_end:d0], rows[src_start:src_end], rows[d0:d1]
        rows[src_start:d1] = seg_dst + middle + seg_src
        return True
    if d1 <= src_start:
        middle, seg_dst, seg_src = rows[d1:src_start], rows[d0:d1], rows[src_start:src_end]
        rows[d0:src_end] = seg_src + middle + seg_dst
        return True
    return False


def apply_order_fix(rows: list[Row], roles: tuple[str, ...]) -> None:
    """把 `roles` 规定的规范顺序，套到当前存在的部件上——反复交换相邻的
    "顺序反了"的两块，直到不再有相邻颠倒为止（冒泡：每次只动一对，块内部
    与块之间的其它内容都不丢不乱）。"""
    for _guard in range(len(roles) * len(roles)):     # 有限步内必然收敛，防御性上限
        violations = find_order_violations(rows, roles)
        if not violations:
            return
        a, b = violations[0]
        if not _swap_blocks(rows, a, b):
            return


def ensure_min_content_rows(rows: list[Row]) -> None:
    """"标签+正文"型结构：至少给一个正文位，避免"拖走一行后没处拖回"。"""
    for key in CONTENT_STRUCTS:
        indices = [i for i, r in enumerate(rows) if r.block == key]
        if not indices:
            continue
        has_content = any(
            rows[i].kind in ("content", "body")
            or (rows[i].kind == "slot" and rows[i].slot_col == "text")
            for i in indices)
        if not has_content:
            add_content_slot(rows, key, added=False)


def add_content_slot(rows: list[Row], key: str, added: bool = True) -> bool:
    """在结构块的最后一行后面插一个"正文位"空位（多一个落点）。

    added=True＝用户手加的（带 ✕ 可删）；默认补的正文位 added=False 不许删。
    """
    indices = [i for i, r in enumerate(rows) if r.block == key]
    if not indices:
        return False
    last = indices[-1]
    rows.insert(last + 1, Row(rid=0, kind="slot", col="text", block=key, slot_for=key,
                              slot_col="text", zone=rows[last].zone, added=added))
    return True


# ================================================================ 窗口


class AdjustWindow:
    def __init__(self, root: tk.Misc, doc_path: pathlib.Path | None,
                 initial_rows: list[Row] | None = None,
                 initial_index: list[Entry] | None = None,
                 preset_baseline: dict | None = None,
                 cover_end: int | None = None,
                 delete_cover: bool = False,
                 open_rules_callback=None,
                 insert_text: dict[str, str] | None = None,
                 renumber_refs: bool = True,
                 reference_style: str | None = None,
                 heading_scheme: str | None = None,
                 heading_scheme_callback=None) -> None:
        """`root` 可以是独立弹窗，也可以是随便一个容器（2026-09-14 起改成
        嵌进主窗口"论文结构"页签，不再自己弹一个 Toplevel）。

        `heading_scheme`：原文的标题编号方案（"chinese"/"dotted"，见
        `recognize.grammar_for`）——这份表格里所有"重新识别"都按它认标题；
        传了 `initial_rows` 时，调用方要保证那份行是按同一个方案识别的。
        `heading_scheme_callback(scheme)`：用户 Ctrl+Z 撤回了"切换标题编号
        方式"时通知调用方（主窗口把导入区的单选拨回去）。

        `preset_baseline`：点"重置"时重新识别，也要重新按"预设∪识别到的
        部件"这条规则定一遍 ENABLED 的初始值（见 `seed_enabled_from`）——
        这份基准就是打开这份文档那一刻拍下来的预设 `structure_required`
        快照，跟"重置"要退回的"刚识别完"的状态是同一个时间点，理应一致。

        `insert_text`：预设「结构部件」里"插入文字"那一格的快照（部件名→
        写法）。结构页拿它**把行显示成导出时的写法**（用户 2026-09-15）：
        摘要行显示【摘要】、引言行显示引　言、编号标题显示规范编号……
        没传（独立跑这个文件的原型模式、或还没拿到预设）就一律显示原文。

        `reference_style`：预设「引注」里的「参考文献编号样式」（`[1][2][3]` /
        `1.2.3.` / `(1)(2)(3)`）——参考文献条目的显示要跟着它变（用户
        2026-09-16：条目得显示成导出时会写成的那一种序号）。

        `initial_index`：跟 `initial_rows` 配对的识别索引（`build_rows()`
        第二个返回值）——两者必须成对传或成对不传，导出时
        `doc_index.build_ops(self.rows, self.index, ...)` 要靠它把面板当前
        的行序换算回身份更新。
        """
        self.root = root
        self.doc_path = doc_path
        self._insert_text = dict(insert_text) if insert_text else {}
        self.reference_style = reference_style
        self.heading_scheme = heading_scheme
        self._heading_scheme_cb = heading_scheme_callback
        # "自动重新编号参考文献"开关（用户 2026-09-15：“允许用户自定”）：表里
        # 那个复选框就是它，跟着导出走（`pipeline.run()` 的
        # `renumber_references` 形参，见 `main_window.on_export_doc`）。
        # 默认**开**。（题注侧那个“自动重新编号题注”复选框 2026-09-16 删掉了：
        # 用户定“我们只管样式、编号与我们无关”，没有可切的东西了。）
        self.renumber_refs = bool(renumber_refs)
        self.cover_end = cover_end
        # 导出时删不删封面（用户 2026-09-15 在"设置封面范围"里加的复选框）。
        self.delete_cover = delete_cover
        # 「查阅识别规则」那个按钮要打开的是主窗口级的说明窗口，这里只存回调；
        # 独立跑这个文件的原型模式没有回调，按钮就只往状态栏说一句话。
        self._open_rules_cb = open_rules_callback
        # 没有文档＝显示空白模板骨架：全表都是空位，只能勾左侧的启用勾选框，
        # 拖不动（本来也没有段落内容可拖）——用来在导入文件之前，先定好
        # "生成空白模板"要包含哪些部件。
        self.blank_mode = doc_path is None
        self._preset_baseline = preset_baseline
        # `initial_rows` 是上一次打开这个窗口时留下的调整结果（同一份文档）：
        # 关闭窗口不等于放弃调整，重新打开要接着上次的状态改，不能每次都退回
        # 刚识别完的原始状态（那样等于每次都要重新调一遍）。空白模板模式下
        # 永远要求调用方显式传 `initial_rows`（build_blank_rows()）——没有
        # doc_path，没法退回去自己识别一份。
        if initial_rows is not None:
            self.rows = copy.deepcopy(initial_rows)
            self.index = initial_index
        elif doc_path is not None:
            self.rows, self.index = build_rows(doc_path, self.cover_end, self.heading_scheme)
        else:
            self.rows = build_blank_rows()
            self.index = None
        # "重置"要退回的那一份：**构造这一刻的行**（＝刚识别完，或主窗口这次
        # 重建时拿到的识别结果）。用户 2026-09-16 定："既然已经有重新识别了，
        # 重置不重新识别"——重置只撤销手动修改，不再读盘重跑识别（读盘归
        # 「重新识别」按钮）。
        self._pristine_rows = copy.deepcopy(self.rows)
        self._pristine_index = self.index
        # 导出实际会用的那一份行序（用户 2026-09-19 明确要求："GUI 不点确认，
        # 即使点导出也不认"）：拖动/换列/调换这些调整只在**点了"确认"**之后
        # 才会写进这里；导出（`main_window._build_ops_for_export`）永远读
        # `confirmed_rows`，不读实时变动的 `self.rows`。`dirty`＝当前 `rows`
        # 跟这份快照是不是已经不一样了——导出前 `main_window` 用它决定要不要
        # 先弹窗提醒"有调整没确认"。刚建好表、什么都还没拖的时候两者相同，
        # 不用现在就点一下确认才能导出。
        # ⚠️ 行序和识别索引是**成对**的（`Row.doc_index` 指的是这份索引里的
        # 段号）：重新识别/设置封面范围/切换标题编号方案都会换掉 `self.index`，
        # 但它们跟拖动一样只是"没确认的调整"、可以 Ctrl+Z——所以确认点也存一
        # 份索引（`confirmed_index`），导出永远拿"上次确认的行序 + 当时的
        # 索引"这一对，不会拿旧行去配新索引（用户 2026-09-19："只有点击确定
        # 才写入"）。
        self.confirmed_rows = copy.deepcopy(self.rows)
        self.confirmed_index = self.index
        self.confirmed_settings = self._settings()      # 封面范围也是"确认了才作数"
        # 部件启用勾选（`ENABLED`）跟行序一样，也得有"确认了才作数"的快照——
        # 它是全模块共享的一个字典，`_build_ops_for_export` 以前直接读它的
        # 实时值，导致"有调整还没确认"的警告（读 `dirty`）跟实际导出结果
        # （读 `ENABLED` 现值）对不上号：警告说"不会应用"，其实已经应用了
        # （2026-09-22 审计揪出的真 bug）。
        self.confirmed_enabled = dict(ENABLED)
        self.dirty = False
        self.font_main = tkfont.Font(family=FAMILY, size=10)
        self.font_bold = tkfont.Font(family=FAMILY, size=10, weight="bold")
        self.font_small = tkfont.Font(family=FAMILY, size=9)
        # 底栏那三条"如果你…，你可以…"的提示词专用：比按钮小一大截，靠
        # wraplength 折成和按钮一样宽（用户 2026-09-15："字号大幅调小并允许
        # 换行，最好和按钮一样长"）。
        self.font_hint = tkfont.Font(family=FAMILY, size=8)
        self.font_title = tkfont.Font(family=FAMILY, size=11, weight="bold")
        # "word自动目录"这类域行专用：斜体，跟正文/标题都不撞
        self.font_field = tkfont.Font(family=FAMILY, size=10, slant="italic")
        # 列宽不写死，按当前 DPI 的真实字体量出来——格子要"装得下最长的字"，
        # 不是赌一个像素数。
        # 结构列 = 最长块名 + 勾选框/"＋行"的余量（92 是 chrome 余量，见
        # work/_cell_w.txt 的量测数据："插图和附表清单" 91px @100% 缩放）。
        self.w_struct = max(W_STRUCT, self.font_bold.measure("插图和附表清单") + 92)
        # 还得装得下**最宽的一组小勾选框**（☐引用 ☐图 ☐表 ☐续）——勾了图/表
        # 之后会多出"续"，装不下就溢出到右边那一列/窗口外（用户 2026-09-17：
        # 「显示续时续超出窗口」）。实测那组宽 ~234px 而列右界只有 233px。
        # 宽度与绘制共用 `_SPECIAL_CHIPS` / `_CHIP_GAP_*`，别再各算一套。
        self.w_struct = max(self.w_struct, self._special_marks_width() + 24)
        # 标签列要装下的最长空位提示词 = "留白将生成占位字符"（8 字，见
        # `_slot_note`）。以前这里按整句"「插图和附表清单」未启用（点左侧勾选
        # 框可启用）"（25 字）算，把这列撑到 500px 上下、白占半屏（用户
        # 2026-09-14：提示词改短了，这一列自然跟着瘦下来）。
        self.w_label = max(170, self.font_small.measure("留白将生成占位字符") + 24)
        # 一~四级标题那四列：只放两字表头 + 一点内边距（COL_LAYOUTS 里的 0 就是
        # "用这个值"，见 _HEAD_KEYS）。高 DPI 下字号量出来更大，这里自动跟着长。
        self.w_head = max(44, self.font_small.measure("一级") + 20)

        # 交互状态
        self.cells: list[tuple] = []       # (x0,y0,x1,y1,row,col)
        self.toggles: list[tuple] = []     # (x0,y0,x1,y1,kind,payload)
        self.draggables: list[tuple] = []  # (x0,y0,x1,y1,row)
        # 整块拖动的抓手（2026-09-19 新增，块头那个"☰"）：(x0,y0,x1,y1,row)，
        # `on_press` 会先查这份列表，命中就标记这一下拖动是"整块"（见
        # `_begin_drag`），不用用户手动全选。
        self.block_handles: list[tuple] = []
        self._drag: dict | None = None
        self._drag_is_block = False
        # 拖动中自动滚动（2026-09-19 新增，见 `_autoscroll_tick`）：定时器
        # 句柄 + 最近一次已知的鼠标位置（窗口相对坐标，不是画布坐标）。
        self._autoscroll_job = None
        self._drag_last_xy = (0, 0)
        self._resize_job = None
        self._last_width = 1500
        self._flash_gen = 0
        # 冻结表头副本的现状：`(当前区, 副本的 y_top, 副本高度)`；None ＝ 现在
        # 没有副本。滚动时靠它判断"挪位置"还是"重建"（见 `_overlay_header`），
        # 点击判定靠它避开"点中副本、命中被它盖住的行"（见 `_overlay_covers`）。
        self._hdr_key: tuple[str, float, float] | None = None
        # relayout 的产物。必须在 yscrollcommand 接线（configure 画布）**之前**
        # 就绪——Tk 配置画布时会立刻回调一次 yscrollcommand，那里要读 laid。
        self.laid: list[dict] = []
        self.content_w = 0
        self.total_h = 0
        self._undo: list[dict] = []          # 撤回点的结构见 `_push_undo`
        # 多选：存被选中 Row 对象的 id()（不用 rid——claim/删空位都会重新编 rid
        # 或者整个换一个新 Row 对象，rid 不是稳定的身份标识）。选中状态只是
        # "准备拖动"这一下的临时状态，表结构一旦被改动（认领/删空位/重置/
        # 撤回）就清空，不追求跨编辑存活。
        self._selected: set[int] = set()

        # ---- 顶栏：只留标题信息，不放按钮——按钮统一挪到最下面（用户
        # 2026-09-14："原有窗口上面那一栏的按钮一概不要"）----
        top = tk.Frame(root, bg="#e9e9e6")
        top.pack(side="top", fill="x")
        tk.Label(top, text="识别结果调整", font=("Microsoft YaHei UI", 12, "bold"),
                 bg="#e9e9e6", fg="#333").pack(side="left", padx=(12, 6), pady=6)
        # 没有文档时原来写的是“数据：（尚未导入文件——空白模板骨架…）”，
        # 用户 2026-09-16 点名删掉，所以这种时候标签留空。
        data_text = f"数据：{doc_path.name}" if doc_path is not None else ""
        tk.Label(top, text=data_text, bg="#e9e9e6", fg="#767676",
                 font=(FAMILY, 9)).pack(side="left")
        # 上栏只披露标题和文件名——封面范围之类的解释性文字一律不摆了
        # （用户 2026-09-15："把所有这种话全部删掉"）。

        # ---- 画布 + 滚动条（纵向 + 横向）----
        # 横向滚动条：内容宽（左侧列展开的、正文长行）超出视口
        # 时出现，靠它看全；纵向滚动时列头由 _overlay_header 钉在视口顶
        # （冻结表头，用户 2026-09-14 要求）。
        mid = tk.Frame(root, bg=C_APP)
        mid.pack(side="top", fill="both", expand=True)
        self.canvas = tk.Canvas(mid, bg=C_APP, highlightthickness=0, yscrollincrement=24)
        vsb = tk.Scrollbar(mid, orient="vertical", command=self.canvas.yview)
        hsb = tk.Scrollbar(mid, orient="horizontal", command=self.canvas.xview)
        # yscroll 走自己的回调：更新滚动条的同时把列头副本钉在视口顶
        self._vsb = vsb
        self.canvas.configure(yscrollcommand=self._on_yscroll, xscrollcommand=hsb.set)
        # pack 顺序有讲究：底部横条先占位，画布才能 fill both expand
        hsb.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # ---- 底栏（用户 2026-09-15 重做）----
        # 布局：① 一行**居中加粗**的总说明；② 左边三条"如果你…，你可以…"提示、
        # 每条下面一个按钮，右边"重置/确认"；③ 最下面一行快捷键与注意事项。
        # 原来那一大段图例（☐引用＝长引文 ｜ ☑图/☑表＝题注 ｜ 六圆点＝…）**整段
        # 删掉了**——用户原话"删除这个页面的全部说明文字"。那些符号的说明改由
        # 「查阅识别规则」那个窗口（将来那份识别规则全表）承担。
        bottom = tk.Frame(root, bg="#e9e9e6")
        bottom.pack(side="bottom", fill="x")
        tk.Label(bottom, text="拖动文字或勾选组件以调整标题结构",
                 font=(FAMILY, 11, "bold"), bg="#e9e9e6", fg="#333").pack(
            side="top", pady=(6, 4))

        hint_row = tk.Frame(bottom, bg="#e9e9e6")
        hint_row.pack(side="top", fill="x", padx=12, pady=(0, 6))
        # 左边一整块（四列提示按钮 + 下面那行灰字）自然宽度；"重置/确认"占**右边剩下的
        # 全部空位**、跟左边整块等高（用户 2026-09-19："在保留合理边距的情况下，在空位里居中
        # 占满"）——两个按钮上下叠放、各占一半高，宽度随空位伸缩。
        left_col = tk.Frame(hint_row, bg="#e9e9e6")
        left_col.pack(side="left", anchor="n")
        actions = tk.Frame(hint_row, bg="#e9e9e6")
        actions.pack(side="left", fill="both", expand=True, padx=(18, 0), pady=(2, 2))
        left = tk.Frame(left_col, bg="#e9e9e6")
        left.pack(side="top", anchor="w")

        hint_defs = (
            ("遇到识别问题，可以", "查阅识别规则", self.on_open_rules, True),
            ("如已修改原文档，请", "重新识别", self.on_reidentify, not self.blank_mode),
            ("如文档包含封面，请", "设置封面范围", self.on_set_cover_range, not self.blank_mode),
            ("部件顺序想恢复规范，请", "一键重排", self.on_apply_order_fix, not self.blank_mode),
        )
        # 三列宽度取齐，且**保证提示词一行装得下**（用户 2026-09-15："还是不要
        # 换行的好"）——所以取"按钮文字宽"和"提示词宽"里的最大值再加一截余量：
        # 按钮靠 `fill="x"` 拉长到列宽，提示词不设 wraplength，永远单行。
        # 按当前字体的实测宽算，不写死像素，高 DPI（本机 175%）下自动跟着变。
        col_w = max(
            [self.font_main.measure(text) for _t, text, _c, _e in hint_defs]
            + [self.font_hint.measure(t) for t, _b, _c, _e in hint_defs]
        ) + 12          # 余量 64→28→12：提示词几乎贴着列宽，给"重置/确认"腾地方

        for col in range(len(hint_defs)):
            left.columnconfigure(col, minsize=col_w, uniform="hintcol")
        hint_buttons = []
        hint_cells = []
        for col, (text, button_text, command, enabled) in enumerate(hint_defs):
            cell = tk.Frame(left, bg="#e9e9e6")
            cell.grid(row=0, column=col, sticky="new", padx=(0, 6))
            hint_cells.append(cell)
            tk.Label(cell, text=text, font=self.font_hint, bg="#e9e9e6", fg="#333",
                     anchor="center").pack(side="top", fill="x")
            button = tk.Button(cell, text=button_text, font=(FAMILY, 10),
                               command=command)
            button.pack(side="top", fill="x", pady=(3, 0))
            if not enabled:
                button.configure(state="disabled")
            hint_buttons.append(button)
        # "重置/确认"：上下叠放（用户 2026-09-19："调成上下关系"），填满右边剩下的空位。
        # 顺序沿用原来并排时的左→右：重置在上、确认在下（确认是收尾动作，放最后）。
        # `actions` 里先放一个 0 高的撑宽条，保证空位再小也有最小宽度（reflow 也按它算）。
        action_min_w = 110
        tk.Frame(actions, bg="#e9e9e6", width=action_min_w, height=1).pack(side="top")
        # 大小**不超过另外四个提示按钮**（用户 2026-09-19："别这么大，最大不超过另外四个按钮
        # 的宽高"）：宽取提示按钮列宽、高取提示按钮实测高；空位比它小就跟着缩窄。
        # 两个按钮叠成一摞，整摞在空位里水平、垂直居中（`place` 居中）。
        self.root.update_idletasks()
        btn_h = hint_buttons[0].winfo_reqheight()
        stack = tk.Frame(actions, bg="#e9e9e6")
        stack.place(relx=0.5, rely=0.5, anchor="center")
        holders = []
        for i, (text, command) in enumerate((("重置", self.on_reset), ("确认", self.on_confirm))):
            holder = tk.Frame(stack, bg="#e9e9e6", width=col_w - 6, height=btn_h)   # 列宽含 6px 间距
            holder.pack(side="top", pady=(0 if i == 0 else 6, 0))
            holder.pack_propagate(False)
            tk.Button(holder, text=text, font=(FAMILY, 10), command=command).pack(
                fill="both", expand=True)
            holders.append(holder)

        def fit_actions(e) -> None:
            w = min(col_w - 6, max(e.width - 4, 60))
            for holder in holders:
                holder.configure(width=w)

        actions.bind("<Configure>", fit_actions, add="+")

        # 窗口更窄（小屏/拖窄了分隔条）连一行四列也装不下时，四个提示按钮改排两行两列，
        # 而不是被裁掉——宽度够了自动恢复一行。
        def reflow_hints(_e=None) -> None:
            need = len(hint_cells) * (col_w + 6) + 18 + action_min_w
            cols = len(hint_cells) if hint_row.winfo_width() >= need else 2
            if getattr(self, "_hint_cols", None) == cols:
                return
            self._hint_cols = cols
            for col in range(len(hint_cells)):
                left.columnconfigure(col, minsize=col_w if col < cols else 0,
                                     uniform="hintcol" if col < cols else "")
            for i, cell in enumerate(hint_cells):
                cell.grid(row=i // cols, column=i % cols, sticky="new", padx=(0, 6),
                          pady=(0 if i < cols else 6, 0))

        hint_row.bind("<Configure>", reflow_hints, add="+")

        # 灰字提示放在左边整块的最下面（原来是整行铺开的一条，"重置/确认"够不到它那一行）；
        # 宽度跟四列提示按钮一样宽，窄了换行，不被裁掉。
        foot_label = tk.Label(left_col, text="右键：插入空行；Ctrl+Z：撤回上一步。"
                                             "如需修改段落实际顺序或内容，请重新编辑原文档！",
                              font=(FAMILY, 9), bg="#e9e9e6", fg="#8a8a8a", anchor="w",
                              justify="left")
        foot_label.pack(side="top", anchor="w", pady=(6, 0))
        left.bind("<Configure>", lambda e: foot_label.configure(wraplength=max(e.width, 200)),
                  add="+")

        # ---- 事件 ----
        self.canvas.bind("<Button-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_motion)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.canvas.bind("<Button-3>", self.on_right_click)
        self.canvas.bind("<Motion>", self.on_hover)
        self.canvas.bind("<Configure>", self.on_configure)
        # 滚轮绑在画布本身上，不用 bind_all——嵌进主窗口之后，bind_all 会在
        # 整个应用范围内生效，切到别的页签滚轮也会被这里劫持。
        self.canvas.bind("<MouseWheel>", self.on_wheel)
        # Shift+滚轮＝横向滚动（横向滚动条的键盘/滚轮补充）
        self.canvas.bind("<Shift-MouseWheel>", self.on_hwheel)
        # Ctrl+Z 没法只靠绑定画布解决（键盘快捷键不看鼠标位置，看谁有焦点），
        # 所以还是挂在顶层，但 on_undo 里会先检查这个页签是不是当前显示的，
        # 不是就什么都不做——不然切到别的页签按 Ctrl+Z 会撤到这张表上来。
        # ⚠️ 这个绑定是**顶层级**的，销毁时必须摘掉：主窗口重建结构页时会先
        # destroy 掉旧的表格，而绑定还指着已销毁实例的 on_undo —— 那时候按
        # Ctrl+Z 就是 on_undo → 已销毁的 canvas（TclError → 崩溃弹窗）。
        self._undo_toplevel = root.winfo_toplevel()
        self._undo_toplevel.bind_all("<Control-z>", self.on_undo)
        # 本类不是 Tk 控件（只是个普通类，控件都挂在传进来的 root 上），所以
        # 销毁钩子挂在画布上——主窗口重建结构页时会 destroy 掉这些子控件。
        self.canvas.bind("<Destroy>", self._on_destroy, add="+")

        root.update_idletasks()
        self._last_width = max(self.canvas.winfo_width(), 1280)
        self.relayout()
        self.draw()

    # ------------------------------------------------------------ 布局

    def relayout(self) -> None:
        """排版两遍法（用户 2026-09-14：表格不再"…"截断，改内容驱动）：

        第一遍按**内容实测**定各列最终宽——**正文一律单行到底、不折行也不缩进**
        （用户 2026-09-14 晚），列宽被该区最长的那一行撑开；标签列同理按内容
        展开；一~四级标题那四列只按表头实测宽给（`w_head`）。第二遍铺 y 坐标
        和行高。内容总宽超过视口的部分靠横向滚动条看（见 __init__ 的 hsb）。
        """
        width = self._last_width
        avail = max(640, width - W_NUM - self.w_struct - W_SELECT - 14)
        pad_left = 9 + GRIP_ROOM           # _draw_cell_content 里文字的左缘偏移
        pad_right = 6
        self.laid: list[dict] = []

        # ---- 第一遍：按区收集内容需求宽 ----
        zone_rows: dict[str, list[Row]] = {}
        zone = None
        for row in self.rows:
            if row.kind == "zone":
                zone = row.zone
                zone_rows.setdefault(zone, [])
            elif zone is not None:
                zone_rows[zone].append(row)

        def chip_extra(row: Row) -> int:
            chip = _caption_chip(row)
            return self.font_small.measure(chip) + 12 + 6 if chip else 0

        def text_need(row: Row) -> int:
            """这一行正文"单行到底"需要多宽（含左侧抓手留白与题注小标签）。

            用户 2026-09-14 晚：**一律不折行、不缩进**——装不下就把列宽撑开，
            超长部分靠横向滚动条看。以前"切成等长两行 + 首行缩进两个全角空格"
            的那套（`_split_two_lines`）已整体删除。

            ⚠️ 量的是 `self.shown(row)`（导出时的写法）而不是 `row.text`：
            显示的就是那个字符串，两边量/画必须同一个来源——拿原文量、拿改后
            的文字画，长出来的那几个字就会越出格线。
            """
            return (self.font_main.measure(self.shown(row)) + pad_left + pad_right
                    + chip_extra(row))

        def col_w(key: str, w, label_w: int) -> int:
            """这一列该给多宽：标签列用本区实测宽，一~四级标题用 w_head（表头实测）。

            正文列（w 为 None）不走这里——它由内容撑开，见下面的 zone_colw。
            """
            if key == "label":
                return label_w
            if key in _HEAD_KEYS:
                return self.w_head
            return w

        zone_colw: dict[str, int] = {}
        zone_fixed: dict[str, int] = {}
        zone_label_w: dict[str, int] = {}
        for zn, rows in zone_rows.items():
            # 标签列先量：它的实测宽是"正文列拿剩余"求和的一部分
            label_w = self.w_label
            for row in rows:
                if row.col == "label":
                    text = self._slot_hint(row) if row.kind == "slot" else self.shown(row)
                    f = self._slot_font(row) if row.kind == "slot" else self.font_main
                    label_w = max(label_w, f.measure(text) + pad_left + pad_right)
            fixed = sum(col_w(k, w, label_w)
                        for k, _h, w in COL_LAYOUTS[zn] if w is not None)
            rest = max(240, avail - fixed)
            colw = rest
            for row in rows:
                if row.col == "text" and row.kind != "slot":
                    # 单行不折行：列宽直接跟着最长的那一行撑开
                    colw = max(colw, text_need(row))
            zone_colw[zn] = colw
            zone_fixed[zn] = fixed
            zone_label_w[zn] = label_w

        # **所有区的最后一列都拉到同一个右缘**（用户 2026-09-17 截图：
        # “「正文」表头比控件稍微短一点”）。原因：前置/后置区的正文列是按
        # **它自己**的内容宽撑开的，而行底纹、分隔线一律画到整表宽 ⇒ 窄区的列头
        # 就比下面那一条短一截。表宽 = max(各区自然需要的最右边界, 视口宽,
        # 整行文字的溢出) —— 取视口宽是为了让内容装得下时表格仍然填满窗口
        # （否则右边会空一条）。
        natural_right = max(W_NUM + self.w_struct + W_SELECT + zone_fixed[zn] + zone_colw[zn]
                            for zn in zone_colw)
        # 题目/副标题/英文题目这类"整行一格"的文字可能比列分配还宽：也算进表宽
        # （只是把列拉宽，文字照旧从左往右画，不会因此被截断）
        for row in self.rows:
            if row.kind in ("title", "subtitle", "entitle") and row.text:
                f = self.font_title if row.kind == "title" else self.font_bold
                natural_right = max(natural_right, W_NUM + self.w_struct + W_SELECT
                                    + f.measure(self.shown(row)) + pad_left + pad_right)
        table_right = max(natural_right, int(width))

        # ---- 第二遍：铺 y、行高、列位 ----
        zone = "front"
        colx: list[tuple] = []
        y = 0
        for row in self.rows:
            if row.kind == "zone":
                zone = row.zone
                self.laid.append({"kind": "zone", "zone": zone, "y0": y, "y1": y + 30})
                y += 30
                spec: list[tuple[str, str, int]] = []
                for key, hdr, w in COL_LAYOUTS[zone]:
                    # 最后一列（w is None）＝右缘对齐到 table_right 的那一列
                    ww = (table_right - (W_NUM + self.w_struct + W_SELECT + zone_fixed[zone])
                          if w is None else col_w(key, w, zone_label_w[zone]))
                    spec.append((key, hdr, ww))
                x = W_NUM + self.w_struct + W_SELECT
                colx = []
                for key, hdr, ww in spec:
                    colx.append((key, hdr, x, ww))
                    x += ww
                self.laid.append({"kind": "colheader", "zone": zone, "cols": colx,
                                  "y0": y, "y1": y + 28})
                y += 28
                continue
            # 空位高一点：装得下两行提示——9号字实测行高 17px，两行按 20px
            # 间距排布还要留边距，42px 明显不够，两行字会互相压住，字也会
            # 顶到虚线框边缘（用户 2026-09-14 反馈"字叠字""字和虚线方框打架"）。
            # 正文行一律单行高（28px）——折行的行高分支已随折行一起删掉。
            # 行高交给 `_row_height()`：复选框被挤到第二行时它会加高（2026-09-17）。
            cell_w = next((w for k, _h, _x, w in colx if k == row.col), 0.0)
            h = self._row_height(row, cell_w)
            self.laid.append({"kind": "row", "row": row, "zone": zone, "cols": colx,
                              "y0": y, "y1": y + h})
            y += h
        self.total_h = y + 16
        # 表宽只认 `table_right`（第一遍已把各区列宽、整行文字溢出都算进去）。
        # 行底纹、分隔线、列头画的都是它 ⇒ 三者右缘永远对齐。
        self.content_w = int(table_right)

    def draw(self) -> None:
        c = self.canvas
        c.delete("all")
        self.cells = []
        self.toggles = []
        self.draggables = []
        self.block_handles = []
        self._hdr_key = None      # delete("all") 把表头副本也清了（见 _overlay_header）
        for item in self.laid:
            if item["kind"] == "zone":
                self._draw_zone_band(item)
            elif item["kind"] == "colheader":
                self._draw_colheader(item)
            else:
                self._draw_row(item)
        # 横向格线统一最后补画：若跟着行一起画，会被下一行的底色矩形整个盖掉，
        # 结果就是"横线一条都看不见"——所以必须等所有块画完再补一遍。
        for item in self.laid:
            c.create_line(0, item["y1"], self.content_w, item["y1"], fill=C_ROW_LINE)
        c.configure(scrollregion=(0, 0, self.content_w, self.total_h))
        # 滚动位置恢复后把列头钉回视口顶（冻结表头）
        self._overlay_header()

    def _draw_zone_band(self, item: dict) -> None:
        c = self.canvas
        zone = item["zone"]
        y0, y1 = item["y0"], item["y1"]
        c.create_rectangle(0, y0, self.content_w, y1, fill=ZONE_BAND[zone], outline="")
        c.create_text(12, (y0 + y1) / 2, text=ZONE_TEXT[zone], anchor="w",
                      font=self.font_bold, fill="#33404e")
        c.create_line(0, y1, self.content_w, y1, fill="#c9c9c4")

    def _draw_colheader(self, item: dict, y_top: float | None = None,
                        tag: str | None = None) -> None:
        """画一个区的列头行。y_top 给值时＝"冻结副本"：画在视口顶（文档 x 坐
        标，横向跟内容一起滚、永远对齐列位）；tag 用于 overlay 的统一增删。"""
        c = self.canvas
        zone = item["zone"]
        if y_top is None:
            y0, y1 = item["y0"], item["y1"]
        else:
            y0, y1 = y_top, y_top + (item["y1"] - item["y0"])
        kw = {"tags": tag} if tag else {}
        fill = ZONE_HDR[zone]
        c.create_rectangle(0, y0, W_NUM, y1, fill=fill, outline="#d5d5cf", **kw)
        c.create_text(W_NUM / 2, (y0 + y1) / 2, text="段号", font=self.font_small,
                      fill="#555", **kw)
        c.create_rectangle(W_NUM, y0, W_NUM + self.w_struct, y1, fill=fill,
                           outline="#d5d5cf", **kw)
        c.create_text(W_NUM + self.w_struct / 2, (y0 + y1) / 2, text="结构",
                      font=self.font_small, fill="#555", **kw)
        # 多选列（2026-09-19 新增，独立一列，夹在"结构"和区特定列之间）
        sel_x0 = W_NUM + self.w_struct
        c.create_rectangle(sel_x0, y0, sel_x0 + W_SELECT, y1, fill=fill,
                           outline="#d5d5cf", **kw)
        c.create_text(sel_x0 + W_SELECT / 2, (y0 + y1) / 2, text="多选",
                      font=self.font_small, fill="#555", **kw)
        for key, hdr, x0, w in item["cols"]:
            c.create_rectangle(x0, y0, x0 + w, y1, fill=fill, outline="#d5d5cf", **kw)
            c.create_text(x0 + w / 2, (y0 + y1) / 2, text=hdr, font=self.font_small,
                          fill="#445", **kw)

    def _draw_row(self, item: dict) -> None:
        c = self.canvas
        row: Row = item["row"]
        y0, y1 = item["y0"], item["y1"]
        off = bool(row.block and not ENABLED.get(row.block, True)) or (
            row.kind == "slot" and not ENABLED.get(row.slot_for or "", True))
        if off:
            bg = "#ededed"
        elif row.special == "引用":
            bg = "#f6f1fb"
        elif row.special == "图":
            bg = "#f1f8f1"
        elif row.special == "表":
            bg = "#eef7f6"
        else:
            bg = "#ffffff"
        c.create_rectangle(0, y0, self.content_w, y1, fill=bg, outline="")
        # 段号
        num = str(row.doc_index) if row.doc_index is not None else "—"
        c.create_text(W_NUM - 10, (y0 + y1) / 2, text=num, anchor="e",
                      font=self.font_small, fill=C_NUM if row.doc_index is not None else "#c0c0c0")
        # 结构格
        self._draw_struct_cell(row, y0, y1)
        # 多选复选框列（2026-09-19 新增，独立一列，夹在"结构"和区特定列
        # 之间，列头"多选"——用户明确要求，替代 ctrl/shift 点选）：只有真
        # 正能拖动的行才画（空位/封面/目录域这几类本来就不进
        # `self.draggables`，画了也点不出效果，索性不画，省得看着像能勾）。
        sel_x0 = W_NUM + self.w_struct
        c.create_rectangle(sel_x0, y0, sel_x0 + W_SELECT, y1, fill="#fbfbf9", outline="")
        c.create_line(sel_x0, y0, sel_x0, y1, fill="#d9d9d4")
        c.create_line(sel_x0 + W_SELECT, y0, sel_x0 + W_SELECT, y1, fill="#d9d9d4")
        checkbox_eligible = (row.kind != "slot" and row.block != COVER_BLOCK
                             and row.kind != "toc_field")
        if checkbox_eligible:
            checked = id(row) in self._selected
            c.create_text(sel_x0 + W_SELECT / 2, (y0 + y1) / 2,
                          text="☑" if checked else "☐", anchor="center",
                          font=self.font_main, fill="#2f6fbe" if checked else "#a8a89e")
            self.toggles.append((sel_x0, y0, sel_x0 + W_SELECT, y1, "select", row))
        # 内容格（全部列都登记为拖放目标）；文字可以一路写到区带最右边
        cols = item["cols"]
        zone_right = cols[-1][2] + cols[-1][3]
        full_width = row.kind in ("title", "subtitle", "entitle") or (
            row.kind == "slot" and row.slot_for in FULL_WIDTH)
        span: tuple[float, float] | None = None
        if full_width:
            # 题目/副标题：不区分"结构标签/正文"，整行一格
            x0 = W_NUM + self.w_struct + W_SELECT
            self.cells.append((x0, y0, zone_right, y1, row, row.col))
            span = self._draw_cell_content(row, item, x0, y0, zone_right, y1, zone_right)
        else:
            for key, _hdr, x0, w in cols:
                x1 = x0 + w
                self.cells.append((x0, y0, x1, y1, row, key))
                if row.col == key:
                    span = self._draw_cell_content(row, item, x0, y0, x1, y1, zone_right)
        # 附录序号题名：在「结构标签」栏右端标一行小字（用户 2026-09-16 定）。
        # 只在这一行仍属附录时画——用户把它拖到别的部件后，标记就跟着消失。
        if row.appendix_entry and row.block == "附录":
            self._draw_appendix_mark(item, y0, y1)
        # 还是 Word 自动编号的一项、而这行**不再是标题/题注**（比如用户把它拖回
        # "正文"）：在标签栏右端标一句小字，说明那个号是 Word 画的、工具不动它
        # ——不然文字前面凭空多出一个编号，看着像工具写的（用户 2026-09-17）。
        elif row.auto_number and row.heading is None and row.kind in ("body", "slot"):
            # 被拖回“正文”的自动编号项：成品里它的 Word 编号会被取消、写成
            # 手打文字。标记出来，不然用户看不出这串号是会跟着进成品的字。
            self._draw_field_mark(item, y0, y1, "word自动编号")
        # 竖线：只画到文字起点为止。文字（哪怕很短）右侧的列界一律不画——
        # 用户 2026-09-13：前面的格子有字了，后面的格子就不再显示竖框线。
        xs = ([W_NUM, W_NUM + self.w_struct, W_NUM + self.w_struct + W_SELECT] if full_width
              else [W_NUM, W_NUM + self.w_struct] + [x0 for (_k, _h, x0, _w) in cols])
        for lx in xs:
            if span is not None and lx > span[0] - 5:
                continue
            c.create_line(lx, y0, lx, y1, fill=C_SEP)
        # 横线不在这里画——见 draw() 的收尾补线，否则会被下一行底色盖掉
        if id(row) in self._selected:
            c.create_rectangle(1, y0 + 1, self.content_w - 1, y1 - 1,
                               outline="#2f6fbe", width=2)

    def _draw_struct_cell(self, row: Row, y0: float, y1: float) -> None:
        c = self.canvas
        cy = (y0 + y1) / 2
        # 底色（块色）
        if row.block:
            enabled = ENABLED.get(row.block, True)
            c.create_rectangle(W_NUM, y0, W_NUM + self.w_struct, y1,
                               fill=C_STRUCT_BG if enabled else C_STRUCT_BG_OFF, outline="")
        else:
            c.create_rectangle(W_NUM, y0, W_NUM + self.w_struct, y1, fill="#fbfbf9", outline="")
        c.create_line(W_NUM, y0, W_NUM, y1, fill="#d9d9d4")
        # 块头控件；手加空位的删除键（放结构栏——全屏时内容格右端的 ✕ 在屏幕另一头）；
        # 其余行显示行内小勾选（引用/图/表）
        if row.block_first:
            self._draw_block_head(row, cy)
        elif row.kind == "slot" and row.added:
            c.create_text(W_NUM + self.w_struct - 20, cy, text="✕", anchor="center",
                          font=self.font_main, fill="#b0664a")
            self.toggles.append((W_NUM + self.w_struct - 44, cy - 12, W_NUM + self.w_struct - 2,
                                 cy + 11, "delete_slot", row))
        elif row.kind in ("body", "caption"):
            self._draw_special_marks(row, cy)
        # 边框装饰
        if row.block:
            c.create_line(W_NUM + 2, y0, W_NUM + 2, y1, fill=C_STRUCT_LINE)
            if row.block_last:
                c.create_line(W_NUM, y1 - 1, W_NUM + self.w_struct, y1 - 1, fill=C_STRUCT_LINE)
        c.create_line(W_NUM + self.w_struct, y0, W_NUM + self.w_struct, y1, fill="#d9d9d4")

    def _draw_block_head(self, row: Row, cy: float) -> None:
        """块头控件：启用勾选＋块名＋拖动整块的抓手（☰）＋"＋行"（加空位）。"""
        c = self.canvas
        enabled = ENABLED.get(row.block, True)
        c.create_text(W_NUM + 8, cy, text="☑" if enabled else "☐", anchor="w",
                      font=self.font_main, fill="#2e7d32" if enabled else "#9a9a9a")
        name = BLOCK_NAME.get(row.block, row.block)
        c.create_text(W_NUM + 30, cy, text=name, anchor="w", font=self.font_bold,
                      fill="#444444" if enabled else "#9a9a9a")
        name_end = W_NUM + 30 + self.font_bold.measure(name)
        self.toggles.append((W_NUM + 2, cy - 12, name_end + 6, cy + 11,
                             "block", row.block))
        # 拖动整块的抓手（2026-09-19 新增，用户原话："结构名后加一个三横线
        # 标识，按住结构名直接拖动，结构里的所有东西都带着动。不必手动
        # 全选"）——单独一块可拖拽区域，跟左边"启用勾选＋块名"那个点击区
        # 分开：点名字是切换启用，按住这个抓手才是拖动整块（见 `on_press`/
        # `_begin_drag`）。
        handle_x = name_end + 14
        c.create_text(handle_x, cy, text="☰", anchor="w", font=self.font_main,
                      fill="#7a7a72")
        self.block_handles.append((handle_x - 4, cy - 12, handle_x + 16, cy + 12, row))
        if enabled and row.block in CONTENT_STRUCTS and not self.blank_mode:
            # "＋行"：给这个结构加一个空位（额外的加入位置）——空白模板模式
            # 没有意义：没有文档、没有段落，加出来的空位永远不会有东西可拖进
            # 去，只是摆设，跟右键插入空位的禁用理由一样。
            c.create_text(W_NUM + self.w_struct - 22, cy, text="＋行", anchor="center",
                          font=self.font_small, fill="#4b8b4b")
            self.toggles.append((W_NUM + self.w_struct - 44, cy - 12, W_NUM + self.w_struct - 2,
                                 cy + 11, "insert", row.block))

    # 行内小勾选框：文字（顺序＝绘制顺序）与间距。**绘制和量宽共用这一份**——
    # `_special_marks_width()` 拿它算结构列要留多宽，`_draw_special_marks`
    # 拿它摆位置（2026-09-17 之前是两边各写一套常量）。
    _SPECIAL_CHIPS = ("引用", "图", "表", "续")
    _CHIP_GAP_AFTER_BOX = 8          # 方框与它的文字之间
    _CHIP_GAP_BETWEEN = 16           # 一组与下一组之间

    def _special_marks_width(self) -> int:
        """一行里那组小勾选框**最宽**要占多少（＝四个框都画时的总宽）。

        给 `__init__` 算 `w_struct` 用：结构列必须装得下它，否则最后一个
        "☐续"会溢出到右边那一列/窗口外（用户 2026-09-17）。
        """
        box = self.font_small.measure("☑")
        labels = sum(self.font_small.measure(v) for v in self._SPECIAL_CHIPS)
        n = len(self._SPECIAL_CHIPS)
        return (n * box + labels
                + self._CHIP_GAP_AFTER_BOX * n
                + self._CHIP_GAP_BETWEEN * (n - 1))

    def _draw_special_marks(self, row: Row, cy: float) -> None:
        r"""行内小勾选：☐引用（长引文）☐图 ☐表（题注）☐续（续表/续图）。

        位置按实际字宽现算，不用写死的间距凑数——"引用"是两个字，比"图"/"表"
        宽一倍，之前统一按固定偏移量摆，方框和文字之间只剩几像素，看着像
        挤在一起（用户 2026-09-14 反馈"引用框那里文字太近了，打架"）。

        **只要勾了图/表就画"续"**（用户 2026-09-17 改的口径）：它是个
        "叠加开关"，勾上＝这一行是续图/续表（仍算题注，但**不挂样式** ⇒
        **不进清单**；也**不涉及编号**：2026-09-16 起工具不写任何域），
        取消＝回到这一行**识别时的角色**（见 `Row.caption_role_base`）。
        ⚠️ 旧版只在 `role != "en"` 时画（2026-09-16 审计为防"英文行点两下
        变成真题注"加的例外）——但用户看到的现象是"自动勾上的图/表不显示续，
        得先取消再重勾"，那个例外反而成了阻碍；现在改成"全画 + 取消时回退
        到原角色"，两个毛病都治了。
        """
        c = self.canvas
        box_w = self.font_small.measure("☑")
        labels = list(self._SPECIAL_CHIPS)
        if row.special not in ("图", "表"):
            labels.remove("续")          # 没勾图/表时"续"没有意义，不画
        label_w = {v: self.font_small.measure(v) for v in labels}
        gap_after_box, gap_between = self._CHIP_GAP_AFTER_BOX, self._CHIP_GAP_BETWEEN
        # 保险丝：算一下这几组东西实际要占多宽，超过"结构"格能给的宽度就按
        # 比例缩间距——不同机器的字体/DPI 渲染宽度不完全一样，写死的间距在
        # 某些环境下可能还是会顶到下一列，这里按实测宽度自适应，不再赌数字。
        #（列宽本身也已经按最宽那组留足了，见 `__init__` 里的 `w_struct`。）
        avail = self.w_struct - 16
        total = sum(box_w + gap_after_box + label_w[v] for v in labels) + gap_between * (len(labels) - 1)
        if total > avail > 0:
            shrink = avail / total
            gap_after_box = max(3, int(gap_after_box * shrink))
            gap_between = max(6, int(gap_between * shrink))
        x = W_NUM + 10
        for label in labels:
            val = label
            # "续"看的是角色（勾上＝角色变成"续"），其余三个看的是 `special`。
            checked = ((row.caption_role or "main") == "cont" if val == "续"
                       else row.special == val)
            fg = C_CHIP[val] if checked else "#b3b3aa"
            c.create_text(x, cy, text="☑" if checked else "☐", anchor="w",
                          font=self.font_small, fill=fg)
            label_x = x + box_w + gap_after_box
            c.create_text(label_x, cy, text=label, anchor="w", font=self.font_small,
                          fill=fg if checked else "#a6a69e")
            label_end = label_x + label_w[val]
            self.toggles.append((x - 4, cy - 11, label_end + 4, cy + 11, "special", (row, val)))
            x = label_end + gap_between

    def _draw_appendix_mark(self, item: dict, y0: float, y1: float) -> None:
        r"""附录序号题名（`附录A …`）在「结构标签」栏右端的小字标记。

        用户 2026-09-16 定："在 GUI 放到标签那一栏吧，但靠右对齐。"靠右是因为
        这一栏本来就是"标签"的位置（左边留给标签行的文字），右端贴边既不跟
        别处抢位置，也能一眼看出这一行不是普通正文：它按规范 [226] 要套 S20
        （五号宋体），跟上下文都不一样。
        """
        self._draw_field_mark(item, y0, y1, "附录题名")

    def _draw_field_mark(self, item: dict, y0: float, y1: float, text: str) -> None:
        """在「结构标签」栏右端画一句小字（斜体灰蓝，跟"word自动目录"同一套
        用语）。目前两个用途：`附录题名`、`word自动编号`。"""
        label = next((c for c in item["cols"] if c[0] == "label"), None)
        if label is None:
            return
        _key, _hdr, x0, w = label
        self.canvas.create_text(x0 + w - 12, (y0 + y1) / 2, text=text,
                                anchor="e", font=self.font_field, fill=C_FIELD)

    def _draw_cell_content(self, row: Row, item: dict, x0: float, y0: float, x1: float,
                           y1: float, right: float) -> None:
        c = self.canvas
        cy = (y0 + y1) / 2
        pad = 9
        if row.kind == "slot":
            # 提示文字不再截断（用户 2026-09-14：表格去"…"，列宽按内容撑开，
            # 超出视口靠横向滚动看）；虚线框仍按格子画。
            c.create_rectangle(x0 + 4, y0 + 4, x1 - 4, y1 - 4, outline=C_SLOT_LINE,
                               dash=(5, 3), width=1)
            hint = self._slot_hint(row)
            key = row.slot_for or ""
            note = self._slot_note(row)
            enabled = ENABLED.get(key, True)
            font = self._slot_font(row)
            # "将生成Word自动目录/清单"这类提示用域行的灰蓝斜体，跟正文区分开
            fill = (C_FIELD if font is self.font_field
                    else (C_SLOT_TEXT if enabled else "#a9a9a9"))
            widest = font.measure(hint)
            if note is not None:
                widest = max(widest, self.font_small.measure(note))
            # 内容按"行"排（见 `_layout_lines`）：提示词、说明、复选框依次往下，
            # 整块在格子里竖直居中。复选框装不下就换到下一行；文字本身还装不下
            # 就折成两行（用户 2026-09-17：“这个超出了虚线框。可以换行处理
            # （但注意高度不要超）”）。行高由 `_row_height()` 用**同一套**排法算，
            # 两边不会各说各话。
            switch_lines, switch_x, _wrapped = self._switch_plan(
                row, x0 + 14 + widest + 20, x1, x0 + 4)
            items = ["hint"] + (["note"] if note is not None else []) \
                + ["switch"] * len(switch_lines)
            ys = self._layout_lines(cy, len(items))
            c.create_text(x0 + 14, ys[0], text=hint, anchor="w",
                          font=self.font_small if note is not None else font,
                          fill=fill)
            if note is not None:
                c.create_text(x0 + 14, ys[1], text=note, anchor="w",
                              font=self.font_small, fill="#b9b9b0")
            if switch_lines:
                # 复选框整块（方格 + 1~2 行文字）居中在它那几行的位置上
                first = ys[len(items) - len(switch_lines)]
                self._draw_switch(row, switch_x,
                                  (first + ys[-1]) / 2, switch_lines)
            return x0 + 14, x0 + 14 + widest

        font = self.font_main
        fill = "#333333"
        chip = None
        draggable = True
        if row.block == COVER_BLOCK:
            # 「已有封面」里的东西**不允许拖动**（用户 2026-09-15）：那几段是
            # 用户自己排的封面，工具不识别、不加工，也不该被拖进任何部件。
            # 文字用灰调，跟"参与排版的内容"视觉上区分开。
            draggable = False
            fill = "#7c7c7c"
        elif row.kind in ("title", "subtitle", "entitle"):
            # 英文题目现在是独立部件了，样式上跟中文题目/副标题呼应一下
            # （加粗、同色），不算普通摘要正文（用户 2026-09-14："英文题目
            # 应该跟中文标题一样单加一个部件"）。
            font = self.font_title if row.kind == "title" else self.font_bold
            fill = C_TITLE
        elif row.kind == "toc_field":
            # 目录/清单域：Word 自动生成的，不是用户的正文——斜体灰蓝跟正文
            # 区分开，也不给拖拉（拖它没有意义，跟目录类部件拒收拖放一个理）
            font = self.font_field
            fill = C_FIELD
            draggable = False
        elif row.kind == "label":
            font = self.font_bold
            fill = C_LABEL
        elif row.kind == "heading":
            fill = C_HEADING.get(row.col, "#333")
            if row.col == "l1":
                font = self.font_bold
        elif row.kind == "caption" and row.special in ("图", "表"):
            chip = _caption_chip(row)
        if row.block and not ENABLED.get(row.block, True):
            fill = "#9a9a9a"

        # 六圆点"抓手"画在文字之前（用户 2026-09-13：放在行尾离文字太远）
        grip_fill = "#c9c9c2" if (row.block and not ENABLED.get(row.block, True)) else "#b3b3ab"
        if draggable:
            self._draw_grip(x0 + pad + 1, cy, grip_fill)
        else:
            c.create_line(x0 + pad + 1, cy, x0 + pad + GRIP_ROOM - 6, cy,
                          fill="#b9c2cb", dash=(3, 2))
        tx = x0 + pad + GRIP_ROOM
        if chip:
            chip_font = self.font_small
            cw = chip_font.measure(chip) + 12
            c.create_rectangle(tx, cy - 9, tx + cw, cy + 9, fill=C_CHIP[row.special], outline="")
            c.create_text(tx + cw / 2, cy, text=chip, font=chip_font, fill="#ffffff")
            tx += cw + 6
        # 正文不"…"截断也不折行（用户 2026-09-14）：relayout 已按单行实测把
        # 列宽撑开，超出视口的部分靠横向滚动条看。
        # 显示的文字＝`shown()`（导出时会写成什么样），量宽/绘制都必须用它：
        # relayout 里 `text_need` 量的就是同一个字符串。
        shown = self.shown(row)
        text_w = font.measure(shown)
        # 文字 + 复选框也按行排：接不到文字右边就换到下一行；文字本身还装不下
        # 就折两行（跟 `_row_height()` 同一套排法）。没有复选框的行 `ys` 只
        # 一项、就是 `cy`，行为跟以前一模一样。
        switch_lines, switch_x, _wrapped = self._switch_plan(
            row, tx + text_w + 16, x1, x0 + 4)
        ys = self._layout_lines(cy, 1 + len(switch_lines))
        c.create_text(tx, ys[0], text=shown, anchor="w", font=font, fill=fill)
        switch_y = cy
        if switch_lines:
            first = ys[len(ys) - len(switch_lines)]
            switch_y = (first + ys[-1]) / 2
        # **已有真域**的那些行后面给一个"重新为文档生成"复选框（用户 2026-09-15）：
        # 勾上就在导出时把源文件里那份域换成本工具生成的（要用户确认，见 _do_toggle）；
        # "将生成Word自动…"那种还没域的行不给——本来就没东西可覆盖。
        if row.kind == "toc_field" and row.doc_index is not None:
            bx = tx + text_w + 16
            c.create_rectangle(bx, cy - 7, bx + 14, cy + 7, fill="#ffffff",
                               outline="#8d97a1")
            if row.regenerate:
                c.create_text(bx + 7, cy, text="✓", font=self.font_small, fill=C_FIELD)
            label = "重新为文档生成"
            c.create_text(bx + 22, cy, text=label, anchor="w",
                          font=self.font_field, fill=C_FIELD)
            self.toggles.append((bx - 4, cy - 11, bx + 24 + self.font_field.measure(label),
                                 cy + 11, "regenerate", row))
        # 可拖区＝从文字所在格到区带右缘：文字越界显示的部分也能按住就拖
        if draggable:
            self.draggables.append((x0, y0, right, y1, row))
        # "自动重新编号参考文献"那个复选框（用户 2026-09-15）
        self._draw_switch(row, switch_x, switch_y, switch_lines)
        return tx, tx + text_w

    def _draw_grip(self, x: float, cy: float, fill: str) -> None:
        """六圆点"抓手"：告诉用户这一行的文字可以拖（画出来，不依赖字体字形）。"""
        c = self.canvas
        for row_i in range(3):
            for col_i in range(2):
                gx, gy = x + col_i * 6, cy - 6 + row_i * 6
                c.create_oval(gx - 2, gy - 2, gx + 2, gy + 2, fill=fill, outline="")

    def shown(self, row: Row) -> str:
        """这一行在表里显示的文字（＝它导出时会写成什么样），见 `shown_text()`。

        **量列宽和实画必须都走这里**：显示的是"改过之后的文字"，还拿原文去量
        宽，多出来的那几个字就会越出格线（【目录】比目录宽、【摘要】同理）。

        参考文献条目额外看一眼"自动重新编号参考文献"那个复选框：勾着就显示
        成品里的样子（`[1] …`——序号由 Word 自动编号画、手打的 `[1]`/`［1］`
        被剥掉）；没勾就原样显示用户自己写的序号。
        """
        text = shown_text(row, self._insert_text)
        if self.renumber_refs and row.block == "参考文献" and row.kind == "content":
            index = self._reference_index(row)
            if index is not None:
                return reference_display_text(text, index, self.reference_style)
        return text

    def _reference_index(self, row: Row) -> int | None:
        """这一行是参考文献里的第几条（1 起算）；不是条目返回 None。

        按**表里的行序**数，跟导出时 `apply_anchor_content_styles` 遍历参考文献
        区间的顺序一致（区块内的顺序被拖放规则锁着，不会跟原文顺序不一致）。
        空文字的行不算（导出那边也跳过空段：夹在条目之间的空行不编号）。
        """
        index = 0
        for r in self.rows:
            if r.block != "参考文献" or r.kind != "content" or not r.text.strip():
                continue
            if r.doc_index is None:
                continue
            index += 1
            if r is row:
                return index
        return None

    def _switch_at(self, row: Row):
        """这一行要不要挂"自动重新编号"复选框；返回 (kind, 文字, 是否勾着)。

        用户 2026-09-15："在结构页「参考文献」和图表清单处，允许自动重新编号
        参考文献（默认为开）、自动重新编号题注（默认为关）"。锚点行＝参考文献
        块的标签行（没有标签就是它的空位行）、以及三张清单里**第一张**的行
        （三个清单部件是互斥的几种形态，复选框只画一次，不重复）。部件停用时
        照样画（灰调）——不然用户没地方把它打开。
        """
        if row.col != "label" or row.kind not in ("label", "slot"):
            return None
        if row.block == "参考文献":
            return ("renumber_refs", "自动重新编号", self.renumber_refs)
        # （“自动重新编号题注”那个复选框 2026-09-16 删掉了：用户定“我们只管样式、
        #   编号与我们无关”，题注不再写域/改号，没有可切的东西。）
        return None

    def _draw_switch(self, row: Row, x: float, cy: float,
                     lines: list[str] | None = None) -> None:
        """在 `x` 处画那个"自动重新编号…"复选框（没有就什么也不做）。

        跟"重新为文档生成"同一套自绘方块：方格 + 勾 + 说明文字，点击区记进
        `self.toggles`（交互在 `_do_toggle` 里接）。用 9pt 小字，跟正文区分开。

        `lines`：文字已经折好的行（见 `_switch_plan`）；多行时方格对齐第一行，
        整块以 `cy` 为中心。不传就单行。
        """
        info = self._switch_at(row)
        if info is None:
            return
        kind, text, checked = info
        rows_text = lines or [text]
        pitch = _SWITCH_LINE_PITCH
        top = cy - (len(rows_text) - 1) * pitch / 2
        c = self.canvas
        c.create_rectangle(x, top - 7, x + 14, top + 7, fill="#ffffff", outline="#8d97a1")
        if checked:
            c.create_text(x + 7, top, text="✓", font=self.font_small, fill=C_FIELD)
        widest = 0.0
        for k, line in enumerate(rows_text):
            c.create_text(x + 22, top + k * pitch, text=line, anchor="w",
                          font=self.font_small, fill=C_FIELD)
            widest = max(widest, self.font_small.measure(line))
        self.toggles.append((x - 4, top - 11,
                             x + 24 + widest,
                             top + (len(rows_text) - 1) * pitch + 11, kind, row))

    def _switch_plan(self, row: Row, same_line_x: float, cell_right: float,
                     below_x: float) -> tuple[list[str], float, bool]:
        """复选框怎么摆：返回（要画的文字行、画在哪个 x、是不是换到了下一行）。

        * 接在文字右边装得下 ⇒ 就接在那儿（1 行）；
        * 装不下 ⇒ 换到下一行（`below_x`）；
        * 换到下一行仍然装不下（"自动重新编号参考文献"比标签列还宽，
          高 DPI 下一定会发生）⇒ 文字**折成两行**。

        传相对坐标也行（只有差值参与比较）——`_row_height()` 就是这么用的：
        两边共用这一份算法，所以“画出来什么布局”和“行高够不够”永远一致。
        """
        info = self._switch_at(row)
        if info is None:
            return [], 0.0, False
        text = info[1]
        if self._switch_width(row) <= cell_right - 6 - same_line_x:
            return [text], same_line_x, False
        avail = max(24.0, cell_right - 6 - below_x - 22)
        lines: list[str] = []
        rest = text
        while rest and len(lines) < 2:
            take = ""
            for i in range(1, len(rest) + 1):
                if self.font_small.measure(rest[:i]) <= avail:
                    take = rest[:i]
                else:
                    break
            if not take:
                take = rest[:1]                  # 一个字都放不下也要前进
            lines.append(take)
            rest = rest[len(take):]
        if rest:
            lines[-1] += rest                    # 折不下就并回第二行（宁可略溢出）
        return lines, below_x, True

    @staticmethod
    def _layout_lines(cy: float, n: int) -> list[float]:
        """把 n 行内容竖直居中排在 cy 上下，返回每行的中心 y。

        间距 `_SWITCH_LINE_PITCH`（19px）＝ 9 号字实测行高 17 加一点余量：
        两行 ÷2 居中、三行、四行都靠它统一排，不再各写一套 ±10/±20。
        """
        if n <= 0:
            return []
        top = cy - (n - 1) * _SWITCH_LINE_PITCH / 2
        return [top + k * _SWITCH_LINE_PITCH for k in range(n)]

    def _switch_width(self, row: Row) -> float:
        """这一行那个"自动重新编号…"复选框占的总宽（方格 + 文字；没有就 0）。"""
        info = self._switch_at(row)
        return 0.0 if info is None else 22 + self.font_small.measure(info[1])

    def _row_height(self, row: Row, cell_w: float) -> float:
        """这一行该多高。

        用户 2026-09-17：“这个超出了虚线框。可以换行处理（但注意高度不要超）”
        ——「自动重新编号参考文献」那个复选框在标签列里放不下，就换到提示词/标签
        **下面一行**；换行后仍然放不下（高 DPI 下必然）就把文字折成两行。
        排法与 `_draw_cell_content` **共用 `_switch_plan()`**，所以高度永远够。
        空位行基础 56（两行）、标签行 28；每多一行加 `_SWITCH_LINE_PITCH`。
        """
        if row.kind == "slot":
            hint_w = self._slot_font(row).measure(self._slot_hint(row))
            note = self._slot_note(row)
            note_w = self.font_small.measure(note) if note else 0.0
            lines, _x, wrapped = self._switch_plan(
                row, 14 + max(hint_w, note_w) + 20, cell_w, 4)
            if not wrapped:
                return 56
            # 提示词（1 行）+ 说明（有就 1 行）+ 复选框（1~2 行）。三行那种用
            # 76（跟以前一模一样，别动这个数）；复选框折成两行再多一档。
            return (76 if note else 56) + _SWITCH_LINE_PITCH * (len(lines) - 1)
        if row.col == "label" and row.kind == "label":
            text_w = self.font_bold.measure(self.shown(row))
            lines, _x, wrapped = self._switch_plan(
                row, 9 + GRIP_ROOM + text_w + 16, cell_w, 4)
            if not wrapped:
                return 28
            return 46 + _SWITCH_LINE_PITCH * (len(lines) - 1)
        return 28

    def set_insert_text(self, insert_text: dict[str, str] | None) -> None:
        """预设「插入文字」换了（用户重新应用了预设）：重铺一遍再画。

        这张表只影响**显示**，不改任何行数据（导出那边始终从预设实时取值），
        所以重铺一次就够；空白模板骨架（没有行文字）重铺也无害。
        """
        new = dict(insert_text) if insert_text else {}
        if new == self._insert_text:
            return
        self._insert_text = new
        self.relayout()
        self.draw()

    def _slot_font(self, row: Row):
        """空位提示词用哪个字体（relayout 量宽与实画共用一份，别各写一套）。

        目录/三张清单（部件启用时）那几条提示是"将生成Word自动…"，用跟域行
        同一套斜体；其余空位照旧用 `font_small`。
        """
        key = row.slot_for or ""
        if key in _AUTO_GENERATED_PARTS and ENABLED.get(key, True):
            return self.font_field
        return self.font_small

    def _slot_hint(self, row: Row) -> str:
        """空位行的**第一行**提示文字。

        relayout（量列宽）和绘制两处共用同一份逻辑，避免两边各写一套、量宽和
        实画对不上。

        用户 2026-09-14 定稿、2026-09-15 改词，**只有这几种写法**：
        * 部件未启用（两种模式一样）："未启用"；
        * 空白模板模式（还没导入文件）："将生成占位字符"。这时候没有文字
          可拖，写"可拖放文字到此"是骗人；
        * 目录 / 三张图表清单：内容由 Word 域自动生成，写成
          "将生成Word自动目录"、"将生成Word自动插图清单"……（用户 2026-09-15：
          "勾选插图附表清单后，显示将生成Word自动……"）。这几个部件的空位
          **不允许拖放**，见 `_check_drop`；
        * 其余（导入后、启用但空着）："可拖放文字到此"（第二行见 `_slot_note`）。
        """
        key = row.slot_for or ""
        if key and not ENABLED.get(key, True):
            return "未启用"
        if key in _AUTO_GENERATED_PARTS:
            # 目录/清单两种模式下都是"生成 Word 域"，所以这条排在空白模板模式
            # 那条前面——说"将生成占位字符"对它们是错的（用户 2026-09-14：
            # "三张清单的空位不是占位的，是真的"）。
            return f"将生成Word自动{key}"
        if self.blank_mode:
            return "将生成占位字符"
        return "可拖放文字到此"

    def _slot_note(self, row: Row) -> str | None:
        """空位行的**第二行**（换行显示）——只有一种情形有，其余返回 None。

        导入文件后 + 部件启用 + 这格还空着：留白在生成时会被补上占位文字，
        所以写"留白将生成占位字符"。空白模板模式（本来就是在生成占位字符）
        和未启用（压根不会生成）都没有第二行。
        """
        key = row.slot_for or ""
        if key and not ENABLED.get(key, True):
            return None
        if self.blank_mode or key in _AUTO_GENERATED_PARTS:
            return None
        return "留白将生成占位字符"

    def _on_yscroll(self, first: str, last: str) -> None:
        """canvas 的 yscrollcommand：喂纵向滚动条，同时把列头钉在视口顶。"""
        self._vsb.set(first, last)
        self._overlay_header()

    def _overlay_header(self) -> None:
        """冻结表头（用户 2026-09-14）：纵向滚动后，把"视口顶所在区"的列头
        重画一份钉在视口顶边。副本用文档 x 坐标——横向滚动时跟内容一起移动，
        永远和列位对齐。

        ⚠️ 这个函数挂在 canvas 的 `yscrollcommand` 上，**每滚一下都要跑**。
        原来每次都 `delete("hdr_overlay")` 再重建 5~8 个图元，列头那一条因此
        在滚动时一直闪（用户 2026-09-16 反馈的"滚动时色块"）。改成：同一个
        区的副本**只挪位置**（`canvas.move`，C 实现），只有跨了区才重建。
        """
        c = self.canvas
        if not self.laid:
            return
        top = c.canvasy(0)
        cur_zone = None
        for item in self.laid:
            if item["kind"] == "zone" and item["y0"] <= top:
                cur_zone = item["zone"]
        hdr = None
        if cur_zone is not None:
            hdr = next((it for it in self.laid
                        if it["kind"] == "colheader" and it["zone"] == cur_zone), None)
        if hdr is None or top - hdr["y0"] < 2:
            # 列头本体就贴在视口顶（或者上面根本没有区）：不需要副本
            if self._hdr_key is not None:
                c.delete("hdr_overlay")
                self._hdr_key = None
            return
        height = hdr["y1"] - hdr["y0"]
        if self._hdr_key is not None and self._hdr_key[0] == cur_zone:
            dy = top - self._hdr_key[1]
            if dy:
                c.move("hdr_overlay", 0, dy)
            self._hdr_key = (cur_zone, top, height)
            return
        c.delete("hdr_overlay")
        self._hdr_key = (cur_zone, top, height)
        self._draw_colheader(hdr, y_top=top, tag="hdr_overlay")

    def _overlay_covers(self, cy: float) -> bool:
        """这个画布纵坐标是不是落在"冻结表头"那份浮动副本上——它是另外画出来的
        图元，跟真正内容格不在同一位置，鼠标点击判定（`on_press`/`_hit_cell`/
        `toggles`/`draggables`/`block_handles`）如果只按内容原来的几何位置算，
        会在这里误判成点中了副本底下被盖住的那一行（误勾开关、误拖文字，
        2026-09-22 审计揪出的真 bug）。凡是要处理点击坐标的地方，先问一句
        这个，命中就当"点在表头上"处理（不往下派发）。"""
        if self._hdr_key is None:
            return False
        _, y0, height = self._hdr_key
        return y0 <= cy < y0 + height

    def on_hwheel(self, e) -> None:
        """Shift+滚轮＝横向滚动。"""
        self.canvas.xview_scroll(-2 * wheel_notch(e), "units")

    # ------------------------------------------------------------ 交互

    def _hit_cell(self, cx: float, cy: float):
        for cell in self.cells:
            x0, y0, x1, y1, _row, _key = cell
            if x0 <= cx <= x1 and y0 <= cy <= y1:
                return cell
        # 落在段号/结构/多选这几列（不属于任何内容格）时，按"这一整行"算落点
        # （2026-09-19 修：抓手就在结构列里，用户很自然地把它拖到另一个部件
        # 的结构列上松手，以前这种落点直接被当成"没落在任何地方"，毫无反应）。
        for item in self.laid:
            if item["kind"] == "row" and item["y0"] <= cy <= item["y1"]:
                row = item["row"]
                return (0, item["y0"], self.content_w, item["y1"], row, row.col)
        return None

    def on_press(self, e) -> None:
        """2026-09-19 改：多选不再靠 ctrl/shift 点选（用户原话"多选交互应该
        是左侧复选框"）——左侧复选框（见 `_draw_row`/`_toggle_row_select`）
        是**唯一**的多选入口，只支持连续 N 行。这里普通点击只管"开始拖"：
        点在已经勾选的行上＝保留整组一起拖，否则清掉勾选只拖这一行。
        """
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        if self._overlay_covers(cy):
            return          # 点在"冻结表头"浮动副本上，不派发给底下被它盖住的行
        for x0, y0, x1, y1, row in self.block_handles:
            if x0 <= cx <= x1 and y0 <= cy <= y1:
                self._selected.clear()
                self._drag = {"row": row, "start": (e.x, e.y), "active": False,
                              "whole_block": True}
                return
        for x0, y0, x1, y1, kind, payload in self.toggles:
            if x0 <= cx <= x1 and y0 <= cy <= y1:
                self._do_toggle(kind, payload)
                return
        for x0, y0, x1, y1, row in self.draggables:
            if x0 <= cx <= x1 and y0 <= cy <= y1:
                if id(row) not in self._selected or len(self._selected) <= 1:
                    self._selected = {id(row)}
                    self.draw()
                self._drag = {"row": row, "start": (e.x, e.y), "active": False}
                return
        self._drag = None
        self._clear_selection()
        # 点到空位/其他格子：只清掉选择，**不写任何"这一格是什么"的提示**
        # （用户 2026-09-15：选中一条就在下面显示这条的信息，多余）。

    def on_right_click(self, e) -> None:
        """右键＝在任意位置插入空位（用户 2026-09-14 要求"允许在所有地方
        加行"，不再局限于摘要/参考文献这类原本就有"＋行"按钮的块）。

        空白模板模式没有意义：没有文档、没有段落，插出来的空位永远不会有
        东西可以拖进去，只是摆设。
        """
        if self.blank_mode:
            return
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        if self._overlay_covers(cy):
            return
        for x0, y0, x1, y1, row in self.draggables:
            if x0 <= cx <= x1 and y0 <= cy <= y1:
                self._show_insert_menu(e, after_row=row)
                return
        for item in self.laid:
            if item["kind"] in ("zone", "colheader") and item["y0"] <= cy <= item["y1"]:
                self._show_insert_menu(e, zone=item["zone"])
                return

    def _show_insert_menu(self, e, after_row: Row | None = None, zone: str | None = None) -> None:
        menu = tk.Menu(self.root, tearoff=0)
        if after_row is not None:
            menu.add_command(label=f"在第 {after_row.doc_index} 段后插入空位",
                             command=lambda: self._insert_blank_after(after_row))
        else:
            menu.add_command(label=f"在「{ZONE_TEXT[zone]}」最前面插入空位",
                             command=lambda: self._insert_blank_at_zone_top(zone))
        menu.tk_popup(e.x_root, e.y_root)

    def _insert_blank_after(self, row: Row) -> None:
        # 一律插"正文位"（col="text"），不跟着 row 自己的列走——一个块正常
        # 只有一个标签，插标签位没什么用；真要腾标签位，走"三步走"：这里插
        # 一个正文位，把占用标签格的东西挪过来，标签格自然就空出来了。
        block = None if row.zone == "body" else row.block
        col = "text"
        self._push_undo("插入空位")
        blank = Row(rid=0, kind="slot", col=col, block=block, slot_for=block,
                   slot_col=col, zone=row.zone, added=True)
        self.rows.insert(self.rows.index(row) + 1, blank)
        self._renumber_rid()
        annotate_blocks(self.rows)
        self.relayout()
        self.draw()

    def _insert_blank_at_zone_top(self, zone: str) -> None:
        # 只有正文区能用这一条。前/后区的空位必须归属某个部件（题目、摘要、
        # 参考文献……）才谈得上"身份"，而"区最前面"那个位置上面什么都没有，
        # 造出来只能是没有 slot_for 的无主空位——认领后得到 kind="content"、
        # block=None 的行，build_manual_structure 里没有分支接它，等于白拖；
        # 更糟的是它是**整体取代**自动识别结果的，被拖进来那一段原有的身份会
        # 被一起抹掉（2026-09-14 实测：把「题目」拖进这种空位，TitleMapping
        # 直接变 None，导出时题目整块消失）。
        # 前/后区要加空位，请在某一行上右键——那个空位会跟着该行的部件走。
        if zone != "body":
            return
        insert_at = 0
        for i, r in enumerate(self.rows):
            if r.kind == "zone" and r.zone == zone:
                insert_at = i + 1
                break
        self._push_undo("插入空位")
        blank = Row(rid=0, kind="slot", col="text", block=None, slot_for=None,
                   slot_col="text", zone=zone, added=True)
        self.rows.insert(insert_at, blank)
        self._renumber_rid()
        annotate_blocks(self.rows)
        self.relayout()
        self.draw()

    def _renumber_rid(self) -> None:
        for i, r in enumerate(self.rows):
            r.rid = i + 1

    # ------------------------------------------------------------ 多选

    def _toggle_row_select(self, row: Row) -> None:
        """左侧复选框（2026-09-19 新增，替代 ctrl/shift 点选——用户原话
        "多选交互应该是左侧复选框"）：只支持**连续 N 行**，不支持跳着勾
        （用户原话"对勾选的连续N行（不支持多选不连续的行）"）。

        勾选一行：跟当前已勾的范围紧挨着（正好在最前一行前面/最后一行
        后面）就并进去扩大范围；离得远就当成"另起一段"，清空重选只留
        这一行——不強行拼接，免得用户搞不清"这两段怎么突然连起来了"。
        取消勾选：在范围两端就直接缩一格；勾掉中间一行会把连续区间拆成
        两段，这种情况直接清空，逼用户重新勾一遍连续的。
        """
        order = [r for _x0, _y0, _x1, _y1, r in self.draggables]
        if row not in order:
            return
        idx = order.index(row)
        if id(row) in self._selected:
            sel_idx = sorted(order.index(r) for r in order if id(r) in self._selected)
            if idx in (sel_idx[0], sel_idx[-1]):
                self._selected.discard(id(row))
            else:
                self._selected.clear()
            self.draw()
            return
        if not self._selected:
            self._selected = {id(row)}
        else:
            sel_idx = sorted(order.index(r) for r in order if id(r) in self._selected)
            lo, hi = sel_idx[0], sel_idx[-1]
            self._selected = ({*self._selected, id(row)} if idx in (lo - 1, hi + 1)
                              else {id(row)})
        self.draw()

    def _clear_selection(self) -> None:
        if self._selected:
            self._selected.clear()
            self.draw()

    def on_motion(self, e) -> None:
        d = self._drag
        if d is None:
            return
        if not d["active"]:
            if abs(e.x - d["start"][0]) + abs(e.y - d["start"][1]) < 6:
                return
            d["active"] = True
            self._begin_drag(d["row"])
        self._update_drag(e)

    def _can_drop_block(self, dst_row: Row | None) -> bool:
        """抓手拖动整块（`self._drag_rows` 是这个部件当前的整段）能不能落在
        `dst_row` 上：**落在哪一行都行**，只要不是落在它自己身上
        （2026-09-19 用户口径"无条件的允许调换"；以前只认"落在另一个具名
        部件上"，落在正文、空位、结构列这些地方都是没反应）。"""
        if dst_row is None:
            return False
        return id(dst_row) not in {id(r) for r in self._drag_rows}

    def _apply_block_drop(self, dst_row: Row) -> None:
        """整块落到 `dst_row`：落点在另一个部件的范围里（前/后区各个部件）
        ＝两整块互换位置；落点在正文区、空位或没有归属的行上＝整块插到
        那一行的前面。"""
        if not self._can_drop_block(dst_row):
            return
        block_rows = list(self._drag_rows)
        dst_i = self.rows.index(dst_row)
        name = block_rows[0].block or "这一块"
        # 落点属于哪个部件就按那个部件的**整块**来（含它的空位），不能落在
        # 一个部件的标签位和正文位之间把它拆开。
        dst_span = None
        if dst_row.block and dst_row.zone != "body":
            sp = _find_block_span(self.rows, dst_row.block)
            if sp is not None and sp[0] <= dst_i < sp[1]:
                dst_span = sp
        if dst_span is not None:
            s, e = dst_span
            seg_b = self.rows[s:e]
            if any(id(r) in {id(x) for x in block_rows} for r in seg_b):
                return
            self._apply_swap_rows(block_rows, seg_b)
            return
        self._push_undo(f"移动「{name}」")
        ids = {id(r) for r in block_rows}
        remaining = [r for r in self.rows if id(r) not in ids]
        at = remaining.index(dst_row)
        self.rows[:] = remaining[:at] + block_rows + remaining[at:]
        self._clear_selection()
        annotate_blocks(self.rows)
        self.relayout()
        self.draw()

    def _begin_drag(self, row: Row) -> None:
        c = self.canvas
        self._ghost = row
        d = self._drag
        # 整块移动只有两条路：☰ 抓手，或者用户自己勾选连续几行。直接拖标签行
        # 只挪标签本身——标签得能单独调整（用户 2026-09-19 明确：不要做
        # "拖标签＝拖整块"，那样标签就没法单独调了）。
        if d is not None and d.get("whole_block") and row.block:
            span = _find_block_span(self.rows, row.block)
            self._drag_rows = self.rows[span[0]:span[1]] if span is not None else [row]
            self._drag_is_block = True
        elif id(row) in self._selected and len(self._selected) > 1:
            by_id = {id(r): r for r in self.rows}
            self._drag_rows = sorted(
                (by_id[i] for i in self._selected if i in by_id),
                key=lambda r: (r.doc_index is None, r.doc_index))
            self._drag_is_block = False
        else:
            self._drag_rows = [row]
            self._drag_is_block = False
        multi = len(self._drag_rows) > 1

        # 拖动幽灵标签的预览**保留截断**（表格本体已全面去"…"，这里是跟
        # 鼠标的浮动提示：拖一段几百字的正文时若显示全文，鼠标旁会跟一条
        # 几千像素宽的横幅）。_fit 已删，这里内联一份仅此一处使用。
        def _ghost_fit(text: str, width: float) -> str:
            if not text or self.font_main.measure(text) <= width:
                return text
            lo, hi = 0, len(text)
            while lo < hi:
                mid = (lo + hi + 1) // 2
                if self.font_main.measure(text[:mid] + "…") <= width:
                    lo = mid
                else:
                    hi = mid - 1
            return text[:lo] + "…"

        label = (f"（{len(self._drag_rows)} 行）" + _ghost_fit(self.shown(row), 260)
                 if multi else _ghost_fit(self.shown(row), 340))
        self._ghost_rect = c.create_rectangle(0, 0, 10, 10, fill=C_GHOST_BG,
                                              outline=C_GHOST_LINE, tags="ghost")
        self._ghost_text = c.create_text(0, 0, text=label, anchor="w", fill="#333", tags="ghost")
        for x0, y0, x1, y1, r, key in self.cells:
            if not multi and r is row and key != row.col:
                c.create_rectangle(x0 + 1, y0 + 1, x1 - 1, y1 - 1, outline="#7fbf7f",
                                   dash=(4, 3), tags="hl")
            elif r.kind == "slot":
                if multi:
                    ok, _reason, _slots = self._check_drop_multi(self._drag_rows, r)
                else:
                    ok, _reason, _action = self._check_drop(row, r, r.col)
                if ok:
                    c.create_rectangle(x0 + 1, y0 + 1, x1 - 1, y1 - 1, outline="#7fbf7f",
                                       dash=(4, 3), tags="hl")
        c.configure(cursor="hand2")
        self._drag_last_xy = self._drag["start"]
        self._schedule_autoscroll()

    def _check_swap_rows(self, dragged_rows: list, target: Row):
        """普通拖拽（不是抓手拖整块）落在一个**已经有内容**的格子上：把拖走
        的这几行跟"从目标行开始、数量相同"的那几行互换位置——不看部件名、
        不要求两边是同一个块，纯按当前排列位置换（2026-09-19 用户口径：
        "允许用户随意调换两段之间的位置……既可以拖动一段文字和相邻的调换
        位置，也可以和远隔重洋的另一端调换位置"）。返回 `(能不能, 要换的
        那一段行)`。
        """
        if target.kind == "slot" or not target.text.strip():
            return False, None
        dragged_ids = {id(r) for r in dragged_rows}
        if id(target) in dragged_ids:
            return False, None
        try:
            t0 = self.rows.index(target)
        except ValueError:
            return False, None
        seg_b = self.rows[t0:t0 + len(dragged_rows)]
        if len(seg_b) < len(dragged_rows) or any(id(r) in dragged_ids for r in seg_b):
            return False, None
        return True, seg_b

    def _apply_swap_rows(self, dragged_rows: list, seg_b: list) -> None:
        try:
            a0 = self.rows.index(dragged_rows[0])
            b0 = self.rows.index(seg_b[0])
        except (ValueError, IndexError):
            return
        a1 = a0 + len(dragged_rows)
        if self.rows[a0:a1] != dragged_rows:
            return
        self._push_undo("调换位置")
        seg_a = dragged_rows
        if a0 > b0:
            a0, a1, seg_a, b0, seg_b = b0, b0 + len(seg_b), seg_b, a0, seg_a
        b1 = b0 + len(seg_b)
        middle = self.rows[a1:b0]
        self.rows[a0:b1] = seg_b + middle + seg_a
        self._clear_selection()
        annotate_blocks(self.rows)
        self.relayout()
        self.draw()

    def _drop_check(self, target: Row, col: str) -> bool:
        """落点合不合法——拖走一段文字就是立马拖走，不分"悬停多久"或"用户
        当初想干嘛"。抓手拖的整块（`self._drag_is_block`）只走"跟另一个
        部件整块互换"（`_can_swap`）；普通拖拽（单行或复选框多选的连续
        N 行）先试"落到空位＝认领"（`_check_drop`/`_check_drop_multi`），
        那条路走不通、目标格已经有内容，就走"跟目标这几行互换位置"
        （`_check_swap_rows`，不要求同一个部件，见该函数说明）。
        """
        if self._drag_is_block:
            return self._can_drop_block(target)
        if len(self._drag_rows) > 1:
            ok, _reason, _slots = self._check_drop_multi(self._drag_rows, target)
        else:
            ok, _reason, _action = self._check_drop(self._ghost, target, col)
        if ok:
            return True
        ok2, _seg_b = self._check_swap_rows(self._drag_rows, target)
        return ok2

    def _update_drag(self, e) -> None:
        self._drag_last_xy = (e.x, e.y)
        self._update_drag_visuals(e.x, e.y)

    def _update_drag_visuals(self, ex: float, ey: float) -> None:
        """按**窗口相对坐标**（不是画布坐标——画布坐标会随滚动改变，窗口
        相对坐标在纯滚动时不变）刷新跟手的幽灵标签和落点高亮。拆成这个
        独立方法是因为自动滚动（`_autoscroll_tick`）要在鼠标没有真的移动
        时也能刷新一遍——不然内容滚过去了，高亮框却还停在滚动前的位置。
        """
        c = self.canvas
        cx, cy = c.canvasx(ex), c.canvasy(ey)
        w = 340
        c.coords(self._ghost_rect, cx + 14, cy - 14, cx + 14 + w, cy + 12)
        c.coords(self._ghost_text, cx + 22, cy - 1)
        c.delete("hover")
        cell = self._hit_cell(cx, cy)
        if cell is None:
            return
        x0, y0, x1, y1, r, key = cell
        ok = self._drop_check(r, key)
        color = "#2e8b57" if ok else "#c0392b"
        c.create_rectangle(x0 + 1, y0 + 1, x1 - 1, y1 - 1, outline=color, width=2, tags="hover")

    def _schedule_autoscroll(self) -> None:
        self._autoscroll_job = self.root.after(60, self._autoscroll_tick)

    def _autoscroll_tick(self) -> None:
        """拖动中靠近视口顶/底边缘时自动滚动（2026-09-19 修：以前完全没有
        这个机制——源和目标不能同时出现在当前视口里时，物理上够不着，拖
        到边缘就卡住不动，用户反馈"拖动任何一个条目都不能跨其他条目放到
        别的位置"，根子就在这——不是调换/认领的判定逻辑坏了，是压根没法
        把目标带进视口）。定时器每 60ms 走一次，只要拖动还在继续、鼠标还
        停在边缘附近就一直滚；鼠标离开边缘或拖动结束就自然停手（不再自我
        续订定时器）。
        """
        self._autoscroll_job = None
        d = self._drag
        if d is None or not d.get("active"):
            return
        c = self.canvas
        h = c.winfo_height()
        margin = 36
        ex, ey = self._drag_last_xy
        if 0 <= ex <= c.winfo_width():
            if ey < margin:
                c.yview_scroll(-1, "units")
            elif ey > h - margin:
                c.yview_scroll(1, "units")
        self._update_drag_visuals(ex, ey)
        self._schedule_autoscroll()

    def on_release(self, e) -> None:
        if self._autoscroll_job is not None:
            self.root.after_cancel(self._autoscroll_job)
            self._autoscroll_job = None
        d = self._drag
        self._drag = None
        if d is None:
            return
        if not d.get("active"):
            # 点一下（没拖动）＝选中这一行。
            # **不再把这一行的信息写进状态栏**：用户 2026-09-15——"我在结构栏
            # 每选中一条，下面就会显示这条的信息，这个删去，多余"；那条带"空位：
            # …把别处的文字拖到这里即可"的提示也在被删之列。
            return
        c = self.canvas
        c.delete("ghost")
        c.delete("hl")
        c.delete("hover")
        c.configure(cursor="")
        cx, cy = c.canvasx(e.x), c.canvasy(e.y)
        cell = self._hit_cell(cx, cy)
        if cell is None:
            return
        x0, y0, x1, y1, r, key = cell
        # 拖走一段文字就是立马拖走：不分悬停多久（判定逻辑跟 `_drop_check`
        # 必须一致，见该函数说明）。
        if self._drag_is_block:
            if self._can_drop_block(r):
                self._apply_block_drop(r)
            return
        if len(self._drag_rows) > 1:
            ok, reason, slots = self._check_drop_multi(self._drag_rows, r)
            if ok:
                self._apply_drop_multi(self._drag_rows, slots)
                return
        else:
            ok, reason, action = self._check_drop(d["row"], r, key)
            if ok:
                self._apply_drop(d["row"], r, action)
                return
        ok2, seg_b = self._check_swap_rows(self._drag_rows, r)
        if ok2:
            self._apply_swap_rows(self._drag_rows, seg_b)

    @staticmethod
    def _slot_claimable(slot: Row) -> bool:
        """这个空位认领之后能不能得到一个"明确的身份"。

        正文区不靠 block 定身份（列即身份：标题列＝标题、正文列＝正文），所以
        没有 slot_for 也没关系；前/后区的身份必须来自 block，没有 slot_for 的
        无主空位认领后只能得到 kind="content"、block=None 的行——
        build_manual_structure 里没有分支接它，等于白拖，而且它会连被拖进来那
        一段原有的身份一起抹掉（2026-09-14 实测：把「题目」拖进这种空位，
        TitleMapping 直接变 None，导出时题目整块消失）。
        """
        return bool(slot.slot_for) or slot.zone == "body"

    def _check_drop(self, dragged: Row, target: Row, col: str):
        """能不能把 `dragged` 放到 `target` 的 `col` 列——返回 (能不能, 原因, 动作)。

        `原因` 现在**不显示给用户**（2026-09-15 结构页撤掉状态栏之后就没出口了）：
        拒收的表现就是"拖过去没反应"。留的几条是给将来做悬浮提示备用；
        2026-09-19 用户核定删掉了两条已与现行行为不符的（"格子已有内容"该落点
        实际被 `_check_swap_rows` 的对调接走；"只能放到本行的其他列"落点罕见），
        归成一条空 reason 的静默拒收。
        """
        if target is dragged:
            if col == dragged.col:
                return False, "已经在这个位置", None
            return True, f"把这一段移到「{COL_NAME.get(col, col)}」列", ("recol", col)
        if target.kind == "slot":
            key = target.slot_for or ""
            if key and not ENABLED.get(key, True):
                return False, f"「{key}」未启用——先点左侧勾选框启用它", None
            if key in _AUTO_GENERATED_PARTS:
                return False, f"「{key}」的内容系自动生成，不能拖放！", None
            if not self._slot_claimable(target):
                return False, "该空位无归属部件，请删除本行后重建以拖入。", None
            if not dragged.text.strip():
                return False, "这一段没有文字", None
            if key in FULL_WIDTH:
                msg = f"把这一段放进「{key}」位"
            elif target.slot_col == "label":
                msg = f"把这一段放进「{key}」的标签位" if key else "把这一段放进这个标签位"
            else:
                msg = f"把这一段放进「{key}」的正文位" if key else "把这一段放进这个空位"
            return True, msg, ("claim", target)
        # 落到这里＝拖到"非空位的格子"上：有内容的格子实际由"对调"逻辑
        # （`_check_swap_rows`，2026-09-19）接走；旧的两条拒收文案（"格子已有
        # 内容"／"只能放到本行的其他列"）已用不到，删掉（2026-09-19 用户核定）。
        # 空 reason＝静默拒收。
        return False, "", None

    def _check_drop_multi(self, dragged_rows: list, target: Row):
        """多选拖拽：从 target 开始往后找足够数量的可用空位，按段号从小到大
        依次对应过去——落位后彼此的先后顺序必须跟原来的段号顺序一致，不能
        借多选把顺序倒过来。target 必须是空位，多选没有"换列"这个概念。
        """
        if target.kind != "slot":
            return False, "多选只能拖到虚线空位", None
        key = target.slot_for or ""
        if key and not ENABLED.get(key, True):
            return False, f"「{key}」未启用——先点左侧勾选框启用它", None
        if key in _AUTO_GENERATED_PARTS:
            return False, f"「{key}」的内容系自动生成，不能拖放！", None
        if not self._slot_claimable(target):
            return False, "这个空位没有归属的部件——拖进来会没有身份。请另选一个空位。", None
        ordered = sorted(dragged_rows, key=lambda r: (r.doc_index is None, r.doc_index))
        if any(not r.text.strip() for r in ordered):
            return False, "选中的行里有没有文字的行", None
        exclude = {id(r) for r in ordered}
        start = self.rows.index(target)
        slots: list = []
        for r in self.rows[start:]:
            if len(slots) >= len(ordered):
                break
            if id(r) in exclude or r.kind != "slot":
                continue
            skey = r.slot_for or ""
            if skey and not ENABLED.get(skey, True):
                continue
            if skey in _AUTO_GENERATED_PARTS:
                continue        # 目录/清单的空位不接受拖放，多选时也跳过
            if not self._slot_claimable(r):
                continue
            slots.append(r)
        if len(slots) < len(ordered):
            return False, "空行数量不足！", None
        return True, f"把选中的 {len(ordered)} 行依次放进从这里开始的 {len(ordered)} 个空位", slots

    def _claim_slot(self, dragged: Row, slot: Row) -> Row:
        """把 dragged 的内容搬进 slot 这个空位，原位置视情况留痕或撤空，
        返回搬过去之后落在新位置的那个 Row（dragged 这个对象本身就作废了，
        后面不要再用它）。

        身份怎么定：正文区的空位按列是不是标题列区分"标题/正文"；题目/
        副标题/英文题目这几个"整行独占、没有标签栏"的部件按块名直接映射
        成对应的 kind；其余前/后区的空位按是不是"标签位"区分"标签/内容"。

        以前这里只有 body/其它 两条分支，题目/副标题的空位落进"其它"那条，
        被一律标成 content——build_manual_structure 只认 kind=="title"/
        "subtitle"，认领进来的段落其实从没被真正记成过标题，拖了也白拖
        （2026-09-14 加英文题目这个独立部件时顺带挖出来的坑，题目/副标题
        原来就有，一并修了）。
        """
        if slot.zone == "body":
            new_kind = "heading" if slot.slot_col != "text" else "body"
        elif slot.block in _FULL_WIDTH_KIND:
            new_kind = _FULL_WIDTH_KIND[slot.block]
        else:
            new_kind = "label" if slot.slot_col == "label" else "content"
        di = self.rows.index(dragged)
        si = self.rows.index(slot)
        new_row = Row(rid=slot.rid, kind=new_kind, text=dragged.text, col=slot.slot_col,
                     doc_index=dragged.doc_index, block=slot.slot_for, zone=slot.zone,
                     special=dragged.special,
                     # 题注角色/推算序号也跟着走：一行图题被拖到别处，它仍然
                     # 是图题（角色不变），显示那块推算出来的编号也用原值。
                     caption_role=dragged.caption_role,
                     # 题注角色基准（“续”是个叠加开关，点掉以后该回哪个角色）
                     # 和自动编号那串文字都得跟着走，不然拖一下就把“原材料”弄丢。
                     caption_role_base=dragged.caption_role_base,
                     auto_number=dragged.auto_number,
                     # 拖过来的那一行自带的"原材料"跟着走：拖到别的标题级别上，
                     # 显示（和导出）就得按新级别重编编号（用户 2026-09-15）。
                     # 掉进标签位的还要补上部件名——普通正文被拖进"摘要"标签位，
                     # 导出时那一段会被写成【摘要】，显示不能还露原文。
                     heading=dragged.heading,
                     # ⚠️ 落进的新位置不是"标签位"时，`label` 必须清成 None——
                     # 以前写成 `dragged.label or (... if new_kind=="label" else
                     # None)`，`or` 短路会在新位置明明不是标签位时，只要
                     # `dragged.label` 原来非空就照样把旧标签留下来（比如一段
                     # 文字先被拖进"摘要"标签格、再被拖到别处，界面仍显示
                     # "【摘要】"，导出还多报一条"重复的部件"，2026-09-22
                     # 审计揪出的真 bug）。改成先看新位置是不是标签位，不是
                     # 就直接清空，不看 `dragged.label` 原来是什么。
                     label=(dragged.label or slot.slot_for) if new_kind == "label" else None)
        self.rows[si] = new_row
        if dragged.block:
            # 原来位置永远留一个空位——拖回去的落脚点（不再"一行消失一行"）
            self.rows[di] = Row(rid=dragged.rid, kind="slot", col=dragged.col,
                                block=dragged.block, slot_for=dragged.block,
                                slot_col=dragged.col, zone=dragged.zone)
        else:
            self.rows.remove(dragged)
        return new_row

    def _apply_drop(self, dragged: Row, target: Row, action) -> None:
        kind, value = action
        if kind == "recol":
            self._push_undo("换列")
            dragged.col = value
            if dragged.zone in ("front", "back") and dragged.kind in ("label", "content"):
                dragged.kind = "label" if value == "label" else "content"
                # 列的归属变了，"原材料"也得跟着换：变标签行就得有 `label`，
                # 变回内容行就得把 `label` 清掉。否则 shown_text() 会按旧原材料
                # 渲染（拖到标签列却还显示原文、拖回内容列却仍显示【摘要】），
                # 而导出是按列走的——两边不一致（审计 2026-09-16 报的）。
                if dragged.kind == "label":
                    dragged.label = dragged.label or dragged.block or dragged.slot_for
                else:
                    dragged.label = None
            elif dragged.zone == "body" and dragged.kind in ("heading", "body"):
                # 列即身份：拖到标题列＝标题，拖到正文列＝正文
                dragged.kind = "heading" if value != "text" else "body"
                if dragged.kind == "body":
                    # 掉回正文列就不再是标题：必须清掉 heading 原材料，否则
                    # shown_text() 仍按标题渲染（页面显示"1. xxx"、导出按正文，
                    # 编号凭空消失，审计 2026-09-16 报的）。
                    dragged.heading = None
        else:  # claim
            self._push_undo("认领空位")
            slot: Row = value
            self._claim_slot(dragged, slot)
        self._clear_selection()
        annotate_blocks(self.rows)
        self.relayout()
        self.draw()

    def _apply_drop_multi(self, dragged_rows: list, slots: list) -> None:
        self._push_undo(f"批量认领（{len(dragged_rows)} 行）")
        ordered = sorted(dragged_rows, key=lambda r: (r.doc_index is None, r.doc_index))
        for row, slot in zip(ordered, slots):
            self._claim_slot(row, slot)
        self._clear_selection()
        annotate_blocks(self.rows)
        self.relayout()
        self.draw()

    def _toggle_regenerate(self, row: Row) -> None:
        """已有真域的目录/清单行：勾"重新为文档生成"。

        用户 2026-09-15："点击后弹窗警告'将覆盖原有自动……！'"——勾之前先问
        一句（取消就不勾），勾上之后导出时会把源文件那份域删掉、换成本工具
        生成的新域（删旧域见 `cleanup.drop_regenerated_fields`）。取消勾选是
        安全方向（回到"不动原有域"），不再多问。
        """
        part = row.block or "目录"
        if not row.regenerate:
            ok = messagebox.askyesno(
                "重新生成", f"将覆盖原有自动{part}！\n\n是否继续？", parent=self.root)
            if not ok:
                self.draw()
                return
            self._push_undo(f"重新生成「{part}」")
            row.regenerate = True
        else:
            self._push_undo(f"不再重新生成「{part}」")
            row.regenerate = False
        self.draw()

    def _list_missing_kinds(self, key: str, enabling: bool) -> set[str]:
        """这张清单要收的题注种类里，文档里一个都没有的那些。

        用户 2026-09-15："不灰，可点，但点了没反应+告警。不留例外。"
        清单域是按**题注样式**收条目的（`TOC \\h \\z \\t "图题,1"`），所以
        "要哪种图/表题"看 `LIST_PART_KINDS`（跟引擎同一份表），"有没有"看
        结构页上的题注行——**判据必须跟引擎那边一模一样**（`wears_caption_style`
        ＝挂样式的才算数）：分图/分表算、**续行不算**（它不挂样式、不进清单）、
        英文对照行不算。不然会出现"结构页看着有、勾选却被拒"。
        """
        if not enabling:
            return set()
        need = LIST_PART_KINDS.get(key)
        if not need:
            return set()
        have = {row.special for row in self.rows
                if row.kind in ("caption", "body", "content")   # 手动勾的行 kind 不变
                and row.special in ("图", "表")
                and (row.caption_role or "main") != "cont"
                and (row.caption_role or "main") != "en"}
        return {kind for kind in need if kind not in have}

    def _do_toggle(self, kind: str, payload) -> None:
        if kind == "select":
            self._toggle_row_select(payload)
            return
        if kind == "regenerate":
            self._toggle_regenerate(payload)
            return
        if kind == "renumber_refs":
            # 只改显示与导出行为，不动任何行数据；也不进撤回栈（这不是"改了
            # 结构"，就是个开关，再点一下就回去了）。它会改变条目在表里的
            # 文字（[1] 前缀）⇒ 得重铺一遍列宽。
            self.renumber_refs = not self.renumber_refs
            self.relayout()
            self.draw()
            return
        if kind == "block":
            key = payload
            cur = ENABLED.get(key, True)
            # 三张清单只能选中这四种形态（用户 2026-09-15）：插图清单／附表清单／
            # 插图和附表清单／插图清单+附表清单。**合并形态与独立清单互斥**——
            # 同时开会把同一张图列两遍（所以引擎那边不再报"形态冲突"，改成
            # 这里直接不让选）。
            bad = _list_shape_conflict(key, enabling=not cur)
            if bad:
                # 用户 2026-09-15（审计核实）：原来这里直接 `return`，什么反馈
                # 都没有——用户点了勾选框却什么也没发生，只能以为"勾选框坏了"。
                # 结构页没有状态栏（`_set_status` 已删），所以用一句短弹窗，
                # 文案就是 `_list_shape_conflict()` 里那句（说清"为什么不能这样勾"）。
                messagebox.showinfo("矛盾的清单", bad, parent=self.root)
                return
            missing = self._list_missing_kinds(key, enabling=not cur)
            if missing:
                # 用户 2026-09-15：勾选框**不置灰**（"不灰，可点，但点了没反应
                # +告警。不留例外"）。清单域是按题注样式收条目的，这一类题注
                # 一个都没有的话，勾了也只会生成一张空清单——所以不改状态，
                # 弹一句告知这次勾选没生效（2026-09-19 起文案从简）。**没有任何例外**（包括"文档里连题注都没有"）。
                # 引擎那条 B9 "不存在的清单对象"告警照旧报，两条不冲突：
                # 那条管的是导出时刻的事实，这条管的是用户当下这一下操作。
                messagebox.showinfo(
                    "没有可列对象", f"{key}没有可列对象，请核实！",
                    parent=self.root)
                return
            if cur:
                holders = [r for r in self.rows if r.block == key and r.text.strip()]
                if holders:
                    self._flash_content(holders)
                    return
            self._push_undo(f"{'停用' if cur else '启用'}「{key}」")
            ENABLED[key] = not cur
        elif kind == "insert":
            self._insert_content_row(payload)
            return
        elif kind == "delete_slot":
            row = payload
            if row.kind == "slot" and row.added and row in self.rows:
                self._push_undo("删空位")
                self.rows.remove(row)
                self._renumber_rid()
                self._clear_selection()
                annotate_blocks(self.rows)
                self.relayout()
                self.draw()
            return
        else:
            row, val = payload
            self._push_undo("改动标记")
            if val == "续":
                # "续"改的是**角色**（不改种类），而且现在是个**叠加开关**（用户
                # 2026-09-17："一旦图表被勾选……这个就应该可用"）：
                #   勾上 → cont（续图/续表：仍算题注，但**不挂样式** ⇒ **不进清单**）
                #   取消 → 回到**识别时的角色**（`caption_role_base`）
                # 分开存 base 的原因：英文对照行也能点这个框了，若是"回退到 main"，
                # 点两下就把它推成挂样式的真题注、清单里凭空多一条（2026-09-16
                # 审计抓到的坑）。要真把它当正题注：取消"图/表"再勾上（下面那一支
                # 会把 base 一并置成 main，那是用户明确的"改判"）。
                if (row.caption_role or "main") == "cont":
                    row.caption_role = row.caption_role_base or "main"
                else:
                    row.caption_role = "cont"
            elif row.special == val:
                row.special = None
                row.caption_role = None      # 取消图/表身份，角色一并清掉
                row.caption_role_base = None
            else:
                row.special = val
                # 手动勾成图/表：身份就是“主图/主表”（占号、挂样式、进清单）；
                # 原来的“分/续/英文”角色只在识别出来的行上才有意义，用户
                # 一旦手动改判，就按主图/主表算（想让它不进清单，再勾上“续”）。
                row.caption_role = "main"
                row.caption_role_base = "main"
        self.draw()

    def _insert_content_row(self, key: str) -> None:
        if not any(r.block == key for r in self.rows):
            return
        self._push_undo(f"给「{key}」加空位")
        add_content_slot(self.rows, key)
        self._renumber_rid()
        annotate_blocks(self.rows)
        self.relayout()
        self.draw()

    def on_hover(self, e) -> None:
        if self._drag is not None:
            return
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        if self._overlay_covers(cy):
            self.canvas.configure(cursor="")
            return
        on_handle = any(x0 <= cx <= x1 and y0 <= cy <= y1 for x0, y0, x1, y1, _r in self.block_handles)
        hit = on_handle or any(
            x0 <= cx <= x1 and y0 <= cy <= y1 for x0, y0, x1, y1, _k, _p in self.toggles)
        if not hit:
            hit = any(x0 <= cx <= x1 and y0 <= cy <= y1 for x0, y0, x1, y1, _r in self.draggables)
        self.canvas.configure(cursor="fleur" if on_handle else ("hand2" if hit else ""))

    def on_wheel(self, e) -> None:
        self.canvas.yview_scroll(-2 * wheel_notch(e), "units")

    def on_configure(self, e) -> None:
        if abs(e.width - self._last_width) < 24:
            return
        self._last_width = e.width
        if self._resize_job is not None:
            self.root.after_cancel(self._resize_job)
        self._resize_job = self.root.after(120, self._relayout_and_draw)

    def _relayout_and_draw(self) -> None:
        self._resize_job = None
        self.relayout()
        self.draw()

    # ------------------------------------------------------------ 按钮

    def on_confirm(self) -> None:
        """「确认」——把当前表格的行序/状态定成"导出要用的那一份"（用户
        2026-09-19："GUI 不点确认，即使点导出也不认"）：在这之前拖动/换列/
        调换这些调整都只停在 `self.rows` 上，导出永远读 `confirmed_rows`，
        点了这里才真正把当前的样子搬过去、`dirty` 清掉。
        """
        self.confirmed_rows = copy.deepcopy(self.rows)
        self.confirmed_index = self.index
        self.confirmed_settings = self._settings()
        self.confirmed_enabled = dict(ENABLED)
        self.dirty = False
        self.canvas.focus_set()

    # -------------------------------------------------- 重新识别 / 封面范围

    def _confirm_source_readable(self) -> bool:
        """「重新识别」「切换标题编号方案」「设置封面范围」都要重新从磁盘读
        `self.doc_path`——如果导入之后原文件被移动、改名或删掉了，以前这里
        直接让 `Document(...)` 的异常一路冒到兜底崩溃弹窗，点"关闭"整个程序
        就退出，这次在结构页做的所有调整（还没点"确认"的）全部丢失
        （2026-09-22 审计揪出的真 bug）。改成先在这里检查一遍、不行就弹一句
        好懂的提示、**不碰任何状态**（不撤回栈、不改行序）就返回，原来的调整
        照样留着。返回 True 才应该继续往下做。
        """
        if self.doc_path is None or not self.doc_path.exists():
            messagebox.showerror(
                "找不到原文件",
                f"原文件不见了（可能被移动、改名或删除）：\n{self.doc_path}\n\n"
                "这一步需要重新读取原文件，暂时做不了。请在主窗口重新导入这份文档。",
                parent=self.root)
            return False
        return True

    def _rerecognize(self) -> None:
        """按当前的封面范围/标题编号方案从磁盘重新识别，换掉 `rows` 和
        `index`；"重置"要退回的也换成这一次的识别结果（重置＝退回**最近
        一次识别**的样子，封面范围、编号方案是用户对这份文档的设定，不跟着
        退）。**不动 `confirmed_*`**：没点确认之前导出仍按上次确认的那一对。"""
        self.rows, self.index = build_rows(self.doc_path, self.cover_end, self.heading_scheme)
        self._pristine_rows = copy.deepcopy(self.rows)
        self._pristine_index = self.index
        self._selected.clear()

    def _reload_rows(self, label: str, keep_enabled: bool, push_undo: bool = True) -> None:
        """重新从磁盘读文档、重新识别，并用新结果重铺表。

        `push_undo=False`：调用方已经自己拍过撤回点了（切换标题编号方式要在
        **改方案之前**拍，快照里才是旧方案）。

        `keep_enabled=True`（重新识别/改封面范围用）：**保留用户当前勾好的
        部件启用状态**，只把新识别到的部件"点亮"——改完原文档或者刚设完封面
        回来，多半是要接着调，把勾选清掉等于白点一遍。
        `keep_enabled=False`（重置用）：退回"刚识别完"那一刻的样子（预设 ∪
        识别到的部件）。
        """
        if self.doc_path is None:
            return
        if not self._confirm_source_readable():
            return
        if push_undo:
            self._push_undo(label)
        kept = dict(ENABLED)
        self._rerecognize()
        seed_enabled_from(self._preset_baseline, self.rows)
        if keep_enabled:
            for part, on in kept.items():
                if on and part in ENABLED:
                    ENABLED[part] = True
        annotate_blocks(self.rows)
        self.relayout()
        self.draw()

    def set_heading_scheme(self, scheme: str | None) -> None:
        """切换原文的标题编号方案，并按新方案整表重新识别。跟"重新识别"一样
        是**没确认的调整**：可以 Ctrl+Z（撤回点连同旧方案、旧索引一起存，
        见 `_push_undo`），点了"确认"才作数。"""
        if scheme == self.heading_scheme or self.doc_path is None:
            return
        if not self._confirm_source_readable():
            return
        self._push_undo("切换标题编号方式")     # 必须在改方案之前拍，快照里才是旧方案
        self.heading_scheme = scheme
        self._reload_rows("切换标题编号方式", keep_enabled=True, push_undo=False)

    def on_apply_order_fix(self) -> None:
        """「一键重排」（2026-09-19 新增，用户明确要求；原名"按预设一键重排"，因为顺序
        其实是写死的本科规范顺序、并不来自预设，2026-09-19 改名）：把前置区/
        后置区里已经存在的部件，按规范要求的顺序（`FRONT_ORDER`/`BACK_ORDER`）
        重新排一遍。识别之后**不再自动弹"是否自动排序"**（用户 2026-09-19
        要求删掉）——顺序不合规范只在导出时、且已经读入预设的情况下告警
        （见 `main_window.on_export_doc`），要重排由用户自己点这个按钮；
        不管当前顺序有没有问题都能点，已经对的就什么都不会变。
        """
        if self.blank_mode:
            return
        self._push_undo("一键重排")
        apply_order_fix(self.rows, FRONT_ROLES)
        apply_order_fix(self.rows, BACK_ROLES)
        annotate_blocks(self.rows)
        self.relayout()
        self.draw()

    def on_reidentify(self) -> None:
        """「重新识别」：用户改了**原文档**之后，重新读一遍再识别一次。

        跟「重置」的区别只有一处：**保留用户已经勾好的部件启用状态**（新识别
        到的部件仍会补上），因为改完文档回来多半是接着调。
        """
        self._reload_rows("重新识别", keep_enabled=True)

    def on_set_cover_range(self) -> None:
        """「设置封面范围」：让用户手敲封面范围（**工具不做任何自动判定**）。

        用户 2026-09-15 的原话："我们不作任何程度的封面自动识别，纯靠用户
        手敲。"——因为"第一页"是排版结果，工具不做排版，猜出来的边界迟早会
        把正文吃掉，不如直接问。
        """
        if self.doc_path is None:
            return
        if not self._confirm_source_readable():
            return
        answer = self._ask_cover_range()
        if answer == "cancel":
            return
        # 先拍快照再改（以前反了：快照里存的就是改完之后的样子，那一步的
        # Ctrl+Z 等于空操作）。
        self._push_undo("设置封面范围")
        self.cover_end, self.delete_cover = answer
        self._rerecognize()
        annotate_blocks(self.rows)
        self.relayout()
        self.draw()

    def _ask_cover_range(self):
        """弹窗让用户指定封面范围。返回 `(cover_end, delete_cover)`，取消返回
        `"cancel"`。

        弹窗布局（用户 2026-09-15 定的）：用户选的是**封面之后的第一段**——选中的
        那一段**不算封面**，它前面的（第 0 段到选中段减一）才是封面；所以选
        "第 1 段"就等于没有封面。列表**列全部段落**（用户要求"不要限制范围，
        从第一段完整显示到最后一段"），封面不可能出现在正文中间——但选到最后
        一段时 `cover_end` 就是 `n-2`，**删不掉整篇**（管线那边另有一道底线）。
        """
        from docx import Document as _Document

        doc = _Document(str(self.doc_path))
        unwrap_section_break_sdt(doc, keep_content=True)
        paras = doc.paragraphs
        # **不限制范围**：从第 1 段到末段全部列出来（用户 2026-09-15：
        # "不要限制用户选择封面段落的范围，从第一段完整显示到最后一段"），
        # 列表自己带滚动条。
        limit = len(paras)

        win = tk.Toplevel(self.root)
        win.title("设置封面范围")
        win.transient(self.root.winfo_toplevel())
        win.configure(bg=PANEL_BG)
        # 整个窗口比内容需要的更宽（用户 2026-09-15："把整个窗口调大，间距各自
        # 都调大一点，现在太臃肿局促了"）：`minsize` 定宽不限高，高度仍随内容走。
        win.minsize(DIALOG_MIN_W, 1)
        # 顶部提示调大（原来跟着 8pt 的提示字体走，太小）；"段落"二字和它前面
        # 的空行按用户 2026-09-15 的要求一并去掉——列表就在提示语下面。
        tk.Label(win, text="如果您的文档含有封面，为防止误识别，请在该界面设置封面范围。",
                 font=(FAMILY, 11), bg=PANEL_BG, fg="#333", anchor="w",
                 justify="left", wraplength=DIALOG_MIN_W - 60).pack(
            fill="x", padx=PAD, pady=(24, 18))

        box = tk.Frame(win, bg=PANEL_BG)
        box.pack(fill="x", padx=PAD, pady=(0, 16))
        scroll = tk.Scrollbar(box, orient="vertical")
        scroll.pack(side="right", fill="y")
        # height=5：**把中间的列表做矮**（用户 2026-09-15："我让你把选择窗口
        # 调小一点，意思是中间那个选择控件做小，不是整个做小。以中间部分能显示
        # 5 个待选段落为好"），不是把整个窗口做小。
        listbox = tk.Listbox(box, font=(FAMILY, 10), height=5, activestyle="none",
                             selectmode="browse", yscrollcommand=scroll.set)
        listbox.pack(side="left", fill="both", expand=True)
        scroll.configure(command=listbox.yview)
        # 只给序号和原文，不加任何提示词。
        for i in range(limit):
            text = " ".join(paras[i].text.split())
            if len(text) > DIALOG_ROW_CHARS:
                text = text[:DIALOG_ROW_CHARS] + "…"
            listbox.insert("end", f"{i + 1}　{text}")
        current = 0 if self.cover_end is None else min(self.cover_end + 1, limit - 1)
        listbox.selection_set(current)
        listbox.see(current)

        # 列表下面那一行：11pt 加粗、**靠右对齐**，下面空一行（用户 2026-09-15）。
        tk.Label(win, text="设置该段（不含该段本身）之前的内容为封面",
                 font=(FAMILY, 11, "bold"), bg=PANEL_BG, fg="#333",
                 anchor="e").pack(fill="x", padx=PAD, pady=(0, 22))

        keep_var = tk.BooleanVar(value=bool(self.delete_cover))
        # 复选框用主窗口那一套（自造方块 + ttk 布局，见 `widgets.py`）；文字再小
        # 一号，整块（方块 + 文字）**居中**（用户 2026-09-15）。
        init_checkbox_style(win, PANEL_BG)
        checkbutton(win, "导出时删除封面", variable=keep_var, font_size=9).pack(
            anchor="center")
        tk.Label(win, text="勾选以后，封面内容会被删除；若不勾选，封面将会保留。"
                           "自动封面如果启用将在最前方插入，不论是否保留原有封面。",
                 font=self.font_hint, bg=PANEL_BG, fg="#8a8a8a",
                 justify="center", wraplength=DIALOG_MIN_W - 60).pack(
            fill="x", padx=PAD, pady=(10, 0))

        result: dict[str, object] = {}

        def ok() -> None:
            sel = listbox.curselection()
            index = sel[0] if sel else 0
            result["value"] = (None if index == 0 else index - 1, bool(keep_var.get()))
            win.destroy()

        def cancel() -> None:
            result["value"] = "cancel"
            win.destroy()

        row = tk.Frame(win, bg=PANEL_BG)
        row.pack(fill="x", pady=(26, 22))
        tk.Button(row, text="取消", font=(FAMILY, 10), padx=14, pady=4,
                  command=cancel).pack(side="right", padx=(10, PAD))
        tk.Button(row, text="确定", font=(FAMILY, 10), padx=14, pady=4,
                  command=ok).pack(side="right")
        listbox.bind("<Double-Button-1>", lambda _e: ok())
        win.grab_set()
        # 二级窗口一律摆在主窗口正中央（用户 2026-09-15）。
        center_on(win, self.root.winfo_toplevel())
        self.root.wait_window(win)
        return result.get("value", "cancel")

    def on_open_rules(self) -> None:
        """「查阅识别规则」：跟关于页「查阅本工具操作逻辑说明」**同一个窗口**。"""
        if self._open_rules_cb is not None:
            self._open_rules_cb()

    def on_reset(self) -> None:
        """重置：撤销手动修改，退回**识别结果**（不读盘、不重跑识别）。

        用户 2026-09-16："既然已经有重新识别了，重置不重新识别。" 所以这里
        恢复的是构造窗口时那一份行数据（`_pristine_rows`），而不是重新
        `build_rows()`——原先那种写法会把 Word 里后来改动的东西也一并读进来，
        跟旁边的「重新识别」按钮重复。
        封面范围不跟着清：它是用户对**这份文档**的设定（属于输入），不是
        "调整过的识别结果"。
        """
        self._push_undo("重置")
        self.rows = copy.deepcopy(self._pristine_rows)
        self.index = self._pristine_index
        self._selected.clear()
        # 退回"刚识别完"的状态，部件启用也要跟着退回那一刻该有的样子
        # （预设∪识别到的部件），不然重置了行、启用状态却还停在调整过的样子。
        # 空白模板模式没有文档可识别，"识别到的部件"永远是空集，退回的就是
        # 纯预设的样子——这跟正常模式是同一条规则，不用另外分支。
        seed_enabled_from(self._preset_baseline, self.rows)
        self.relayout()
        self.draw()

    # ------------------------------------------------------------ 撤回

    def _settings(self) -> dict:
        """"这份文档的设定"——跟行数据分开、但同样要能撤回的那几项。

        以前快照只存 rows + ENABLED，所以 `cover_end`/`delete_cover`/两个
        "自动重新编号"开关**撤不回去**（2026-09-15 审计核实为真）；而且
        `on_set_cover_range` 是先改后拍快照，那一步的 Ctrl+Z 干脆是个空操作。
        """
        return {
            "cover_end": self.cover_end,
            "delete_cover": self.delete_cover,
            "renumber_refs": self.renumber_refs,
            "heading_scheme": self.heading_scheme,
        }

    def _push_undo(self, label: str) -> None:
        """记一个撤回点（整表快照 + 这份文档的设定——几十行的小对象，随便存）。

        这里是所有"要改 `self.rows`"的动作共用的入口（拖动/换列/调换/
        增删空位/启用开关/重置/重新识别/设置封面范围/切换标题编号方式……），
        顺手在这标脏：改了就是改了，没点"确认"之前导出不该悄悄生效。

        ⚠️ **识别索引也在快照里**（`index`，以及"重置"的基线 `pristine_*`）：
        行里的 `doc_index` 指的是索引里的段号，重新识别会换掉整份索引，只
        撤回行不撤回索引的话，撤回之后行和索引就对不上了。索引在拖动这类
        编辑里不会被改（导出时才被 `build_ops` 原地改写），所以存引用即可，
        不用深拷贝；`pristine_rows` 同理（只读、只在重新识别时整个换掉）。
        """
        self._undo.append({
            "label": label,
            "rows": copy.deepcopy(self.rows),
            "enabled": dict(ENABLED),
            "settings": self._settings(),
            "index": self.index,
            "pristine_rows": self._pristine_rows,
            "pristine_index": self._pristine_index,
        })
        if len(self._undo) > 100:
            self._undo.pop(0)
        self.dirty = True

    def _on_destroy(self, event) -> None:
        """自己被销毁时，把挂在顶层的 Ctrl+Z 绑定摘掉（见 __init__ 里那段说明），
        顺带取消还没触发的自动滚动定时器——不然它下一跳会去操作一个已经
        销毁的画布（TclError）。"""
        if event.widget is not self.canvas:
            return                      # 子控件销毁也会冒上来，只认画布自己
        if self._autoscroll_job is not None:
            try:
                self.root.after_cancel(self._autoscroll_job)
            except tk.TclError:
                pass
            self._autoscroll_job = None
        top = getattr(self, "_undo_toplevel", None)
        if top is not None and top.winfo_exists():
            top.unbind_all("<Control-z>")

    def on_undo(self, event=None) -> None:
        # Ctrl+Z 绑在顶层窗口上（键盘快捷键没法只靠画布绑定生效），嵌进
        # 主窗口之后，切到别的页签这个绑定还在——用画布是不是"看得见"
        # （所在页签是不是当前选中的那个）判断这次 Ctrl+Z 是不是冲着这张
        # 表来的，不是就什么都不做，别把用户在别的页签按的 Ctrl+Z 撤到这来。
        if not self.canvas.winfo_exists():
            return                      # 已经被销毁（重建失败/换文档）
        if not self.canvas.winfo_viewable():
            return
        if self._drag is not None:
            return  # 拖动中途不给撤，避免状态打架
        if not self._undo:
            return
        snapshot = self._undo.pop()
        settings = snapshot["settings"]
        self.rows = snapshot["rows"]
        self.index = snapshot["index"]
        self._pristine_rows = snapshot["pristine_rows"]
        self._pristine_index = snapshot["pristine_index"]
        self._selected.clear()
        ENABLED.clear()
        ENABLED.update(snapshot["enabled"])
        self.cover_end = settings["cover_end"]
        self.delete_cover = settings["delete_cover"]
        self.renumber_refs = settings["renumber_refs"]
        if settings["heading_scheme"] != self.heading_scheme:
            # 撤回的是"切换标题编号方式"：主窗口导入区那组单选也得拨回去。
            self.heading_scheme = settings["heading_scheme"]
            if self._heading_scheme_cb is not None:
                self._heading_scheme_cb(self.heading_scheme)
        self.dirty = True   # 撤回也是"跟上次确认不一样"，同样要求重新确认
        annotate_blocks(self.rows)
        self.relayout()
        self.draw()

    # ------------------------------------------------------------ 提示闪动

    def _flash_content(self, rows: list[Row]) -> None:
        """红框快闪三下：指出"挡着停用的就是这几行"。

        状态栏负责解释原因，闪动负责指认位置——闪完即收，不无限闪。
        """
        c = self.canvas
        self._flash_gen += 1
        gen = self._flash_gen
        c.delete("flash")
        target = {id(r) for r in rows}
        for item in self.laid:
            if item["kind"] == "row" and id(item["row"]) in target:
                c.create_rectangle(2, item["y0"] + 1, self._last_width - 2, item["y1"] - 1,
                                   outline="#d93025", width=2, tags="flash")
        self._flash_tick(gen, 5)

    def _flash_tick(self, gen: int, ticks: int) -> None:
        c = self.canvas
        # 窗口/画布可能已经销毁（用户关掉结构页之后定时器才轮到）——那时
        # `find_withtag`/`itemconfigure` 会抛 TclError（审计 2026-09-15 核实）。
        if not c.winfo_exists() or not self.root.winfo_exists():
            return
        if gen != self._flash_gen or not c.find_withtag("flash"):
            return
        if ticks <= 0:
            c.delete("flash")
            return
        c.itemconfigure("flash", state="normal" if ticks % 2 else "hidden")
        self.root.after(170, lambda: self._flash_tick(gen, ticks - 1))


# ================================================================ 入口


def do_check(path: pathlib.Path) -> None:
    rows, _index = build_rows(path)
    print(f"共 {len(rows)} 行：")
    for row in rows:
        if row.kind == "zone":
            print(f"──── {ZONE_TEXT[row.zone]} ────")
        elif row.kind == "slot":
            print(f"  [空位] {row.slot_for}（启用={ENABLED.get(row.slot_for)}）")
        else:
            print(f"  [{str(row.doc_index):>3}] {row.kind:<8} col={row.col:<5} "
                  f"block={str(row.block):<6} {row.text[:30]}")
    print("ENABLED =", ENABLED)


def main() -> None:
    parser = argparse.ArgumentParser(description="识别结果调整窗口")
    parser.add_argument("--doc", type=pathlib.Path, default=DEFAULT_DOC)
    parser.add_argument("--check", action="store_true", help="只打印模型，不开窗")
    args = parser.parse_args()

    if args.check:
        do_check(args.doc)
        return

    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = tk.Tk()
    root.title("识别结果调整")
    try:
        root.state("zoomed")
    except Exception:
        root.geometry("1680x950")
    root.minsize(1280, 720)
    AdjustWindow(root, args.doc)
    print("窗口已打开。", flush=True)
    root.mainloop()


if __name__ == "__main__":
    main()
