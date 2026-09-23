"""主窗口（2026-09-13 按用户的整体布局图制作；2026-09-14 接上引擎和预设存盘）。

布局（照用户草图）：
- 上方六个页签：文件 / 页面 / 格式 / 结构 / 封面 / 关于；
- **右侧常设整张流程图**（不属于任何页签）：默认按面板宽度铺满（不看高度，
  用户 2026-09-13 定的），Ctrl＋滚轮可缩放（2026-09-19 按新流程图重建）；
  点图上节点＝**跳到对应页签＋让该点的按钮闪烁提示**（不直接执行功能）；
- 底部一条状态窗，所有按钮点击的反馈写在这里。

已经接上的：
- 「导入文件……」选 docx → 记路径、显示文件名；
- 「校对和调整识别结果」→ 真开调整窗口（AdjustWindow，吃已导入的 docx）；
- 「预设和配置」列表：默认从本地 JSON 读入/存出（`preset_io.py`），Excel 只是
  边缘产品——「从Excel导入配置文件」把 Excel 转成 JSON 落盘，每行的
  「导出为Excel」把 JSON 配置另存一份 Excel 供人可视化改；「保存当前设置为
  新配置」把「页面」「格式」「结构」「封面」当前的表单值收集起来存成新的
  JSON；「删除」真删文件（内置预设除外，二次确认）；
- 页面/格式/结构/封面四个页签——预设有数据就填、没有就用默认（控件都在）；
- 封面页签摆两张**真封面**样张（tools/make_cover_preview.py 烘的 PNG，点开可放大）；
- 关于页签：使用协议／《规范》原文浏览（spec_viewer）／下载官方版／
  工具操作逻辑说明，另有第三方依赖库许可的展开条目；
- 「导出空白模板到……」「导出论文到……」都已接上引擎（`template.build_template`/
  `pipeline.run`），封面字段从「封面」页签的"填写内容"区收集；
- 点击流程图节点 → 跳页签＋闪烁提示对应按钮（功能仍然由按钮自己执行）。

「联系开发者」相关入口（关于页署名、崩溃弹窗、启动失败兜底）已接通：统一
直接打开开发者主页（链接见 DEVELOPER_HOMEPAGE；用户 2026-09-19 拍板）。
使用协议/工具操作逻辑两个弹窗是定稿内容（AGREEMENT_TEXT / OPERATION_LOGIC_TEXT），
别把它们当成占位。
"""

from __future__ import annotations

import hashlib
import os
import pathlib
import re
import shutil
import sys
import tkinter as tk
import tkinter.font as tkfont
import traceback
from datetime import datetime
from tkinter import filedialog, messagebox, simpledialog, ttk

ROOT = pathlib.Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# 开发者主页——「关于」页署名、以及所有「联系开发者」入口统一打开它
# （2026-09-19 用户拍板：链接 = https://github.com/Grandliu66）。
DEVELOPER_HOMEPAGE = "https://github.com/Grandliu66"

# 「开发者的话」弹窗正文（用户 2026-09-19 提供，原文照录）。
# 排版约定（`_open_text_popup(formatted=True)` 认这套标记）：
#   "@" 开头的行＝居中加粗的小标题；">" 开头的行＝靠右对齐的落款（不缩进）；
#   其余每个自然段＝首行缩进两个汉字。
DEVELOPER_WORDS = """@关于本工具
读到这段话的人，你好！我是这款CUPL本科论文自动排版工具的开发者，很荣幸能在我的工具里和你见面。
我一开始其实没有想做一个公开分享、功能完备的工具，而只是用Agent帮我给自己的论文排版。排着排着，我诞生了“用Agent写一个可以自动按照法大论文标准排版的工具”的想法。
我其实并没有什么编程基础，如果非得说的话，只有作为计算机通识课的Python语言程序设计。不过诚实的说，开发这款工具并没有用到那门通识课的知识，因为所有代码工作都是AI完成的。尽管如此，这依然是比我想象的困难的多——AI总是听不明白话，而我也并不总是能说明白话；主要的识别算法逻辑也几乎是我自己设计的。在和AI扯皮了两周多以后，这款工具总算是开发完成了。
@本工具的功能
如你所见，本工具大概是用来识别本科论文，然后自动排版的。为什么是本科论文呢？其实我曾经考虑并且推进过硕博论文自动识别排版工作，但是由于学校规范的原因放弃了。学校的规范写得实在是太过自相矛盾，就连“某一级标题究竟应该居中还是靠左”这种基础问题都无法得出不自相矛盾的结论。无奈之下，只好放弃。
为此，本工具目前着重于本科论文的识别，主要有四个方面：毕业论文、学年论文、课程考核、学术十星。如果用本工具能给各位师兄弟姐妹的学业带来一点方便，那也就达到了我的目的。
不过，无论如何，人有失手、马有漏蹄，一些问题总是无可避免的。恳请各位师兄弟姐妹在使用本工具时仍审慎核对结果，不要太过相信。祝愿你们的论文取得成功。
@致谢
感谢Aeiex同学反馈的关于程序的问题以及提出的建议。
在这些为我提供过帮助的人之外，似乎我朦胧间仍然应该感谢一点什么。一个毫无编程能力的法学生有朝一日竟能以开发者的身份写下这段话，不得不说是并不心安理得的。究竟应该感谢什么呢？感谢LLM的开发公司，好像太抬举了一手交钱一手交货的买卖。感谢AI本身，他们似乎也理解不了，也并不在意。思来想去，或许应该感谢的是——
@感谢这个波澜壮阔的时代。

>2026年9月20日 于北京
"""


# 「更新日志」弹窗正文（用户 2026-09-20 提供）。用 `_open_text_popup(markup=True)`：
# `**…**` 加粗；以后的新版本追加在最前面。
CHANGELOG_TEXT = """**V1.1.1 2026/9/22**
**1. 修复了如下问题：**
（1）程序崩溃
· 修复了原文件移动、改名后点击“重新识别”“切换编号方案”“设置封面范围”会导致程序崩溃和设置丢失的问题。
· 修复了无法读取的TXT文件会导致程序崩溃的问题。
· 修复了TXT转换失败时文件不回滚、上一份已导入文档丢失的问题。
· 修复了设置文件内容有误时会导致程序崩溃的问题。
（2）文档内容与格式
· 修复了勾选“重新为文档生成”会导致目录域和清单域一并丢失的问题。
· 修复了来源文档的分节定义不在正文末尾时，导出后页面设置（含纸张尺寸）丢失的问题。
· 修复了清除直接格式时会错误清除正文中间分节符的问题。
· 修复了原文参考文献序号被包在超链接内时，剥除序号多删正文文字的问题。
· 修复了全篇没有识别到编号标题时，缺失部件被插到文末而非应在位置的问题。
· 修复了课程封面中填入超长信息时标签被拉变形的问题。
（3）设置项与编号
· 修复了行距“单倍”“双倍”“1.5倍”“最小值（磅）”四个选项导出时报错的问题。
· 修复了导入自带页眉的文档时，前置与正文两节页眉无法分别设置、互相覆盖的问题。
· 修复了参考文献编号定义在反复导出后改指、或与正文其他编号列表共用计数器的问题。
· 修复了结构页“部件启用”勾选未点击“确认”时，提示“不会应用”实际却已生效的问题。
· 修复了“每页重新编号”“序号后空格”两项设置只能勾选、取消勾选不生效的问题。
（4）识别与提示
· 修复了文字被拖入标签格再拖出后，仍显示旧标签、导出还误报“重复的部件”的问题。
· 修复了封面范围内的文字被错误识别、套上题注样式的问题。
· 修复了阿拉伯数字层级方案（1/1.1/1.1.1）下目录条目层级识别错误、全部归为同一级的问题。
（5）界面操作与显示
· 修复了结构页纵向滚动后，鼠标点击冻结表头会穿透命中被遮住的那一行的问题。
· 修复了说明窗口表格里，拖选的起止点落在非首/末行时高亮与复制内容错位的问题。
（6）配置文件读取
· 修复了Excel配置文件依靠空行分节、误删一行会导致后续内容错读的问题。
· 修复了Excel“脚注序号样式”下拉里有选项系统实际不识别的问题。
2. 本次更新不会影响原有JSON预设和Excel配置文件的读取，但修复了原有Excel配置文件中的一些读取错误，建议您重新核对并导入已有的Excel配置。

**V1.1.0 2026/9/20**
**1. 修复了如下问题：**
· 修复了在点击“取消导入”后协议勾选状态和导出按钮状态不一致的问题。（by Aeiex）
· 修复了部件标签下的占位语会被重复插入两次的问题。
· 修复了参考文献处超链接在导出版本中失去其颜色格式的问题。
· 修复了在标题后存在空行时，重复导出会导致重复插入副标题占位语的问题。
**2. 增加了如下功能：**
· 现在允许向窗口内拖入文件以导入。
· 现在允许设置各个部件标签前是否应用段前分页。
· 现在在格式页允许用户清除文档中的完全空行和分页符，但是该设置不会影响已经设置为封面的范围。
· 现在允许用户在格式页设置保留原文中的哪些直接格式，而不必直接修改配置文件。允许保留的直接格式增加了“颜色”项。
**3. 对既有功能做了如下优化：**
· 在封面中填入超长内容需要折行时，增加的行的下划线和第一行的下划线长度现在一致。
**4.** 本次更新更新了预设项和内容，但原有JSON预设和Excel预设仍然能够正常读取。本次更新还更新了致谢。
"""


def open_developer_homepage() -> bool:
    """打开开发者主页，返回是否真的打开了（链接没配 / 系统调不起浏览器就 False）。

    写成模块级函数（不是 `MainWindow` 的方法）是故意的：**崩溃弹窗也要用它**，
    那时候主窗口可能已经半死了。
    """
    if not DEVELOPER_HOMEPAGE:
        return False
    try:
        import webbrowser
        return bool(webbrowser.open(DEVELOPER_HOMEPAGE))
    except Exception:                                          # noqa: BLE001
        return False


def open_developer_contact() -> None:
    """「联系开发者」统一动作：直接开默认浏览器；打不开就把地址亮出来。

    2026-09-19 用户拍板：不再有“内容待补充”占位弹窗，所有入口一步到位。
    崩溃路径也会用到它，所以整体兜异常。
    """
    try:
        if not open_developer_homepage():
            messagebox.showinfo(
                "联系开发者", f"打不开浏览器。\n开发者主页：{DEVELOPER_HOMEPAGE}")
    except Exception:                                          # noqa: BLE001
        pass


def _enable_dpi_awareness() -> None:
    """Windows 高分屏适配第一步：声明进程 DPI 感知。

    **必须在本进程创建第一个 Tk 窗口之前调用**——创建以后再调不生效，Windows
    已经按"这个进程不感知 DPI"把窗口画成位图再拉伸了，字体/边框全糊
    （换开发机、显示器缩放比例跟旧机器不一样时最容易看出来）。放在模块顶层、
    import 时就跑，不管这个文件是被直接运行还是被别的脚本 `import
    essay_tool.gui.main_window` 拿去用，都能在 `MainWindow()`（继承
    `tk.Tk`）真正创建窗口之前抢先执行到。三级回退：Per-Monitor-V2
    （Win10 1703+，效果最好）→ shcore 的系统级感知（Win8.1+）→ 最老的
    `user32.SetProcessDPIAware()`。拿不到、或者已经被声明过（比如
    python.exe 自带清单已经设置），异常直接吞掉，不影响程序继续跑。
    """
    try:
        import ctypes
        # DPI_AWARENESS_CONTEXT 本质是指针大小的句柄，裸 Python int 在 64 位
        # 上传参会传错（实测：传 -4 这个 int，调用"成功"但返回值是失败 0，
        # 不抛异常也不生效）——必须包成 c_void_p 才对（2026-09-13 用诊断脚本
        # 实测确认：包完 SetProcessDpiAwarenessContext 才真的把 DPI 感知从
        # "不感知"改成了"按显示器"）。三个 API 的返回值**都要检查**，不能只
        # 看有没有抛异常：不抛异常不代表调用真的生效了。
        try:
            ok = ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
            if ok:
                return
        except Exception:
            pass
        try:
            # SetProcessDpiAwareness 返回 HRESULT，S_OK=0 才是成功；已经被上面
            # 设置过一次的话这里会拿到 E_ACCESSDENIED（DPI 感知只能设一次）。
            hr = ctypes.windll.shcore.SetProcessDpiAwareness(1)
            if hr == 0:
                return
        except Exception:
            pass
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


_enable_dpi_awareness()


def _fix_tk_scaling(root: tk.Tk) -> None:
    """高分屏适配第二步：校正 Tk 自己的点→像素换算。

    进程声明 DPI 感知只是让 Windows 别再拿位图硬拉伸；Tk/Tcl 8.6 自己算
    `tk scaling`（多少像素等于一磅）时，某些机器上仍然按固定 96 DPI 走，
    跟真实显示器的 DPI 对不上，字体点数虽然没变、量出来的像素尺寸却偏小，
    观感依然发虚。这里用 `winfo_fpixels("1i")`（DPI 感知生效后这个值就是
    真实的"每英寸像素数"）反推、显式钉一次 `tk scaling`。必须在
    `super().__init__()` 建完根窗口、但**在任何子控件/字体被创建之前**调用。
    """
    try:
        dpi = root.winfo_fpixels("1i")
        if dpi > 0:
            root.tk.call("tk", "scaling", dpi / 72.0)
    except Exception:
        pass


from essay_tool.engine.recognize import DEFAULT_HEADING_SCHEME  # noqa: E402
from essay_tool.gui.flow_nav import FlowPanel, C_LINK  # noqa: E402
from essay_tool.gui.text_wrap import wrap_text  # noqa: E402
from essay_tool.gui.operation_logic_text import OPERATION_LOGIC_TEXT  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING  # noqa: E402
from essay_tool.config.units import CHINESE_FONT_SIZES  # noqa: E402
from essay_tool.config.schema import HEADING_LABEL_PARTS  # noqa: E402
from essay_tool import __version__  # noqa: E402
from essay_tool.config.integrity import IntegrityError, verify_official  # noqa: E402
from essay_tool.config.paths import (builtin_presets_dir, cleanup_session_work_dir,  # noqa: E402
                                     import_backup_dir, purge_stale_work_dirs, resource_dir,
                                     session_work_dir, user_presets_dir)
from essay_tool.config.settings import GuiSettings, load_settings, save_settings  # noqa: E402

FAMILY = "Microsoft YaHei UI"
C_APP = "#f4f4f2"
C_PANEL = "#ffffff"
C_LINE = "#c9c9c5"
C_HINT = "#8a8a8a"
C_STRIP_HOT = "#e4e4e0"      # 折叠后最右边"展开使用向导"竖条：鼠标移上去时的底色
C_TEXT = "#222222"
# 封面总开关关掉时，页签里"复选框以下的全部内容"（区块标题、提示语、填写框、
# 样张图）统一压到这个灰度（用户 2026-09-14："只是不能选了，不是灰了"）。
C_DIM = "#b5b5b5"

# 页脚页码编码的显示名（对应 OOXML w:pgNumType 的 fmt 值）；下拉选项直接取
# 这张表的值列表，反查表自动跟着生成——加新格式只改这里。
PAGE_FMT_LABEL = {"decimal": "阿拉伯数字", "lowerRoman": "小写罗马",
                  "upperRoman": "大写罗马", "upperLetter": "大写字母",
                  "lowerLetter": "小写字母"}
# 反查：GUI 下拉框里选的中文标签 -> OOXML fmt 值（保存当前设置时要用）
PAGE_FMT_LABEL_REV = {v: k for k, v in PAGE_FMT_LABEL.items()}

# 格式表格用到的词汇表与常量都在 fmt_grid.py（自绘表格）里
from essay_tool.gui.fmt_grid import CELL_LINE, FmtGrid  # noqa: E402

# 结构部件名单，及规范规定的"缺失时插入的标签文字"默认值。2026-09-14 起
# 这份"插入文字"不再允许用户自定义（本来就只是缺失部件的标签怎么写，规范
# 已经明文规定了，开放自定义挑不出正面用途、只有把它改错的风险，见当天
# 讨论）；"默认启用"这一列是**不套用任何预设时的默认值**（用户 2026-09-14
# 定：只有题目/副标题/摘要/目录 默认开，其余一律关——跟
# adjust_preview.ENABLED 那份表保持一致）；ENABLED（调整表格的实时状态）
# 提到某个部件时以它为准。
# 引言/结论**不在名单里**（用户 2026-09-16 定："把这两个从预设里删掉，当完整
# 的标题"）：它们是正文里的一级标题，位置即身份，引擎按标题文字认角色，
# 没有开关、也不会因为"没识别到"而被补出一段占位文字。
# (部件名, 默认启用, 缺失时插入的标签文字, 说明)
STRUCTURE_PARTS = (
    ("题目", True, "", "内容每篇填写，这里只是开关"),
    ("副标题", True, "", "内容每篇填写，这里只是开关"),
    ("摘要", True, "【摘要】", ""),
    ("关键词", False, "【关键词】", ""),
    ("英文题目", False, "", "内容每篇填写，这里只是开关（题目本身是内容，没有标准答案可填）"),
    # 英文副标题默认**关**（2026-09-15 审计核实后改）：原来这里写 True，而同表的
    # 英文题目、英文摘要、英文关键词都是 False——用户点名的"不套预设时的默认值"
    # 只列了题目/副标题/摘要/目录/引言/结论，英文那几个都不在里面。开着它的后果
    # 是：文档里没有英文题目，导出却会多插一行"如不需要英文副标题，请把这一段
    # 整段删除"。要英文页就整组打开，别只开一半。（`adjust_preview.ENABLED` 里
    # 这一项一直就是 False，今次是把两处对齐。）
    ("英文副标题", False, "", "2026-09-15 起是独立部件（英文题目下面那一行副标题）"),
    ("英文摘要", False, "【ABSTRACT】", "插入文字＝摘要标签；英文关键词现在是独立部件"),
    ("英文关键词", False, "【KEYWORDS】", "2026-09-14 起从「英文摘要」里拆出来的独立部件"),
    ("目录", True, "【目录】", ""),
    ("插图清单", False, "插图清单", "规范[91]：与「插图和附表清单」二选一"),
    ("附表清单", False, "附表清单", "规范[91]：表格较多时才用"),
    ("插图和附表清单", False, "插图和附表清单", "规范[92]：与上面两行二选一"),
    ("参考文献", False, "参考文献", ""),
    ("附录", False, "附  录", ""),
    ("致谢", False, "致  谢", ""),
)

# 格式页「标签写法」里能自定义文字的部件（界面显示顺序）：STRUCTURE_PARTS 里默认
# 写法不为空的那些（＝标签类部件；题目/副标题/英文题目/英文副标题是内容占位语，
# 不许自定义）。引言、结论另有一行"两字之间空几格"（`HEADING_LABEL_PARTS`），
# 不在这张表里。
LABEL_PART_NAMES = tuple(n for n, _on, txt, _note in STRUCTURE_PARTS if txt)

# 格式表的四大块（用户 2026-09-13 拍板：整表按论文出现顺序排，分四块）。
# 名字要和预设「样式」表里的"名称"列对上；不在名单里的样式兜底进「其它」。
FMT_GROUPS = (
    ("前置部分", ("中文题目", "副标题", "摘要标签", "摘要正文", "关键词标签",
                  "关键词内容", "英文题目", "英文副标题", "英文摘要标签", "英文摘要正文",
                  "英文关键词标签", "英文关键词内容", "目录标题",
                  "目录一级", "目录二级", "目录三级")),
    ("正文部分", ("一级标题", "二级标题", "三级标题", "四级标题", "正文",
                  "长引文", "图题", "表题", "表格文字")),
    ("正文以后", ("参考文献条目", "附录标题", "附录序号题名", "致谢正文")),
    ("跨页元素", ("脚注文本", "页眉", "页脚页码")),
)

PRESET_DIR = builtin_presets_dir()
COVER_CARD_W = 236      # 样张卡片缩略图宽度（原图与"变灰版"共用同一尺寸）
BUILTIN_THESIS = "中国政法大学-本科毕业论文"
BUILTIN_YEAR = "中国政法大学-学年论文"
BUILTIN_COURSE = "中国政法大学-课程考核"
BUILTIN_TENSTAR = "中国政法大学-学术十星"
# 内置预设在列表里的固定顺序。**不能按文件名排**：启动时"不套用预设"的基线
# 取列表第一条（`_load_preset_defaults`），必须是毕业论文——按文件名排的话
# "学年论文"（学）排在"本科毕业论文"（本）前面，基线就悄悄变成学年论文了。
_BUILTIN_ORDER = (BUILTIN_THESIS, BUILTIN_YEAR, BUILTIN_COURSE, BUILTIN_TENSTAR)


def _preset_sort_key(path: pathlib.Path) -> tuple[int, str]:
    stem = path.stem
    rank = _BUILTIN_ORDER.index(stem) if stem in _BUILTIN_ORDER else len(_BUILTIN_ORDER)
    return rank, stem

# 「文件」页签里的配置列表：最少/最多多高（像素）。
# 用户 2026-09-16：「添加超过空位的预设时，将会将窗口无限拉长，直到将下面的
# 按钮挤没」——列表原来是"有多少行就多高"、那一格还 expand，每加一份预设
# 窗口就长一截。现在套一层定高滚动区：行少时仍自适应，超过上限就出滚动条。
PRESET_LIST_MIN_H = 34      # 大约一行（只有内置那两份时看着跟以前一样）
PRESET_LIST_MAX_H = 300     # 大约 8 行，再多就滚（内置预设已有四份，200 时四行就很挤）

# 「备份设置」那一块整体比别的块小一号（用户 2026-09-17：“将备份设置的要素
# 全都调小，放到配置文件之下”）：它是安全网、不是日常流程，缩一档就不跟
# “导入文件 / 预设和配置 / 导出”那几块争视线。要再调只改这四个数字。
BACKUP_HEAD_FONT = 11       # “备份设置”标题（原来 13）
BACKUP_TEXT_FONT = 9        # 复选框文字、输入框、下拉框（原来 12 / 10）
BACKUP_BTN_FONT = 9         # 两个小按钮（原来 10）
BACKUP_HINT_FONT = 8        # 灰字提示与“当前备份路径：…”（原来 9）

# 底部日志区（2026-09-14 新增）：默认展开的行数、最多保留的历史行数
LOG_LINES = 4           # 日志窗口默认显示几行（用户 2026-09-14：4 行）
LOG_MAX_LINES = 500
LOG_FONT_SIZE = 10      # 日志与窄栏的字号（用户 2026-09-14：调大一点）
# 结构页识别告警弹窗末尾的提示语（用户 2026-09-15 指定的意思）
REVIEW_ALERTS_HINT = "后续如需再次查看，可以前往屏幕下方状态窗口。"
# 切到「结构」页签后隔多久真去识别+弹窗——把同一瞬间的重复事件合并掉，
# 详见 `_on_tab_changed`。（250ms 用户感觉不到延迟，但足够把重复事件拼成一次）
REVIEW_DEBOUNCE_MS = 250
LOG_GRIP_H = 6          # 日志区上面那条"拖动条"的高度（px，用户嫌 8 太粗）
LOG_GRIP_BG = "#d9d9d4"         # 拖动条底色（鼠标移上去加深，示意它能拖）
LOG_GRIP_BG_HOVER = "#c2c2bb"
LOG_MIN_H = 60          # 拖动条能压到的最低高度；再矮请用「收起日志」
NB_MIN_H = 220          # 拖动时给上面记事本留的最小高度（窗口尺寸不动，从它那里挤）
LOG_WRAP_MARGIN = 8     # 日志换行时从 Text 宽度里留出的余量（px），免得贴边
# 封面样张图：tools/make_cover_preview.py 烘焙产出（GUI 运行时只读 PNG）
COVER_ASSETS_DIR = resource_dir() / "essay_tool" / "gui" / "assets"   # 打包后在 _internal 里
APP_ICON = COVER_ASSETS_DIR / "app.ico"


def apply_app_icon(win: tk.Misc) -> None:
    """标题栏/任务栏图标（替换 Tk 默认的羽毛）。`default=` 让之后建的所有弹窗
    也自动用它。图标缺失或读不了不是故障，静默沿用默认图标。"""
    try:
        win.iconbitmap(default=str(APP_ICON))
    except Exception:
        pass

from essay_tool.gui.widgets import (checkbutton, checkbutton_style_on, init_checkbox_style,
                                    wheel_notch)
from essay_tool.gui.win_place import center_on

# 「关于使用本工具的协议」正文——用户自己拟的条款，2026-09-14 定稿。
# 联系方式单独走一个按钮，不放在这份协议正文里，见 on_open_agreement 附近。
# 排版标记（2026-09-20，程序内弹窗与安装向导的 RTF 共用同一份原文，见 packaging/make_license.py）：
# 行首 "@" ＝ 居中加粗的大标题；"**…**" ＝ 加粗。去掉标记后就是逐字不变的原文。
AGREEMENT_TEXT = """@一、本工具的性质

1. 本工具为离线运行的论文格式识别与排版辅助软件，仅用于辅助用户完成文档格式整理工作。

2. **开发者不对排版结果的准确性、完整性及适用性作出任何保证。**用户在使用本工具生成排版结果后，仍应自行核对格式是否符合期刊、机构或提交单位的具体要求。**因排版误差、格式遗漏或识别偏差所导致的一切后果，由用户自行承担，开发者不承担任何责任。**

3. 本工具承诺：本工具不会修改用户的原始文件，一切修改将在复制后进行。但是，出于数据安全的考虑，开发者仍然建议用户自行备份重要文档。在确认本工具输出之文档足以达到用户对数据完整性的要求之前，不得删除原文档。**在使用本工具时，因用户操作不当或其它开发者难以预见或避免的原因造成数据丢失、损坏的，开发者不因此承担责任。**

4. 经由本工具排版的任何文档，其智力成果归属和一切合法权利仍归属原权利人；开发者不对经由本工具处理的文档主张任何权利，对自动取得的此类权利声明放弃。

5. **本工具系中国政法大学学生开发为公益之用，同中国政法大学官方没有关系**，开发过程中亦无得到校方帮助。本工具的任何源代码、预设、说明等，不代表中国政法大学官方观点；如有出入，则应以官方通知为准。

@二、用户使用本工具的范围

1. 用户使用本工具处理的文档内容应当合法合规。用户不得利用本工具处理、生成或传播含有违反国家法律法规的内容。

2. 用户不得利用本工具处理、生成或传播侵犯他人合法权益的文档，不得利用本工具从事学位论文代写、剽窃、伪造数据等《中华人民共和国学位法》所禁止的学术不端行为。

3. 用户利用本工具处理、生成或传递涉及国家秘密或商业秘密的文档，应当遵循国家法律法规和有关规范。本工具不对工具的安全性和漏洞情况做任何保证。

4. **因用户违反本条款而产生的法律责任或其他责任，由用户自行承担。**

@三、版权协议

本工具的程序、代码依据MIT协议共享。前述许可仅适用于本工具源代码本身的复制、修改及再分发，不影响用户在使用本工具（含已编译程序）时应当遵守本声明其他条款的义务。

@四、适用和争议解决

1. 上述条款自用户安装或使用本工具之时起生效。**用户使用本工具，即视为已阅读并同意本声明全部内容。**

2. 对涉及本工具的任何争议，应当适用中华人民共和国法律解决。**开发者和用户约定之管辖法院为北京市朝阳区人民法院。**"""

# 「工具操作逻辑」正文由工作簿生成，见 essay_tool/gui/operation_logic_text.py；
# 那份文件由 tools/make_operation_logic_text.py 从
# reference/工具操作逻辑写作表.xlsx 重出。要改文字就改工作簿再重跑那个脚本。

# 本工具自身的许可（2026-09-20 用户定：MIT，版权持有人写网名 Grandliu66，年份 2026）。
# 单一出处：关于页链接、项目根目录的 LICENSE、安装包里的 licenses\LICENSE.txt 都从这里来
# （后两者由 packaging/make_notices.py 生成，不要手改）。
# 注意范围：协议第三条写明 MIT 只适用于**源代码本身**；已编译程序的使用另须遵守
# 《关于使用本工具的协议》。弹窗里在许可全文后面补了这一句说明（`_OWN_LICENSE_NOTE`），
# 许可全文本身保持标准 MIT 原样。
_OWN_LICENSE = ("CUPL本科论文自动排版工具", "MIT License", """MIT License

Copyright (c) 2026 Grandliu66

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.""")
_OWN_LICENSE_NOTE = ("\n\n————————————————————\n"
                     "说明：上述许可适用于本工具的源代码。使用已编译的程序，"
                     "另须遵守《关于使用本工具的协议》（见「关于」页）。")

# 关于页「第三方依赖库版权声明」：原文照抄各库自带的 LICENSE 文件，不是转述。
# 只收本工具实际运行会用到的库（python-docx/openpyxl/lxml/Pillow 及其各自的
# 直接依赖）——pywin32、PyMuPDF 只在 tools/ 下的开发期脚本（跑 Word COM、烘封面
# 样张 PNG）里用到，不会进最终发给朋友的程序，不需要在这里声明。
_THIRD_PARTY_NOTICES = (
    ("python-docx", "MIT License", """The MIT License (MIT)
Copyright (c) 2013 Steve Canny, https://github.com/scanny

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE."""),
    ("openpyxl", "MIT License", """This software is under the MIT Licence

Copyright (c) 2010 openpyxl

Permission is hereby granted, free of charge, to any person obtaining a
copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be included
in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE."""),
    ("et_xmlfile", "MIT License", """This software is under the MIT Licence

Copyright (c) 2010 openpyxl

Permission is hereby granted, free of charge, to any person obtaining a
copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be included
in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE."""),
    ("lxml", "BSD 3-Clause License", """BSD 3-Clause License

Copyright (c) 2004 Infrae. All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are
met:

  1. Redistributions of source code must retain the above copyright
     notice, this list of conditions and the following disclaimer.

  2. Redistributions in binary form must reproduce the above copyright
     notice, this list of conditions and the following disclaimer in
     the documentation and/or other materials provided with the
     distribution.

  3. Neither the name of Infrae nor the names of its contributors may
     be used to endorse or promote products derived from this software
     without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
"AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL INFRAE OR
CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE."""),
    ("Pillow", "MIT-CMU License (PIL)", """The Python Imaging Library (PIL) is

    Copyright © 1997-2011 by Secret Labs AB
    Copyright © 1995-2011 by Fredrik Lundh and contributors

Pillow is the friendly PIL fork. It is

    Copyright © 2010 by Jeffrey 'Alex' Clark and contributors

Like PIL, Pillow is licensed under the open source MIT-CMU License:

By obtaining, using, and/or copying this software and/or its associated
documentation, you agree that you have read, understood, and will comply
with the following terms and conditions:

Permission to use, copy, modify and distribute this software and its
documentation for any purpose and without fee is hereby granted,
provided that the above copyright notice appears in all copies, and that
both that copyright notice and this permission notice appear in supporting
documentation, and that the name of Secret Labs AB or the author not be
used in advertising or publicity pertaining to distribution of the software
without specific, written prior permission.

SECRET LABS AB AND THE AUTHOR DISCLAIMS ALL WARRANTIES WITH REGARD TO THIS
SOFTWARE, INCLUDING ALL IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS.
IN NO EVENT SHALL SECRET LABS AB OR THE AUTHOR BE LIABLE FOR ANY SPECIAL,
INDIRECT OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES WHATSOEVER RESULTING FROM
LOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE
OR OTHER TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR
PERFORMANCE OF THIS SOFTWARE.

（Pillow 还按需打包了 libjpeg/libpng 等编解码库各自的许可证，完整版见其
官方仓库 LICENSE 文件，这里只收 Pillow 自身的许可声明。）"""),
    ("typing_extensions", "Python Software Foundation License Version 2",
     """PYTHON SOFTWARE FOUNDATION LICENSE VERSION 2

1. This LICENSE AGREEMENT is between the Python Software Foundation
("PSF"), and the Individual or Organization ("Licensee") accessing and
otherwise using this software ("Python") in source or binary form and
its associated documentation.

2. Subject to the terms and conditions of this License Agreement, PSF hereby
grants Licensee a nonexclusive, royalty-free, world-wide license to reproduce,
analyze, test, perform and/or display publicly, prepare derivative works,
distribute, and otherwise use Python alone or in any derivative version,
provided, however, that PSF's License Agreement and PSF's notice of copyright,
i.e., "Copyright (c) 2001-2023 Python Software Foundation; All Rights
Reserved" are retained in Python alone or in any derivative version prepared
by Licensee.

3. In the event Licensee prepares a derivative work that is based on
or incorporates Python or any part thereof, and wants to make
the derivative work available to others as provided herein, then
Licensee hereby agrees to include in any such work a brief summary of
the changes made to Python.

4. PSF is making Python available to Licensee on an "AS IS"
basis. PSF MAKES NO REPRESENTATIONS OR WARRANTIES, EXPRESS OR
IMPLIED. BY WAY OF EXAMPLE, BUT NOT LIMITATION, PSF MAKES NO AND
DISCLAIMS ANY REPRESENTATION OR WARRANTY OF MERCHANTABILITY OR FITNESS
FOR ANY PARTICULAR PURPOSE OR THAT THE USE OF PYTHON WILL NOT
INFRINGE ANY THIRD PARTY RIGHTS."""),
)


# ------------------------------------------------------------ 崩溃提示

def _make_crash_code(exc_type: type, tb_text: str) -> str:
    """"至少有点识别意义"的错误代码：异常类型名 + traceback 全文的短哈希。

    不是手工维护一张错误码表——那种做法对个人/小范围分发的工具工作量不
    成比例（每加一处新的可能异常都要回来登记一个码）。这里换一个几乎不
    花成本的办法：同一个 bug 每次触发，traceback 文本一样，算出来的代码
    就一样；用户报"错误代码 XXX"过来，光看代码就能判断是不是同一个已知
    问题，不需要对方先把完整堆栈发过来才能确认。
    """
    digest = hashlib.sha1(tb_text.encode("utf-8", "replace")).hexdigest()[:8].upper()
    return f"{exc_type.__name__}-{digest}"


def _log_crash(code: str, exc_type, exc_value, tb) -> None:
    """把完整堆栈追加写进本地日志文件，供以后真联系开发者时附上细节。

    纯本地文件、不联网，跟《使用须知》"完全离线运行"那句不冲突；写失败
    （比如磁盘满了）不应该拦住崩溃弹窗本身照常显示，所以吞掉异常。
    """
    try:
        from essay_tool.config.paths import user_data_dir
        log_path = user_data_dir() / "crash.log"
        text = "".join(traceback.format_exception(exc_type, exc_value, tb))
        with log_path.open("a", encoding="utf-8") as f:
            f.write(f"\n===== {datetime.now().isoformat(timespec='seconds')}"
                    f"  错误代码：{code} =====\n")
            f.write(text)
    except OSError:
        pass


def _render_crash_popup(master, code: str, backup_dir, on_contact, on_close) -> tk.Toplevel:
    """崩溃提示窗口本体——`MainWindow` 正常运行时崩溃、和"连主窗口都建不起来"
    这两种情况共用同一份界面（后者传一个干净的 `tk.Tk()` 当 master）。"""
    win = tk.Toplevel(master)
    win.title("出现问题")
    win.configure(bg=C_PANEL)
    win.resizable(False, False)

    def label(text: str, **kw) -> None:
        kw.setdefault("font", (FAMILY, 10))
        kw.setdefault("bg", C_PANEL)
        kw.setdefault("fg", C_TEXT)
        kw.setdefault("justify", "left")
        # 换行宽度（2026-09-19 从 430 放宽到 560：窗口已按内容自适应，
        # 旧值会让说明文字被折得太窄、右侧空一大块）
        kw.setdefault("wraplength", 560)
        tk.Label(win, text=text, **kw).pack(anchor="w", padx=18, pady=(0, 10))

    tk.Label(win, text="哎呀！出现了一点小问题，程序已崩溃。", font=(FAMILY, 12, "bold"),
             bg=C_PANEL, fg=C_TEXT).pack(anchor="w", padx=18, pady=(18, 12))
    label("你的原文件未受本次崩溃影响。\n"
          f"如果打开了原文件备份，你可以到 {backup_dir} 查看原文件之备份；"
          "特别是如果你已经删除了原文件。")
    label(f"错误代码：{code}", font=(FAMILY, 9), fg=C_HINT)

    # 崩溃路径里点按钮**再**出错，绝不能静默——那就真成"点不动"了。
    # 关闭类动作一旦失败就直接硬退出：宁可进程走，也不能把用户扣在一个
    # 关不掉的窗口里（用户 2026-09-17：“必须 Alt+F4 才能关闭”）。
    def _guard(action, *, must_exit: bool = False):
        def inner():
            try:
                action()
            except Exception:                                  # noqa: BLE001
                if must_exit:
                    os._exit(1)
        return inner

    run_close = _guard(on_close, must_exit=True)
    run_contact = _guard(on_contact)

    row = tk.Frame(win, bg=C_PANEL)
    row.pack(anchor="w", padx=18, pady=(0, 6))
    tk.Label(row, text="如果该问题反复出现，你可以选择：", font=(FAMILY, 10),
             bg=C_PANEL, fg=C_TEXT).pack(side="left")
    tk.Button(row, text="联系开发者", font=(FAMILY, 10), padx=8, pady=2,
              command=run_contact).pack(side="left", padx=(8, 0))

    tk.Button(win, text="关闭程序", font=(FAMILY, 10), padx=14, pady=5,
             command=run_close).pack(pady=(16, 14))

    # 窗口按内容自适应（2026-09-19 用户实拍：显示缩放大的机器上，写死 480×340
    # 会把标题和「联系开发者」按钮挤没——按钮只剩 1 像素宽）。让 tk 先算完各
    # 控件的需求尺寸，再按需求加留白定窗口，最后居中；任何缩放下都不会再挤。
    win.update_idletasks()
    win_w = max(win.winfo_reqwidth() + 24, 520)
    win_h = max(win.winfo_reqheight() + 16, 360)
    win.geometry(f"{win_w}x{win_h}")
    center_on(win, master, win_w, win_h)

    win.protocol("WM_DELETE_WINDOW", run_close)
    # 键盘退路：不管鼠标为什么点不动，Esc / 回车都能把它关掉
    win.bind("<Escape>", lambda _e: run_close())
    win.bind("<Return>", lambda _e: run_close())
    win.transient(master)
    # ⚠️ 之前可能有人（某个模态弹窗、或上一个弹窗）按着 grab——不先放掉的话，
    # 鼠标点击会被转给它，表现出来就是"这个窗口的按钮怎么点都没反应，只能
    # Alt+F4"（用户 2026-09-17 实测；Alt+F4 走窗口管理器、不受 grab 管，
    # 所以那是唯一还能用的路）。
    for holder in (master, win):
        try:
            holder.grab_release()
        except Exception:                                      # noqa: BLE001
            pass
    win.grab_set()
    # 崩溃发生时用户可能没在看这个窗口（切去了别的应用/别的 Toplevel 在最上层）
    # ——grab_set 只管"点不到主窗口"，不保证这个窗口本身显示在最前面、抢到
    # 焦点，所以另外强制提到最前排。
    win.lift()
    win.focus_force()
    return win


class MainWindow(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        _fix_tk_scaling(self)
        apply_app_icon(self)
        self.title("CUPL本科论文自动排版工具")
        self._flow_width = self._set_default_geometry()
        # 默认最大化：必须在窗口第一次映射**之前**就定好（geometry() 之后、任何
        # update 之前）。放到构造完再最大化的话，用户会先看到一个普通大小的窗口
        # 闪一下再撑满（用户 2026-09-20 反馈）。
        try:
            self.state("zoomed")
        except tk.TclError:
            pass
        self.minsize(880, 520)
        self.configure(bg=C_APP)

        # 原文的标题编号方案（导入区那组单选；用户 2026-09-19：只有"中文"和
        # "点分"两项，强制单选，默认中文）。识别（结构页/告警/导出）都读它，
        # 状态放在主窗口而不是结构页实例上：用户没打开过结构页时导出走自动
        # 识别，也得按这个来。
        self.var_heading_scheme = tk.StringVar(value=DEFAULT_HEADING_SCHEME)
        self._heading_scheme_applied = DEFAULT_HEADING_SCHEME   # 结构页当前实际用的
        self.doc_path: pathlib.Path | None = None
        # 用户真正选中的那个文件（`.txt` 会被转成 docx 副本再识别，见
        # `on_import_file`）：界面显示、备份都用它。
        self.src_path: pathlib.Path | None = None
        self.applied_preset: pathlib.Path | None = None
        # 内嵌在"论文结构"页签里的调整表格实例（见 _rebuild_struct_adjust）；
        # 导出的时候要拿它当前的 rows+index 换算成指令表（见 _build_ops_for_export）。
        self._adjust_instance = None
        # 结构页最近一次写进日志的识别结论（(文档, 那几行)）——同一份只记一次，
        # 见 _log_review_once：来回切页签不该把日志刷满。
        self._review_logged: tuple | None = None
        # 待兑现的"切到结构页就识别"任务（见 _on_tab_changed 的防抖）
        self._review_job: str | None = None
        # 这一次进结构页已经报告过的文档（字符串路径）；离开这一页就清账。
        # ttk 会补投重复的页签事件，靠这个账把"同一次驻留"锁成只报一次，
        # 见 _on_tab_changed / _run_review。
        self._review_popped: str | None = None
        self._preset_extra: dict = {}   # 表单管不到的几项预设值，见 _load_page_from
        self.gui_settings: GuiSettings = load_settings()   # 备份相关的几项持久化设置
        self._preset_rows: list[dict] = self._load_preset_rows()
        self._init_checkbox_images()

        self._build_paned()          # 里面顺带建左侧那一栏的日志区，见 _build_bottom
        self._set_exports_enabled(False)
        self._load_preset_defaults()
        self.status("启动成功，你可以根据右侧向导图使用本工具")
        from essay_tool.gui.win_drop import enable_file_drop
        enable_file_drop(self, self._on_files_dropped)

    def _init_checkbox_images(self) -> None:
        """注册自制的复选框样式（实现在 `gui/widgets.py`，两个窗口共用）。

        自造方块图 + ttk 自定义布局的原因（原生 indicator 不跟字号放大、选中态
        还会右下偏移）写在那个模块的模块说明里，这里不再重复。
        """
        init_checkbox_style(self, C_PANEL)

    def _checkbutton(self, parent: tk.Widget, text: str = "", *,
                     variable: tk.BooleanVar, command=None,
                     font_size: int = 12) -> ttk.Checkbutton:
        """统一入口造出来的复选框（ttk + 自造方块，见 `widgets.checkbutton`）。"""
        return checkbutton(parent, text, variable=variable, command=command,
                           font_size=font_size)

    def _set_default_geometry(self) -> int:
        """默认打开尺寸、右侧流程图的宽度都按屏幕尺寸算，不写死像素数。

        以前是写死的 1380x860／流程图侧栏 566px——不同缩放比例/分辨率的屏幕上
        "合适"的像素数差很远：这台机器实测是 175% 缩放（真实 DPI 168，见
        `_fix_tk_scaling` 的注释），DPI 感知修好之后同样的像素数比以前占的
        屏幕比例小了一截，窗口和流程图看着都偏小（用户 2026-09-13 反馈）。
        改成按 `winfo_screenwidth/height`（DPI 感知生效后这两个数就是真实
        物理像素）算比例，各种屏幕上都合适。返回值是流程图侧栏该给多宽——
        用户明确要求"调大占比"，不是简单跟着窗口变大而按比例放大，所以
        这里把占比从原来的 566/1380≈41% 提到约 46%。
        """
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w = max(1380, min(int(sw * 0.82), sw - 80))
        h = max(860, min(int(sh * 0.85), sh - 80))
        # 默认最大化打开（用户 2026-09-20）：流程图侧栏按整块屏幕宽度算占比；
        # 上面的 w×h 是"还原"（取消最大化）之后的窗口尺寸。
        w_flow = sw
        # 向导栏默认占窗口宽度的三分之一（用户 2026-09-20）
        # 首屏不给位置，让系统居中；窗口一映射就立即用 _pin_window_geometry()
        # 把位置显式写回 geometry（不带位置的话 Tk 不记位置，内容尺寸一变
        # Windows 就会自己挪窗口，见 _pin_window_geometry 的注释）。
        self.geometry(f"{w}x{h}")
        return max(420, int(w_flow / 3))

    def _set_window_size(self, w: int, h: int) -> None:
        """改窗口尺寸，**并要求保持当前位置**。

        只写 "WxH" 的话 Tk 不把位置视为"用户指定"，之后窗口的"请求尺寸"
        一变（拖动日志条、展开日志面板都会让内容请求变），Windows 就会自己
        挪窗口——用户 2026-09-14 看到的"一拉条上面往上跑、下界往下跑"就是
        这么来的。带上 +x+y 之后位置就钉死了。
        """
        self.geometry(f"{int(w)}x{int(h)}+{self.winfo_x()}+{self.winfo_y()}")

    def _pin_window_geometry(self) -> None:
        """窗口映射完成后，把当前位置显式写回 geometry（见 _set_window_size）。"""
        if self.state() == "normal" and self.winfo_viewable():
            self._set_window_size(self.winfo_width(), self.winfo_height())

    def _on_window_mapped(self, event=None) -> None:
        """首次映射出来时钉住几何，随后注销这个绑定（只需做一次）。"""
        if event is not None and event.widget is not self:
            return
        self.after_idle(self._pin_window_geometry)
        self.unbind("<Map>")

    # ------------------------------------------------------------ 崩溃提示

    def report_callback_exception(self, exc, val, tb) -> None:
        """Tkinter 的钩子：任何按钮/事件回调里没被 try/except 接住的异常都会
        走这里。默认行为是打印到控制台、程序继续跑；这里换成告诉用户的
        弹窗，并且**弹窗关闭后退出程序**——出了没预料到的异常，界面状态可能
        已经不对了，装作没事继续跑比明说"重开一下"更危险。

        这个方法本身也可能出错（比如弹窗都建不起来）——那种情况别再套娃，
        退回最朴素的 messagebox，然后还是退出。
        """
        try:
            tb_text = "".join(traceback.format_exception(exc, val, tb))
            code = _make_crash_code(exc, tb_text)
            _log_crash(code, exc, val, tb)
            self._show_crash_dialog(code)
        except Exception:
            try:
                messagebox.showerror("出现问题", "程序遇到了没预料到的错误，即将关闭。")
            except Exception:
                pass
            self.destroy()

    def _show_crash_dialog(self, code: str) -> None:
        def close() -> None:
            win.destroy()
            self.destroy()

        win = _render_crash_popup(self, code, self._effective_backup_dir(),
                                  on_contact=open_developer_contact, on_close=close)

    def on_open_developer_homepage(self) -> None:
        """「关于」页署名点开：直接用默认浏览器打开开发者主页。"""
        if not open_developer_homepage():
            self.status(f"打开浏览器失败。开发者主页：{DEVELOPER_HOMEPAGE}")

    def on_open_developer_words(self) -> None:
        """「关于」页「开发者的话」点开：弹一个信式排版的文字窗口。"""
        self._open_text_popup("开发者的话", DEVELOPER_WORDS, formatted=True)

    def on_open_changelog(self) -> None:
        """「关于」页「更新日志」点开：弹文字窗口（`**` 加粗）。"""
        self._open_text_popup("更新日志", CHANGELOG_TEXT, markup=True)

    # ------------------------------------------------------------ 布局骨架

    def _build_bottom(self, master: tk.Frame) -> None:
        """左侧主控面板（记事本）**下方**那一栏：可上下翻的日志区 + "最新提示"窄栏。

        用户要求把原来的单行状态栏扩展成"常驻窗口、能上下滚动看历史提示"，
        并选定 B 方案、**默认展开**；2026-09-14 又定稿了位置：挂在**左侧主控
        面板下面**，右侧流程图占满整窗高度（原来日志区是横跨全窗底部、把右侧
        流程图截成半截）。

        - **上半**是日志区（只读 Text + 竖向滚动条），每写一条提示追加一行，
          行首带 `[时:分:秒]`，只留最近 `LOG_MAX_LINES` 行；
        - **最下**窄栏**只在日志区折叠时**显示最近一条日志，展开时留空
          （用户 2026-09-15："禁用这个位置显示任何事情"＋"在日志窗口末折叠
          时留空，在日志窗口折叠时显示最近的日志"），另加「收起/展开日志」
          和「清空」两个按钮；
        - **日志区上沿**那条 "拖动条"（`_log_grip`）按住上下拖，只改日志区高度
          （往上拉就从上面的记事本里挤；上下限见 `_clamp_log_px`）。
          手柄要拿在**被拉区域的边上**，而且**别往里放 Label**：Frame 的 height
          只是"建议值"，有子部件时会被撑高（实测会变成二十几像素）。
          **不用 Tk 的 sash**：panedwindow 的 sash 又细、又要先把 paned 自己撑高
          才拖得动，实测拖不动。

        **拖动和收起/展开都不动主窗口**（用户 2026-09-14）：收起时腾出来的高度
        直接还给上面的记事本，所以它不再挡住记事本（用户原话）。

        三段的 pack 顺序（都 side="bottom"，**先 pack 的贴最下**）：窄栏 → 日志区
        → 拖动条，自上而下才是 拖动条 / 日志区 / 窄栏。`_set_log_px()` 里补 pack 时
        也按这个顺序，别弄反。

        日志区只在**用户本来就停在底部**时自动滚到最新：他正往上翻历史时
        不该被拽回来（滚动窗口的标准礼貌）。
        """
        bottom = tk.Frame(master, bg="#e9e9e6")
        bottom.pack(side="bottom", fill="x")
        self._bottom_box = bottom

        # —— 窄栏（最下）：最新一条 + 两个按钮
        bar = tk.Frame(bottom, bg="#e9e9e6")
        self._log_bar = bar
        self.btn_log_toggle = tk.Button(bar, text="收起日志", font=(FAMILY, 9),
                                        bd=0, bg="#e9e9e6", fg=C_HINT,
                                        activebackground="#e9e9e6",
                                        cursor="hand2", command=self._toggle_log)
        self.btn_log_toggle.pack(side="right", padx=(4, 8), pady=4)
        tk.Button(bar, text="清空", font=(FAMILY, 9), bd=0, bg="#e9e9e6",
                  fg=C_HINT, activebackground="#e9e9e6", cursor="hand2",
                  command=self._clear_log).pack(side="right", pady=4)
        self.lbl_status = tk.Label(bar, text="", font=(FAMILY, LOG_FONT_SIZE),
                                   bg="#e9e9e6", fg=C_TEXT, anchor="w")
        self.lbl_status.pack(side="left", padx=10, pady=4, fill="x", expand=True)
        bar.pack(side="bottom", fill="x")

        # —— 日志区（中间，紧贴在窄栏上方；高度由 _set_log_px 控制）
        self._log_box = tk.Frame(bottom, bg="#fbfbf9")
        self._log_box.pack_propagate(False)     # 高度只听自己的 height，不听 Text
        # 换行由我们自己算（`text_wrap.wrap_text`），不用 Tk 自带的那两种：
        # `wrap="word"` 在中文里会把整句当一个词、然后在任意字符处硬折
        # （用户 2026-09-15 抓到的"第 13 / 段｜目录：…"），`wrap="char"` 又会
        # 把标点甩到行首、把英文单词拆开。所以这里关掉 Tk 的换行、宽度变化时
        # 自己重排，见 `_render_log`。
        self.log_text = tk.Text(self._log_box, wrap="none", bd=0,
                                highlightthickness=0, bg="#fbfbf9", fg=C_TEXT,
                                font=(FAMILY, LOG_FONT_SIZE), cursor="arrow")
        log_vsb = ttk.Scrollbar(self._log_box, orient="vertical",
                                command=self.log_text.yview)
        # 滑块同步走自己的回调：除了 set 滑块，还要记"用户停没停在底部"
        # （见 `_on_log_yscroll`——尺寸变化也会走到这里，不能当成滚动）
        self._log_view_h = 0
        self._log_follow = True          # True＝自动跟最新的那几条

        def _sync_log__scroll(first, last) -> None:
            log_vsb.set(first, last)
            self._on_log_yscroll(first, last)

        self.log_text.configure(yscrollcommand=_sync_log__scroll)
        log_vsb.pack(side="right", fill="y")
        self.log_text.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=2)
        self.log_text.configure(state="disabled")   # 只读，但可以选中复制
        # 原始（未换行）的提示行；显示时按当前宽度现算换行，见 `_render_log`
        self._log_raw: list[str] = []
        self._log_render_job = None
        self.log_text.bind("<Configure>", self._on_log_configure, add="+")
        self._log_box.pack(side="bottom", fill="x")

        # —— 拖动条（最上，压在日志区上沿）：自己实现拖拽，见方法说明
        self._log_grip = tk.Frame(bottom, bg=LOG_GRIP_BG, height=LOG_GRIP_H,
                                  cursor="sb_v_double_arrow")
        self._log_grip.pack_propagate(False)    # 别让子部件把这条撑粗
        self._log_grip.pack(side="bottom", fill="x")
        self._log_grip.bind("<Enter>", lambda e: self._log_grip.config(bg=LOG_GRIP_BG_HOVER))
        self._log_grip.bind("<Leave>", lambda e: self._log_grip.config(bg=LOG_GRIP_BG))
        self._log_grip.bind("<ButtonPress-1>", self._on_grip_press)
        self._log_grip.bind("<B1-Motion>", self._on_grip_drag)
        self._log_grip.bind("<ButtonRelease-1>", self._on_grip_release)

        self._log_open = True
        self._log_h_px = 0          # 展开时的高度（像素），_init_log_height 里量
        self._grip_y0 = None        # 拖动条按住时的鼠标 y / 日志区高度
        self._grip_h0 = None
        self.after_idle(self._init_log_height)
        # 窗口映射后把位置钉住（见 _pin_window_geometry）。用 <Map> 而不是
        # after_idle：__init__ 里窗口还没被映射，那时 winfo_viewable() 仍是假。
        self.bind("<Map>", self._on_window_mapped, add="+")

    def _init_log_height(self) -> None:
        """布局稳定后量出"LOG_LINES 行"的日志区高度（像素），并应用一次。

        这次调用**不动窗口**：启动时窗口高度是 _set_default_geometry() 按屏幕
        算好的，日志区只是从左栏里分一块出来。
        """
        self.log_text.configure(height=LOG_LINES)
        self.update_idletasks()
        self._log_h_px = self.log_text.winfo_reqheight() + 4   # ＋pack 的 pady
        self._set_log_px(self._log_h_px)

    def _set_log_px(self, px: int) -> None:
        """把日志区高度设成 px（**主窗口尺寸、位置一概不动**）。

        用户 2026-09-14 终稿：日志区挂在左侧主控面板下面，不管拖动还是收起/
        展开，都只改左栏内部的高度分配（多出来的高度从上面的记事本里挤，
        收起时直接还给记事本）。不碰窗口尺寸也就不会有"窗口自己上下跳"的
        那些毛病（踩过：几何只写 WxH 时 Tk 不把位置当用户指定，Windows 会
        自己挪窗口，见 `_set_window_size`）。
        """
        px = max(0, int(px))
        if px <= 0:
            self._log_box.pack_forget()
            self._log_grip.pack_forget()
            self._log_open = False
        else:
            if not self._log_open:
                # 先日志区、后拖动条 ⇒ 自上而下 拖动条/日志区/窄栏（同 _build_bottom）
                self._log_box.pack(side="bottom", fill="x")
                self._log_grip.pack(side="bottom", fill="x")
                self._log_open = True
            self._log_box.configure(height=px)
        self.update_idletasks()    # 让 _clamp_log_px 读到的 nb 高度是新的

    def _toggle_log(self) -> None:
        """收起／展开日志区（默认展开）。只改左栏高度分配，不动主窗口。"""
        if self._log_open:
            self._log_h_px = max(LOG_GRIP_H + 40, self._log_box.winfo_height())
            self._set_log_px(0)
            self.btn_log_toggle.config(text="展开日志")
        else:
            self._set_log_px(self._log_h_px)
            self.btn_log_toggle.config(text="收起日志")
            self.log_text.see("end")
        self._refresh_log_bar()

    def _on_grip_press(self, event) -> None:
        self._grip_y0 = event.y_root
        self._grip_h0 = self._log_box.winfo_height()

    def _clamp_log_px(self, px: int) -> int:
        """拖动时的高度上下界：最矮 `LOG_MIN_H`（再矮请用「收起日志」按钮），
        最高让上面的记事本至少还剩 `NB_MIN_H`。

        左栏里"记事本＋日志区"的总高度是固定的（窗口尺寸不动），所以上限
        就是"按现在的总高给记事本留够 NB_MIN_H"。
        """
        room = (self.nb.winfo_height() + self._bottom_box.winfo_height()
                - self._log_bar.winfo_height() - LOG_GRIP_H)
        top = max(LOG_MIN_H, room - NB_MIN_H)
        return max(LOG_MIN_H, min(int(px), top))

    def _on_grip_drag(self, event) -> None:
        """按住拖动条上下拖：往上拖＝日志区变高（占上面记事本的地方）。"""
        if self._grip_y0 is None or self._grip_h0 is None:
            return
        self._set_log_px(self._clamp_log_px(
            self._grip_h0 - (event.y_root - self._grip_y0)))

    def _on_grip_release(self, event) -> None:
        self._on_grip_drag(event)
        self._grip_y0 = self._grip_h0 = None

    def _clear_log(self) -> None:
        self._log_raw.clear()
        self._render_log()
        self._refresh_log_bar()

    def _log_append(self, msg: str) -> None:
        """往日志区追加一行。

        只存**未换行的原文**，显示时按当前宽度现算换行——窗口变宽变窄时
        重排一遍就还是整齐的，不会留下按旧宽度折出来的断行。
        """
        # 追加前先看用户是不是停在底部；不在底部就不要把他拽回来
        at_bottom = self._log_follow or self.log_text.yview()[1] >= 0.999
        self._log_raw.append(f"[{datetime.now():%H:%M:%S}] {msg}")
        if len(self._log_raw) > LOG_MAX_LINES:      # 只留最近若干行
            del self._log_raw[:len(self._log_raw) - LOG_MAX_LINES]
        self._render_log(scroll_to_end=at_bottom)
        self._refresh_log_bar()

    def _on_log_yscroll(self, first: str, last: str) -> None:
        """日志区滚动回调：除了同步滑块（调用方做），还要记"用户停没停在底部"。

        ⚠️ 这个回调**尺寸一变也会被调到**（Tk 靠它同步滑块），而"把日志区拉大
        又收回"正是尺寸变化：此时视口上端不动、下端跟着变，直接按当前值记就会
        把"本来停在底部"记成"没停在底部" ⇒ 收回后视图被钉在最早那几条上
        （用户 2026-09-17 报的正是这个）。所以只在**高度没变**时才更新标记；
        高度变了说明这是尺寸事件，保持原标记（＝保持"跟随最新"的意图）。
        """
        try:
            height = self.log_text.winfo_height()
        except Exception:                                      # noqa: BLE001
            return
        if height == self._log_view_h:
            try:
                self._log_follow = float(last) >= 0.999
            except (TypeError, ValueError):
                pass
        self._log_view_h = height

    def _on_log_configure(self, event) -> None:
        """日志区大小变了（拖分栏条、改窗口大小）→ 重新算一遍换行。

        防抖 150 ms：拖动过程中会连发几十次 Configure，每次都重排会很卡。
        """
        if getattr(event, "widget", None) is not self.log_text:
            return
        if self._log_render_job is not None:
            try:
                self.after_cancel(self._log_render_job)
            except Exception:                      # noqa: BLE001
                pass
        self._log_render_job = self.after(150, self._render_log)

    def _render_log(self, scroll_to_end: bool | None = None) -> None:
        """按当前宽度把 `_log_raw` 重排一遍填进 Text。

        重排会把手动滚动的位置弄丢，所以按**比例**恢复；`scroll_to_end=True`
        时（或者本来就在底部）直接贴到底。
        """
        self._log_render_job = None
        if not self._log_raw:
            self.log_text.configure(state="normal")
            self.log_text.delete("1.0", "end")
            self.log_text.configure(state="disabled")
            return
        first = self.log_text.yview()[0]
        # ⚠️ "在不在底部"看 `_log_follow`（**尺寸变化不会改它**），不能只在
        # 这里现算 yview：日志区一被拉大／收回，视口上端不动、下端跟着变，
        # 现算出来的就是"没在底部"——于是收回后视图被钉在**最早的**那几条上，
        # 更新的反而被挡在下面（用户 2026-09-17："正确的做法是锁定底部的
        # 状态，让上方窗口挡住更早的状态"）。yview 只当兼底（内容整屏装得下
        # 时它恒为 (0,1)，见 `_on_log_yscroll` 里那段说明）。
        at_bottom = (scroll_to_end if scroll_to_end is not None
                     else (self._log_follow or self.log_text.yview()[1] >= 0.999))
        font = tkfont.Font(font=self.log_text.cget("font"))
        # 可用宽度：Text 本体宽度减掉两侧的内边距（bd=0、无边框，只剩点余量）
        width = max(80, self.log_text.winfo_width() - LOG_WRAP_MARGIN)
        cache: dict[str, float] = {}

        def measure(text: str) -> float:
            total = 0.0
            for ch in text:
                got = cache.get(ch)
                if got is None:
                    got = float(font.measure(ch))
                    cache[ch] = got
                total += got
            return total

        rendered: list[str] = []
        for raw in self._log_raw:
            rendered.extend(wrap_text(raw, measure, width))
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.insert("1.0", "\n".join(rendered))
        self.log_text.configure(state="disabled")
        if at_bottom:
            self.log_text.see("end")
        else:
            self.log_text.yview_moveto(first)


    def _build_paned(self) -> None:
        """左边记事本、右边流程图，中间加一条可拖动的分隔条（用户 2026-09-13
        要求"加一个拖动调整窗口比例的功能"）。

        以前两块各自 `pack` 出固定宽度（流程图侧栏宽度是 `_set_default_geometry`
        按屏幕算好写死的），没法手动调。换成 `tk.PanedWindow`（横向）：中间的
        sash 本身就是可拖动的分隔条，比自己写鼠标拖拽事件简单可靠。
        `_flow_width` 仍然当"默认宽度"用（sash 的初始位置），拖动之后就由
        用户自己定了。记事本这边 `stretch="always"`——窗口整体拉伸时多出来的
        空间给它，流程图那边保持用户拖出来的宽度不跟着变。

        **左侧这一栏是竖着分两块的**（2026-09-14 用户终稿）：上面记事本、下面
        日志区（`_build_bottom`），右侧流程图占满整窗高度、不再被日志区拦腰
        截断。pack 顺序要先日志区后记事本：`pack` 是**按顺序切走空间**的，
        先切走底边给日志区，记事本再占剩下的。
        """
        self.paned = tk.PanedWindow(self, orient="horizontal", sashwidth=6,
                                    sashpad=0, sashrelief="raised", bd=0,
                                    bg=C_LINE, showhandle=False)
        self.paned.pack(side="left", fill="both", expand=True)

        nb_holder = tk.Frame(self.paned, bg=C_PANEL)
        self._build_bottom(nb_holder)      # 必须比 _build_notebook 先 pack
        self._build_notebook(nb_holder)
        self.paned.add(nb_holder, minsize=560, stretch="always")

        side = tk.Frame(self.paned, bg=C_APP)
        self._build_flow_side(side)
        self.paned.add(side, minsize=420, width=self._flow_width, stretch="never")
        self._guide_side = side
        self._build_guide_strip()

    # ------------------------------------------------------------ 向导栏折叠
    #
    # 用户 2026-09-20：非最大化时允许把右侧向导栏折起来（默认展开）。折叠＝
    # **窗口本身变窄**（右边缘向左收，减少的宽度就是向导栏占的那一块），左边
    # 的宽度一点不变——不是把腾出的宽度分给左边。折叠后窗口最右边留一条醒目的
    # "展开使用向导"竖条；点它，窗口向右变宽同样的量、向导栏原样回来。
    # 最大化时不折叠（没有"缩窗口"这回事）：折叠控件隐藏，已折叠的自动展开。

    _GUIDE_STRIP_W = 46

    def _build_guide_strip(self) -> None:
        self._guide_collapsed = False
        self._guide_w = self._flow_width
        self._last_win_state = "normal"
        # 只用字、不用色块（用户 2026-09-20）：跟窗口同色的窄条，左边一道分隔线，
        # 竖排的字用深色加粗；箭头指向"窗口边缘将要移动的方向"——展开＝窗口向右长，
        # 所以是 ▸。
        strip = tk.Frame(self, bg=C_APP, width=self._GUIDE_STRIP_W, cursor="hand2")
        strip.pack_propagate(False)
        self._guide_strip = strip
        tk.Frame(strip, bg=C_LINE, width=1).pack(side="left", fill="y")
        text = "▸\n\n展\n开\n使\n用\n向\n导"
        lbl = tk.Label(strip, text=text, font=(FAMILY, 11, "bold"), bg=C_APP,
                       fg=C_TEXT, cursor="hand2", justify="center")
        lbl.pack(side="left", expand=True)
        for w in (strip, lbl):
            w.bind("<Button-1>", lambda _e: self._expand_guide())
            w.bind("<Enter>", lambda _e: (strip.configure(bg=C_STRIP_HOT), lbl.configure(bg=C_STRIP_HOT)))
            w.bind("<Leave>", lambda _e: (strip.configure(bg=C_APP), lbl.configure(bg=C_APP)))
        self.bind("<Configure>", self._on_main_configure, add="+")
        self.btn_guide_collapse.place(relx=1.0, x=-6, rely=0.5, anchor="e")   # 最大化时会被收起

    def _on_main_configure(self, event=None) -> None:
        """窗口状态在"还原 ⇄ 最大化"之间切换时，调整向导栏的折叠控件。"""
        if event is not None and event.widget is not self:
            return
        state = self.state()
        if state == self._last_win_state:
            return
        self._last_win_state = state
        if state == "zoomed":
            self.btn_guide_collapse.place_forget()
            if self._guide_collapsed:
                self._expand_guide(resize=False)     # 最大化：只把栏摆回来，不改窗口大小
        elif state == "normal":
            self.btn_guide_collapse.place(relx=1.0, x=-6, rely=0.5, anchor="e")

    def _collapse_guide(self) -> None:
        if self._guide_collapsed or self.state() != "normal":
            return
        side = self._guide_side
        self.update_idletasks()
        self._guide_w = max(side.winfo_width(), 420)
        sash = int(str(self.paned.cget("sashwidth")))
        width_before, height = self.winfo_width(), self.winfo_height()
        self.paned.forget(side)
        self._guide_strip.pack(side="right", fill="y", before=self.paned)
        self._guide_collapsed = True
        # 窗口减掉"向导栏 + 分隔条 − 竖条"那么多，左边保持原宽
        self._set_window_size(width_before - (self._guide_w + sash - self._GUIDE_STRIP_W), height)

    def _expand_guide(self, resize: bool = True) -> None:
        if not self._guide_collapsed:
            return
        sash = int(str(self.paned.cget("sashwidth")))
        width_before, height = self.winfo_width(), self.winfo_height()
        self._guide_strip.pack_forget()
        self.paned.add(self._guide_side, minsize=420, width=self._guide_w, stretch="never")
        self._guide_collapsed = False
        if resize:
            grow = self._guide_w + sash - self._GUIDE_STRIP_W
            room = self.winfo_screenwidth() - self.winfo_x()      # 别把窗口撑出屏幕右边
            self._set_window_size(min(width_before + grow, max(room, width_before)), height)

    def _build_flow_side(self, side: tk.Frame) -> None:
        head = tk.Frame(side, bg=C_APP)
        head.pack(fill="x", padx=(4, 6), pady=(6, 0))
        # 标题行：标题对齐流程图的中轴（不是整栏居中），右边放"收起"按钮
        title_row = tk.Frame(head, bg=C_APP, height=44)
        title_row.pack(fill="x")
        title_row.pack_propagate(False)
        self._guide_title = tk.Label(title_row, text="使用向导", font=(FAMILY, 16, "bold"),
                                     bg=C_APP, fg=C_TEXT)
        self._guide_title.place(relx=0.5, rely=0.5, anchor="center")
        tk.Label(head, text="点击划线文本可直接跳转；Ctrl＋滚轮缩放，空白处双击复位",
                 font=(FAMILY, 9), bg=C_APP, fg=C_HINT).pack(fill="x", anchor="center")
        # 折叠按钮：只在窗口没最大化时出现（见 `_on_main_configure`）。箭头指向"窗口
        # 右边缘将要移动的方向"——收起＝窗口向左缩，所以是 ◂。
        self.btn_guide_collapse = tk.Button(
            title_row, text="收起向导 ◂", font=(FAMILY, 12, "bold"), relief="flat", bd=0,
            padx=16, pady=6, bg="#e4e4e0", activebackground="#d6d6d1", fg=C_TEXT,
            cursor="hand2", command=self._collapse_guide)

        self.flow = FlowPanel(side, on_activate=self._on_flow)
        self.flow.pack(fill="both", expand=True, padx=(4, 6), pady=6)
        # 流程图的位置/缩放/滚动一变，标题就重新对齐它的中轴
        self.flow.on_layout = self._align_guide_title
        title_row.bind("<Configure>", lambda _e: self._align_guide_title())
        self.after_idle(self._align_guide_title)

    def _align_guide_title(self) -> None:
        """把"使用向导"四个字摆到流程图中轴的正上方（标题行是它的位置基准）。"""
        row = self._guide_title.master
        if not row.winfo_exists() or row.winfo_width() < 20:
            return
        axis_root = self.flow.canvas.winfo_rootx() + self.flow.axis_view_x()
        x = axis_root - row.winfo_rootx()
        half = self._guide_title.winfo_reqwidth() / 2
        # 别压到右边的按钮、也别出左边界
        limit_right = row.winfo_width() - (self.btn_guide_collapse.winfo_reqwidth() + 16
                                            if self.btn_guide_collapse.winfo_ismapped() else 0) - half
        x = max(half + 4, min(x, limit_right))
        self._guide_title.place(x=x, rely=0.5, relx=0.0, anchor="center")

    def _build_notebook(self, master: tk.Frame) -> None:
        style = ttk.Style(self)
        style.configure("TNotebook.Tab", font=(FAMILY, 11), padding=(26, 8))
        self.nb = ttk.Notebook(master)
        self.nb.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
        # 切到「结构」页签时识别 + 报告警（方案 3，用户 2026-09-15 选定）
        self.nb.bind("<<NotebookTabChanged>>", self._on_tab_changed, add="+")

        tab1 = tk.Frame(self.nb, bg=C_PANEL)
        self.nb.add(tab1, text="文件")
        self._build_tab_files(tab1)

        tab_page = tk.Frame(self.nb, bg=C_PANEL)
        self.nb.add(tab_page, text="页面")
        self._build_tab_page(tab_page)

        tab_fmt = tk.Frame(self.nb, bg=C_PANEL)
        self.nb.add(tab_fmt, text="格式")
        self._build_tab_fmt(tab_fmt)

        tab_struct = tk.Frame(self.nb, bg=C_PANEL)
        self.nb.add(tab_struct, text="结构")
        self._build_tab_struct(tab_struct)

        tab_cover = tk.Frame(self.nb, bg=C_PANEL)
        self.nb.add(tab_cover, text="封面")
        self._build_tab_cover(tab_cover)

        tab_about = tk.Frame(self.nb, bg=C_PANEL)
        self.nb.add(tab_about, text="关于")
        self._build_tab_about(tab_about)

        # 页签顺序（用户 2026-09-20）：文件－结构－封面－格式－页面－关于。各页签是
        # 按上面的顺序建的（有的建页要用到前面已建好的东西），建完再挪位置。
        self.nb.insert(1, tab_struct)
        self.nb.insert(2, tab_cover)
        self.nb.insert(3, tab_fmt)

        # 「锁定窗口」：压在页签条最右边（"关于"右边），勾上＝窗口始终在最前面
        self.v_topmost = tk.BooleanVar(value=False)
        bar_bg = self._tabbar_bg()
        self._chk_lock = checkbutton(master, "锁定窗口", variable=self.v_topmost,
                                     command=self._toggle_topmost, font_size=10)
        self._chk_lock.configure(style=checkbutton_style_on(self, 10, bar_bg))
        self._chk_lock.place(in_=self.nb, relx=1.0, x=-12, y=8, anchor="ne")
        self.nb.bind("<Configure>", lambda _e: self._place_lock(), add="+")
        self.after_idle(self._place_lock)

    def _tabbar_bg(self) -> str:
        """页签条（TNotebook）的底色，取不到就用常见的浅灰。"""
        color = ttk.Style(self).lookup("TNotebook", "background") or "#f0f0f0"
        try:
            r, g, b = self.winfo_rgb(color)
            return f"#{r >> 8:02x}{g >> 8:02x}{b >> 8:02x}"
        except tk.TclError:
            return "#f0f0f0"

    def _place_lock(self) -> None:
        """让"锁定窗口"在页签条里上下居中（页签条高度＝第一个页面内容区的上沿）。"""
        strip_h = self.nb.winfo_children()[0].winfo_y() if self.nb.winfo_children() else 0
        h = self._chk_lock.winfo_reqheight()
        y = max(2, (strip_h - h) // 2) if strip_h > h else 4
        self._chk_lock.place(in_=self.nb, relx=1.0, x=-12, y=y, anchor="ne")

    def _toggle_topmost(self) -> None:
        self.attributes("-topmost", bool(self.v_topmost.get()))

    # ------------------------------------------------------------ 结构页签

    def _build_tab_struct(self, tab: tk.Frame) -> None:
        """结构页签＝直接内嵌"识别结果调整"这张表（2026-09-14 重做）。

        以前这里是一张独立的"部件启用"表格，跟"校对和调整识别结果"弹窗各管
        一份、靠单向同步对齐，还得先弹一段警告劝退用户别在这边直接改。现在
        直接把那张表嵌进来：启用勾选就是表格里每个块自己的表头勾选，不再
        需要一张对照表，也不用担心两边打架——只有一份状态，没有"另一边"。
        导入文件之后自动建好这张表，不需要单独一个"打开调整窗口"的按钮。

        **告警在这一页上打印**：导入文件时不识别，用户切到这一页才算"开始
        核对结构"，那一刻识别并逐条报告（方案 3，用户 2026-09-15 选定），
        见 `_on_tab_changed`。
        """
        self.struct_tab = tab
        self._rebuild_struct_adjust()

    def _rebuild_struct_adjust(self) -> None:
        """(重新)填"论文结构"页签——内嵌的调整表格从这里建。

        每次都是"对着当前这份文档重新建表"：部件启用的初始值按"预设∪识别到的
        部件"这条规则重定（`seed_enabled_from`）——不能沿用上一份文档、甚至
        上次启动程序时残留的状态。"这份文档的设定"（封面范围 / 导出时删不删
        封面 / 两个"自动重新编号"开关）也跟着回到默认值（新文档本来就该从
        默认开始；同一份文档的设定由用户自己在结构页上维持）。

        以前这里有一个 `reseed` 参数和一段"沿用上一份 rows"的分支（审计
        2026-09-15 核实：两个调用点都传 `reseed=True`，那段代码一次都没执行过，
        而 docstring 却承诺了"重绘时沿用当前 rows"）。删掉：真需要"重绘但保留
        调整"时，直接用 `self._adjust_instance.relayout()/draw()`，不要再重建。

        `self.doc_path is None`（还没导入文件）时不再显示"先导入文件"的
        占位提示——改显示一份空白模板骨架（`build_blank_rows()`）：没有正文，
        每个前/后置部件一个空位，只能点左侧勾选框决定"生成空白模板"要哪些
        部件，拖不动（本来也没有段落可拖）。用户 2026-09-14 要求：这样才能
        在导入文件之前就先把空白模板要用的部件定好，不用非得先随便导入一份
        文档才能碰这一页。
        """
        from essay_tool.gui.adjust_preview import (
            AdjustWindow, build_blank_rows, build_rows, seed_enabled_from,
        )
        for w in self.struct_tab.winfo_children():
            w.destroy()
        self._heading_scheme_applied = self.var_heading_scheme.get()

        # 部件启用的"基线"：**没应用任何预设**时用 STRUCTURE_PARTS 的"默认
        # 启用"列——那就是用户 2026-09-14 定的"不套用任何预设时的默认值"
        # （题目/副标题/摘要/目录/引言/结论 开，其余关）；应用了预设的，就
        # 拿预设那份覆盖上来（预设里没写到的部件仍有默认值兜底）。以前这里
        # 没预设时是 None，等于"预设没预期、文档里也没有的部件一律关"，
        # 连题目/摘要这种最基本的部件都没有默认值。
        preset_structure: dict[str, bool] = {
            n: on for n, on, _txt, _note in STRUCTURE_PARTS}
        if self.applied_preset is not None:
            try:
                from essay_tool.config.preset_io import load_preset
                preset_structure.update(load_preset(self.applied_preset).structure_required)
            except Exception as exc:                  # noqa: BLE001
                # 不能静默（审计 2026-09-15 核实）：预设文件被改坏/读不到时，
                # 这里原来 `pass`，结果结构页勾选项静默回到"不套预设"的默认
                # 值——用户看到的是"预设里明明勾了附录"，却找不到任何提示。
                self.status(f"读取预设「{self.applied_preset.stem}」失败，"
                            f"结构部件启用状态按默认值显示：{exc}")

        if self.doc_path is None:
            fresh_rows = build_blank_rows()
            fresh_index = None
        else:
            try:
                fresh_rows, fresh_index = build_rows(
                    self.doc_path, heading_scheme=self.var_heading_scheme.get())
            except Exception as exc:
                self._adjust_instance = None
                tk.Label(self.struct_tab, text=f"识别失败：{exc}", font=(FAMILY, 11),
                        bg=C_PANEL, fg="#b3261e", wraplength=700, justify="left"
                        ).pack(expand=True, pady=60)
                return
        seed_enabled_from(preset_structure, fresh_rows)

        self._adjust_instance = AdjustWindow(
            self.struct_tab, self.doc_path, initial_rows=fresh_rows,
            initial_index=fresh_index,
            preset_baseline=preset_structure,
            open_rules_callback=self.on_open_recognition_rules,
            insert_text=self._structure_insert_text(),
            reference_style=self._structure_reference_style(),
            heading_scheme=self.var_heading_scheme.get(),
            heading_scheme_callback=self._on_heading_scheme_restored)

    # ------------------------------------------------------------ 封面页签

    # 卡片顺序 = 界面显示顺序；形态名单本身以 `schema.COVER_FORM_NAMES` 为准
    # （下面这张表只补"样张图文件名"，`_check_cover_forms` 启动时对一下）。
    COVER_FORMS = (("毕业论文封面", "cover_thesis.png"),
                   ("学年论文封面", "cover_year.png"),
                   ("课程考核封面", "cover_course.png"),
                   ("学术十星封面", "cover_tenstar.png"))

    # 填写区分三组：毕业论文/学年论文同一组字段，课程考核一组，十星一组
    # （十星封面只有题目和副标题，不含个人信息）。形态名 → 字段组名。
    COVER_FIELD_GROUP = {"毕业论文封面": "thesis", "学年论文封面": "thesis",
                         "课程考核封面": "course", "学术十星封面": "tenstar"}

    # 学院下拉列表（用户 2026-09-13 指定的顺序）
    COLLEGE_OPTIONS: tuple[str, ...] = (
        "法学院", "民商经济法学院", "国际法学院/涉外法治学院", "刑事司法学院",
        "政治与公共管理学院", "商学院", "人文学院", "法律学院", "外国语学院",
        "国际教育学院", "法治信息管理学院", "社会学院", "中欧法学院",
        "马克思主义学院", "光明新闻传播学院", "比较法学院", "纪检监察学院",
        "海南国际学院", "国家安全学院/依法治疆研究院", "国际儒学院",
        "人权研究院（当代法治研究院）", "数据法学院（数据法治研究院）",
        "证据科学研究院", "其他（请直接输入）",
    )

    def _build_tab_cover(self, tab: tk.Frame) -> None:
        page = self._make_scroll_page(tab)
        self._cover_page = page
        col = tk.Frame(page, bg=C_PANEL)
        col.pack(fill="both", expand=True, padx=46, pady=(20, 14))

        top = tk.Frame(col, bg=C_PANEL)
        top.pack(fill="x", pady=(0, 12))
        tk.Label(top, text="自动封面", font=(FAMILY, 13, "bold"), bg=C_PANEL,
                 fg=C_TEXT).pack(side="left")
        self.lbl_cover_src = tk.Label(top, text="（未应用配置）", font=(FAMILY, 9),
                                      bg=C_PANEL, fg=C_HINT)
        self.lbl_cover_src.pack(side="left", padx=(10, 0))

        self.cv_cover: dict[str, tk.Variable] = {}

        def section(title: str) -> tuple[tk.Frame, tk.Frame]:
            box = tk.Frame(col, bg=C_PANEL, highlightbackground=C_LINE,
                           highlightthickness=1)
            box.pack(fill="x", pady=(0, 12))
            tk.Label(box, text=title, font=(FAMILY, 10, "bold"), bg=C_PANEL,
                     fg=C_TEXT).pack(anchor="w", padx=12, pady=(8, 0))
            inner = tk.Frame(box, bg=C_PANEL)
            inner.pack(fill="x", padx=12, pady=(6, 10))
            return box, inner

        def check(parent: tk.Frame, key: str, text: str, on: bool = False,
                  pady=(0, 0), command=None) -> tk.Checkbutton:
            self.cv_cover[key] = tk.BooleanVar(value=on)
            cb = self._checkbutton(parent, text, variable=self.cv_cover[key],
                                   command=command)
            cb.pack(anchor="w", pady=pady)
            return cb

        def hint(parent: tk.Frame, text: str, pady=(3, 0)) -> None:
            tk.Label(parent, text=text, font=(FAMILY, 9), bg=C_PANEL, fg=C_HINT,
                     wraplength=620, justify="left"
                     ).pack(anchor="w", pady=pady)

        # —— 设置（总开关 + 悬挂缩进并排；用户 2026-09-20：原来的「格式」框并进来，
        #    框标题「开关」改成「设置」）
        box1, s1 = section("设置")
        self.cover_sw_box = box1       # 流程图点「自动封面」闪的就是这个框
        # 默认不勾（用户 2026-09-14：不套用任何预设时"不使用封面"）。
        # 勾上/取消会走 _cover_power()，把下面整块内容一起点亮/变灰。
        power_cb = check(s1, "on", "为论文插入封面", on=False, command=self._cover_power)
        power_cb.pack_configure(side="left")
        # 总开关自己和这个框的标题不跟着变灰（不然关了就没法再打开）；同框里的
        # 悬挂缩进要跟着灰，所以 `_cover_dim_widgets` 只跳过这两个控件，不跳整个框。
        self._cover_no_dim = {box1.winfo_children()[0], power_cb}
        self._cover_switchables: list[tk.Widget] = []
        hanging_cb = check(s1, "hanging", "为超长内容启用悬挂缩进", on=True)
        hanging_cb.pack_configure(side="left", padx=(32, 0))
        self._cover_switchables.append(hanging_cb)

        # —— 形态
        _box2, s2 = section("形态")
        self.cv_cover["form"] = tk.StringVar(value="毕业论文封面")
        hint(s2, "双击打开大图。")
        cards = tk.Frame(s2, bg=C_PANEL)
        cards.pack(anchor="w", pady=(8, 0))
        self._cover_cards: dict[str, dict] = {}
        self._cover_dim_photos: dict[str, object] = {}   # 总开关关闭时的变灰版样张
        self._cover_photo_refs: dict[str, object] = {}   # 防 PhotoImage 被 GC
        # 四张卡片排成 2×2：左侧面板只有窗口宽度的一半多一点，一排放不下四张。
        for i, (form, img_name) in enumerate(self.COVER_FORMS):
            self._cover_cards[form] = self._mk_cover_card(
                cards, form, img_name, row=i // 2, column=i % 2)

        # —— 填写内容
        _box4, s4 = section("填写内容")
        hint(s4, "该内容将不会保存进预设。")
        self.cv_cover_fields: dict[str, tk.StringVar] = {}
        # 填写区的输入框，按界面上从上到下的顺序攒着——回车跳到下一栏用
        # （用户 2026-09-16：「希望可以允许在封面页面填写信息时使用回车
        # 切换到下一个栏目」）。存的是（所属那一组，控件）对，跳栏时只取
        # 当前显示的那一组。
        self._cover_inputs: list[tuple[tk.Frame, tk.Misc]] = []

        def field(parent: tk.Frame, key: str, label: str, hint_text: str = "",
                  width: int = 24) -> None:
            row = tk.Frame(parent, bg=C_PANEL)
            row.pack(fill="x", pady=(3, 0))
            tk.Label(row, text=label, font=(FAMILY, 10), bg=C_PANEL, fg=C_TEXT,
                     width=10, anchor="w").pack(side="left")
            # 同名字段共用一个变量：十星封面的"论文题目/副标题"和毕业论文封面
            # 的是同一份内容，切换形态时填过的不丢。
            var = self.cv_cover_fields.get(key) or tk.StringVar(value="")
            self.cv_cover_fields[key] = var
            ent = tk.Entry(row, textvariable=var, font=(FAMILY, 10), width=width)
            ent.pack(side="left")
            ent.bind("<Return>", self._cover_focus_next)
            self._cover_inputs.append((parent, ent))
            self._cover_switchables.append(ent)
            if hint_text:
                tk.Label(row, text=hint_text, font=(FAMILY, 9), bg=C_PANEL,
                         fg=C_HINT).pack(side="left", padx=(8, 0))

        def combo(parent: tk.Frame, key: str, label: str, values: tuple[str, ...],
                  width: int = 22) -> None:
            row = tk.Frame(parent, bg=C_PANEL)
            row.pack(fill="x", pady=(3, 0))
            tk.Label(row, text=label, font=(FAMILY, 10), bg=C_PANEL, fg=C_TEXT,
                     width=10, anchor="w").pack(side="left")
            var = tk.StringVar(value="")
            self.cv_cover_fields[key] = var
            cb = ttk.Combobox(row, textvariable=var, values=values, width=width,
                              font=(FAMILY, 10))
            cb.pack(side="left")
            cb.bind("<Return>", self._cover_focus_next)
            self._cover_inputs.append((parent, cb))
            self._lock_wheel(cb)
            self._cover_switchables.append(cb)

        self._cover_fields_thesis = tk.Frame(s4, bg=C_PANEL)
        for key, label, hint_text in (
            ("title", "论文题目", ""),
            ("subtitle", "副标题", "可留空"),
            ("student_name", "姓名", ""),
            ("student_id", "学号", ""),
        ):
            field(self._cover_fields_thesis, key, label, hint_text)
        combo(self._cover_fields_thesis, "college", "学院", self.COLLEGE_OPTIONS)
        for key, label, hint_text in (
            ("major", "专业", ""),
            ("grade", "年级", "如“2022级”"),
            ("date", "完成日期", ""),
            ("advisor_name", "指导教师姓名", ""),
            ("advisor_title", "指导教师职称", "可留空"),
            ("number", "论文编号", "通常留空，由学校填写"),
            ("score", "论文成绩", "通常留空，由老师填写"),
        ):
            field(self._cover_fields_thesis, key, label, hint_text)

        self._cover_fields_course = tk.Frame(s4, bg=C_PANEL)
        for key, label, hint_text in (
            ("course_name", "课程名称", ""),
            ("course_code", "课程代码", ""),
            ("course_seq", "课程序号", ""),
            ("teacher_name", "任课教师", "默认只写姓名，不带职称"),
        ):
            field(self._cover_fields_course, key, label, hint_text)

        # 十星封面只有题目和副标题：规范"封面及全文不能出现作者姓名、院系、
        # 指导老师姓名等个人信息"，"论文编号"栏保持空白、不给输入。
        self._cover_fields_tenstar = tk.Frame(s4, bg=C_PANEL)
        for key, label, hint_text in (
            ("title", "论文题目", ""),
            ("subtitle", "副标题", ""),
            ("session", "届数", "如“二十四”；同时影响封面和页眉"),
        ):
            field(self._cover_fields_tenstar, key, label, hint_text)
        # 届数默认二十四（范例就是第二十四届）；下一届来了改成“二十五”即可，
        # 不用等新预设（用户 2026-09-19："在 GUI 允许用户填写届数"）。
        self.cv_cover_fields["session"].set("二十四")

        self._cover_select("毕业论文封面")
        self._cover_power()
        wire = getattr(page, "_page_wire", None)
        if wire:
            wire()

    def _cover_focus_next(self, event=None):
        """封面填写区按回车：焦点跳到下一个栏目（到最后一行就绕回第一行）。

        用户 2026-09-16：「希望可以允许在封面页面填写信息时使用回车切换到
        下一个栏目。」只在本页签当前显示的那一组（论文封面 / 课程考核封面）
        里跳。

        ！判“哪一组”**不能**用 `winfo_ismapped()`：填写区自己也套了滚动页，
        字段被滚出视野时它会变成假，回车就默默失灵了。按形态取值更稳。
        """
        pairs = getattr(self, "_cover_inputs", [])
        if not pairs:
            return "break"
        form = ""
        try:
            form = str(self.cv_cover["form"].get())
        except Exception:                       # noqa: BLE001
            form = ""
        box = self._cover_field_box(form)
        inputs = [w for parent, w in pairs if parent is box]
        if not inputs:                          # 兜底：形态认不出就跳全部
            inputs = [w for _, w in pairs]
        current = getattr(event, "widget", None)
        try:
            index = inputs.index(current)
        except ValueError:
            index = -1
        nxt = inputs[(index + 1) % len(inputs)]
        nxt.focus_set()
        if isinstance(nxt, tk.Entry):           # 选中已有内容，直接打字就替换
            nxt.selection_range(0, "end")
        return "break"

    def _mk_cover_card(self, parent: tk.Frame, form: str, img_name: str,
                       row: int = 0, column: int = 0) -> dict:
        """一张样张卡片：图片＋标题（点＝选形态，双击＝开大图）。"""
        card = tk.Frame(parent, bg=C_PANEL, highlightthickness=2,
                        highlightbackground=C_PANEL)
        card.grid(row=row, column=column, padx=(0, 26), pady=(0, 14), sticky="nw")
        photo = self._cover_photo(img_name, COVER_CARD_W)
        if photo is not None:
            pic = tk.Label(card, image=photo, bd=0, bg=C_PANEL, cursor="hand2",
                           highlightbackground=C_LINE, highlightthickness=1)
            pic.image = photo                    # 标签上再存一份，双保险
            self._cover_photo_refs[form] = photo
        else:
            pic = tk.Label(card, width=28, height=12, font=(FAMILY, 9),
                           text="（样张图未生成：\n跑 tools\\make_cover_preview.py）",
                           bg="#f7f7f5", fg=C_HINT, cursor="hand2",
                           highlightbackground=C_LINE, highlightthickness=1)
        pic.pack()
        cap = tk.Label(card, text=form, font=(FAMILY, 9), bg=C_PANEL, fg=C_HINT)
        cap.pack(pady=(4, 0))
        for w in (pic, cap):
            w.bind("<Button-1>", lambda e, f=form: self._cover_card_click(f))
            w.bind("<Double-Button-1>",
                   lambda e, n=img_name, t=form: self._open_cover_zoom(n, t))
        return {"card": card, "cap": cap, "img": img_name, "pic": pic, "w": COVER_CARD_W}

    def _cover_card_click(self, form: str) -> None:
        """样张卡片单击＝选形态。

        总开关关着时忽略：卡片是 Label，没法像别的控件那样 `state=disabled`，
        只能在回调里拦一道（`_cover_power` 会把鼠标指针恢复成箭头，看着就不像
        能点的）。双击开大图不受影响——那只是看样张，不改任何设置。
        """
        if not self.cv_cover["on"].get():
            self.status("封面总开关未开启，请先勾选为论文插入封面。")
            return
        self._cover_select(form)

    def _cover_select(self, form: str | None = None) -> None:
        """选封面形态：radio 值和卡片高亮（蓝边＋标题加粗）保持同一份状态。"""
        cards = getattr(self, "_cover_cards", {})
        if form is None:
            form = self.cv_cover["form"].get()
        self.cv_cover["form"].set(form)
        for name, c in cards.items():
            on = name == form
            c["card"].configure(highlightbackground=C_LINK if on else C_PANEL)
            c["cap"].configure(fg=C_TEXT if on else C_HINT,
                               font=(FAMILY, 9, "bold" if on else "normal"))
        if hasattr(self, "_cover_fields_tenstar"):      # 三组字段框都建好了
            show = self._cover_field_box(form)
            for box in (self._cover_fields_thesis, self._cover_fields_course,
                        self._cover_fields_tenstar):
                if box is show:
                    box.pack(anchor="w", pady=(6, 0))
                else:
                    box.pack_forget()

    def _cover_field_box(self, form: str) -> tk.Frame:
        """这个封面形态对应的填写字段框（毕业论文/学年论文同一组）。"""
        group = self.COVER_FIELD_GROUP.get(form, "thesis")
        return {"thesis": self._cover_fields_thesis,
                "course": self._cover_fields_course,
                "tenstar": self._cover_fields_tenstar}[group]

    def _cover_power(self) -> None:
        """总开关关掉时，把形态／选项区的控件一起锁上（不参与的设置一眼可见）。

        卡片（样张图）是 Label，没有 `state` 可设，所以另做两件事：把鼠标指针
        恢复成箭头（看着就不像能点）、点击由 `_cover_card_click` 拦掉。

        除此之外，**复选框以下的全部内容（区块标题、提示语、填写框、样张图）
        在关闭时统一变灰**（走 `_dim_cover_content`）——只"点不动"而不变色，
        用户看不出这一整块当前不生效（2026-09-14 反馈"那个封面只是不能选了，
        不是灰了"）。
        """
        on = bool(self.cv_cover["on"].get())
        state = "normal" if on else "disabled"
        for w in getattr(self, "_cover_switchables", []):
            try:
                w.configure(state=state)
            except Exception:
                pass
        for c in getattr(self, "_cover_cards", {}).values():
            for w in (c.get("pic"), c.get("cap")):
                if w is not None:
                    try:
                        w.configure(cursor="hand2" if on else "")
                    except Exception:
                        pass
        self._dim_cover_content(not on)
        if on:
            # 变灰期间可能换过配置（_load_cover_from → _cover_select），恢复时
            # 让"选中形态"的高亮/加粗按当前值重画一遍，别用缓存里的旧颜色。
            self._cover_select(None)

    def _cover_dim_widgets(self) -> list:
        """封面页签里"跟着总开关一起变灰"的控件：总开关自己和「设置」框标题以外的全部。

        递归走完整棵控件树（含纯文字 Label），所以区块标题、提示语、填写框、
        样张卡片、同框的悬挂缩进复选框都在内；总开关复选框和「设置」标题
        （`_cover_no_dim`）跳过——不然连开关都灰了就没法再把它打开。
        """
        skip = getattr(self, "_cover_no_dim", set())
        out: list = []

        def walk(w) -> None:
            if w in skip:
                return
            out.append(w)
            for ch in w.winfo_children():
                walk(ch)

        page = getattr(self, "_cover_page", None)
        if page is not None:
            walk(page)
        return out

    def _dim_cover_content(self, dim: bool) -> None:
        """把封面页签的文字/样张图统一变灰（dim=True）或恢复原色。

        原色在第一次变灰时记进 `_cover_fg_cache`（键 = 控件路径 + 选项名），
        恢复时照原样写回——不靠"猜原来该是什么颜色"，免得跟 `_cover_select`
        那套"选中=蓝边+黑字加粗"的配色打架。
        """
        cache = getattr(self, "_cover_fg_cache", None)
        if cache is None:
            cache = self._cover_fg_cache = {}
        for w in self._cover_dim_widgets():
            # 1) 经典 tk 控件：直接改文字色。连 disabledforeground 一起改——
            #    输入框被 disable 之后走的是它，只改 fg 看不出变灰。
            for opt in ("fg", "disabledforeground"):
                try:
                    original = w.cget(opt)
                except Exception:
                    continue
                key = (str(w), opt)
                try:
                    if dim:
                        cache.setdefault(key, original)
                        w.configure(**{opt: C_DIM})
                    elif key in cache:
                        w.configure(**{opt: cache.pop(key)})
                except Exception:
                    pass
            # 2) ttk 控件（复选框/下拉框）的文字色由样式决定，cget("fg") 拿不到
            #    ——临时套一个 "Dim." 派生样式，恢复时写回原样式名。
            try:
                style_name = w.cget("style")
            except Exception:
                continue
            if not style_name:
                continue
            key = (str(w), "style")
            try:
                if dim:
                    cache.setdefault(key, style_name)
                    dim_style = "Dim." + style_name
                    ttk.Style(self).configure(dim_style, foreground=C_DIM,
                                              background=C_PANEL)
                    w.configure(style=dim_style)
                elif key in cache:
                    w.configure(style=cache.pop(key))
            except Exception:
                pass
        # 3) 样张卡片：边框 + 图片。tk 改不了已有图片的像素，所以另烘一张
        #    "去饱和+提亮"的版本顶着（没有 Pillow 就只变边框）。
        for form, c in getattr(self, "_cover_cards", {}).items():
            card = c.get("card")
            if card is not None:
                try:
                    card.configure(highlightbackground=C_LINE if dim else C_PANEL)
                except Exception:
                    pass
            pic = c.get("pic")
            if pic is None or c.get("img") is None:
                continue
            try:
                if dim:
                    photo = self._cover_dim_photos.get(form)
                    if photo is None:
                        photo = self._cover_photo_dim(c["img"],
                                                      c.get("w", COVER_CARD_W))
                        if photo is None:
                            continue        # 没 Pillow：保留原图，边框已变灰
                        self._cover_dim_photos[form] = photo
                    pic.configure(image=photo)
                    pic.image = photo
                else:
                    original = self._cover_photo_refs.get(form)
                    if original is not None:
                        pic.configure(image=original)
                        pic.image = original
            except Exception:
                pass

    def _cover_photo_dim(self, img_name: str, width: int):
        """样张图的"变灰"版：去饱和 + 往白里提亮（总开关关闭时用）。"""
        path = COVER_ASSETS_DIR / img_name
        if not path.exists():
            return None
        try:
            from PIL import Image, ImageEnhance, ImageTk
            with Image.open(path) as im:
                im = im.convert("RGB")
                im = im.resize((width, round(im.height * width / im.width)),
                               Image.LANCZOS)
            im = ImageEnhance.Color(im).enhance(0.0)
            im = Image.blend(im, Image.new("RGB", im.size, "#ffffff"), 0.6)
            return ImageTk.PhotoImage(im)
        except Exception:
            return None

    def _cover_photo(self, img_name: str, width: int):
        """assets/ 里的封面 PNG → 缩到指定宽的 PhotoImage（没图返回 None）。"""
        path = COVER_ASSETS_DIR / img_name
        if not path.exists():
            return None
        try:
            from PIL import Image, ImageTk
            im = Image.open(path)
            im = im.resize((width, round(im.height * width / im.width)),
                           Image.LANCZOS)
            return ImageTk.PhotoImage(im)
        except Exception:
            pass
        try:                        # 没有 Pillow 的退路：Tk 自带 PNG 支持，整数倍缩小
            img = tk.PhotoImage(file=str(path))
            return img.subsample(max(1, round(img.width() / width)))
        except Exception:
            return None

    def _open_cover_zoom(self, img_name: str, title: str) -> None:
        """样张大图：独立窗口，图按屏幕高限幅；点画面／Esc 关。"""
        path = COVER_ASSETS_DIR / img_name
        if not path.exists():
            self.status("封面样张加载失败：请先运行 tools\\make_cover_preview.py 生成样张。")
            return
        try:
            from PIL import Image
            with Image.open(path) as im:
                ratio = im.height / im.width
        except Exception:
            ratio = 1.414
        width = int(self.winfo_screenheight() * 0.82 / ratio)
        photo = self._cover_photo(img_name, max(300, min(width, 1300)))
        if photo is None:
            self.status("封面样张加载失败：请先运行 tools\\make_cover_preview.py 重新生成样张。")
            return
        win = tk.Toplevel(self)
        win.title(f"{title} —— 样张")
        win.configure(bg="#4a4a4a")
        lbl = tk.Label(win, image=photo, bd=0, bg="#4a4a4a")
        lbl.image = photo
        lbl.pack(padx=14, pady=(14, 6))
        tk.Label(win, text="点画面任意处，或按 Esc 关闭", font=(FAMILY, 9),
                 bg="#4a4a4a", fg="#d8d8d8").pack(pady=(0, 10))
        win.bind("<Button-1>", lambda e: win.destroy())
        win.bind("<Escape>", lambda e: win.destroy())
        win.transient(self)
        # 样张窗口摆在主窗口正中央（用户 2026-09-15：二级窗口一律居中，
        # 不许靠屏幕/父窗口左上角）。
        center_on(win, self)

    def _load_cover_from(self, cfg, label: str) -> None:
        """预设「封面」小节 → 封面页签（悬挂缩进开关；开关与形态）。

        2026-09-14 起预设能自带"要不要封面、用哪种封面"（`cover_enabled` /
        `cover_form`，见 `schema.PresetConfig`）：课程考核那份预设就是靠
        `cover_form="课程考核封面"` 把形态定下来的。旧预设里没这两项时保持
        老行为——用封面，形态按预设名里有没有"课程"猜。
        封面里填写的内容不归预设管。
        """
        self.lbl_cover_src.config(text=f"当前配置：{label}")
        self.cv_cover["hanging"].set(bool(getattr(cfg, "cover_hanging_indent", True)))
        form = getattr(cfg, "cover_form", None) or (
            "课程考核封面" if "课程" in label else
            "学年论文封面" if "学年" in label else "毕业论文封面")
        if form not in self.COVER_FIELD_GROUP:      # 预设里写了不认识的形态名
            form = "毕业论文封面"
        self._cover_select(form)
        enabled = getattr(cfg, "cover_enabled", None)
        self.cv_cover["on"].set(True if enabled is None else bool(enabled))
        self._cover_power()

    # ------------------------------------------------------------ 关于页签

    def _build_tab_about(self, tab: tk.Frame) -> None:
        page = self._make_scroll_page(tab)
        col = tk.Frame(page, bg=C_PANEL)
        col.pack(fill="both", expand=True, padx=46, pady=(56, 14))

        def part() -> tk.Frame:
            box = tk.Frame(col, bg=C_PANEL)
            box.pack(pady=(0, 40))
            return box

        def label(parent, text, pady=(0, 0), **kw) -> None:
            kw.setdefault("font", (FAMILY, 10))
            kw.setdefault("bg", C_PANEL)
            kw.setdefault("fg", C_TEXT)
            tk.Label(parent, text=text, **kw).pack(anchor="center", pady=pady)

        def button(parent, text, cmd, **pack_kw) -> tk.Button:
            b = tk.Button(parent, text=text, font=(FAMILY, 11), padx=14, pady=5,
                         command=cmd)
            b.pack(anchor="center", **pack_kw)
            return b

        # —— 第一部分：使用协议
        p1 = part()
        label(p1, "如果使用本工具，即代表你已阅读并同意")
        button(p1, "关于使用本工具的协议", self.on_open_agreement, pady=(10, 0))

        # —— 第二部分：规范依据、操作逻辑
        # 前两行文字共用一组按钮、加上十星那行一共四个按钮（用户 2026-09-19 晚：
        # "前两行文字对一行按钮"）。用户定稿：① 规范按钮保持按钮内折行（两行）；
        # ② 另存按钮字号不许跟着缩（10＝旁标签字号）；③ 高度对齐不靠缩字号凑
        # （6pt 太小、被否），靠 sticky="nsew"——grid 直接把四个按钮拉成一样高；
        # ④ 跨两行的那组正好铺满"两行提示词＋行距"的总高——按钮不矮于两行提示词。
        # 前两行提示词合成一张两行标签（两行相距＝字体行距，比各占一行近 6px）；
        # 两行按钮同宽。
        p2 = part()
        grid = tk.Frame(p2, bg=C_PANEL)
        grid.pack(anchor="center")

        def row_label(r: int, text: str, span: int = 1) -> None:
            tk.Label(grid, text=text, font=(FAMILY, 10), bg=C_PANEL, fg=C_TEXT,
                     justify="right").grid(row=r, column=0, rowspan=span,
                                           sticky="e", padx=(0, 14))

        def button_pair(r: int, spec_text: str, open_cmd, save_cmd,
                        span: int = 1) -> None:
            tk.Button(grid, text=spec_text, font=(FAMILY, 8), padx=8, pady=3,
                      command=open_cmd).grid(row=r, column=1, rowspan=span,
                                             padx=(0, 10), sticky="nsew")
            tk.Button(grid, text="另存官方版", font=(FAMILY, 10), padx=14, pady=5,
                      command=save_cmd).grid(row=r, column=2, rowspan=span,
                                             sticky="nsew")

        row_label(0, "毕业论文预设依据\n课程考核预设参考", span=2)
        button_pair(0, "中国政法大学\n本科毕业论文撰写格式规范",
                    self.on_open_spec, self.on_save_official_spec, span=2)
        row_label(2, "学术十星预设依据")
        button_pair(2, "第二十四届学术十星论文大赛\n格式要求",
                    self.on_open_star10_spec, self.on_save_official_star10)
        label(p2, "如果你遇到了任何关于本工具的问题并需要排查，你可以", pady=(20, 0))
        button(p2, "查阅本工具操作逻辑说明", self.on_open_logic, pady=(10, 0))

        # —— 第三部分：版本信息
        p3 = part()
        label(p3, "CUPL本科论文自动排版工具", font=(FAMILY, 13, "bold"))
        # 署名：只对 @Grandliu66 划线和可点（点开跳开发者主页，链接见模块顶部
        # 的 DEVELOPER_HOMEPAGE）；"Developed by" 不划线、不可点
        # （2026-09-19 用户要求）。
        credit = tk.Frame(p3, bg=C_PANEL)
        credit.pack(pady=(8, 0))
        tk.Label(credit, text="Developed by", font=(FAMILY, 10), bg=C_PANEL,
                 fg=C_TEXT).pack(side="left", padx=(0, 5))
        credit_link = tk.Label(credit, text="@Grandliu66",
                               font=(FAMILY, 10, "underline"), bg=C_PANEL,
                               fg=C_LINK, cursor="hand2")
        credit_link.pack(side="left")
        credit_link.bind("<Button-1>",
                         lambda _e: self.on_open_developer_homepage())
        credit_row = tk.Frame(p3, bg=C_PANEL)
        credit_row.pack(anchor="center", pady=(4, 0))
        tk.Label(credit_row, text="With the assistance of", font=(FAMILY, 10),
                bg=C_PANEL, fg=C_TEXT).pack(side="left", padx=(0, 12))
        # 两行名字合成一张两行标签（行距＝字体行距，比两张单行标签各占一行
        # 近 ~6px；2026-09-19 用户"再靠近一点"）。justify="center" 保持居中观感。
        tk.Label(credit_row, text="Claude Sonnet 5\nDeepSeek V4.1 Flash",
                 font=(FAMILY, 10), bg=C_PANEL, fg=C_TEXT,
                 justify="center").pack(side="left")
        label(p3, f"v{__version__}", pady=(10, 0))      # 版本号唯一出处：essay_tool/__init__.py
        # 「更新日志」「开发者的话」：小字、划线、可点——点开文字窗口
        # （正文见模块顶部的 CHANGELOG_TEXT / DEVELOPER_WORDS；2026-09-19/20 用户要求）。
        log = tk.Label(p3, text="更新日志", font=(FAMILY, 9, "underline"),
                       bg=C_PANEL, fg=C_LINK, cursor="hand2")
        log.pack(pady=(6, 0))
        log.bind("<Button-1>", lambda _e: self.on_open_changelog())
        words = tk.Label(p3, text="开发者的话", font=(FAMILY, 9, "underline"),
                         bg=C_PANEL, fg=C_LINK, cursor="hand2")
        words.pack(pady=(4, 0))
        words.bind("<Button-1>", lambda _e: self.on_open_developer_words())

        # —— 本工具自身的版权声明（MIT；点开看全文）
        label(p3, "本工具版权声明", font=(FAMILY, 10, "bold"), pady=(28, 10))
        own_name, own_license, own_text = _OWN_LICENSE
        own = tk.Label(p3, text=f"{own_name} —— {own_license}",
                       font=(FAMILY, 10, "underline"), bg=C_PANEL, fg=C_LINK,
                       cursor="hand2")
        own.pack(anchor="center", pady=(0, 6))
        own.bind("<Button-1>", lambda _e: self._open_license_popup(
            own_name, own_license, own_text + _OWN_LICENSE_NOTE))

        # —— 第三方依赖库版权声明（各是一个链接，点开弹窗看原文）
        label(p3, "第三方依赖库版权声明", font=(FAMILY, 10, "bold"), pady=(24, 10))
        for name, license_name, text in _THIRD_PARTY_NOTICES:
            lnk = tk.Label(p3, text=f"{name} —— {license_name}",
                          font=(FAMILY, 10, "underline"), bg=C_PANEL, fg=C_LINK,
                          cursor="hand2")
            lnk.pack(anchor="center", pady=(0, 6))
            lnk.bind("<Button-1>",
                    lambda e, n=name, l=license_name, t=text:
                    self._open_license_popup(n, l, t))

        wire = getattr(page, "_page_wire", None)
        if wire:
            wire()

    def _open_text_popup(self, title: str, text: str, font_size: int = 9,
                         width: int = 760, height: int = 620,
                         formatted: bool = False, markup: bool = False) -> None:
        """一个带滚动条的只读文本弹窗——协议正文、第三方许可证原文共用这个。

        `formatted=True` 是「开发者的话」那种信式排版（用户 2026-09-19）：
        每个自然段首行缩进两个汉字（折行只在第一行少算这两个字的宽），
        "@" 开头的行去掉 "@"、居中加粗当小标题；窗口缩放重排后 tag 重打。

        `markup=True`（协议弹窗用，2026-09-20）：行首 "@" 是居中加粗的大标题，
        `**…**` 之间的字加粗；其余排版不变（不缩进）。

        `font_size` 只影响正文，标题字号跟着同比例放大，不传就是原来的大小。
        这里原来还有一个 `center: bool = True` 参数，2026-09-15 审计核实：**全仓
        0 个调用点传它**（历史遗留——“二级窗口不允许在屏幕左上角弹出”，现在一律
        摆主窗口正中央，那个开关早就没有作用了），跟 `status(log=)` 一样删掉。

        换行由我们自己算（跟日志区同一套 `text_wrap.wrap_text`）：Tk 的
        `wrap="word"` 在中文里会把整句当一个词，把“1.”这种编号单独甩在一行
        （用户 2026-09-19 实拍“每一条的 1. 2. 后换行”），所以正文先按窗口
        宽度折好再插进 Text，窗口尺寸变了重排一次。
        """
        win = tk.Toplevel(self)
        win.title(title)
        win.configure(bg=C_PANEL)
        win.geometry(f"{width}x{height}")
        center_on(win, self, width, height)
        tk.Label(win, text=title, font=(FAMILY, font_size + 3, "bold"),
                 bg=C_PANEL, fg=C_TEXT).pack(anchor="w", padx=16, pady=(16, 8))
        box = tk.Frame(win, bg=C_PANEL)
        box.pack(fill="both", expand=True, padx=16)
        sb = tk.Scrollbar(box)
        sb.pack(side="right", fill="y")
        t = tk.Text(box, wrap="none", bd=0, bg=C_PANEL, fg=C_TEXT, font=(FAMILY, font_size),
                   yscrollcommand=sb.set, highlightthickness=0)
        t.pack(fill="both", expand=True)
        sb.config(command=t.yview)
        # 信式排版的小标题样式（formatted=True 时给 "@" 行打这个 tag）
        t.tag_configure("devh", justify="center",
                        font=(FAMILY, font_size, "bold"),
                        spacing1=12, spacing3=6)
        t.tag_configure("bold", font=(FAMILY, font_size, "bold"))

        # —— 折行：不用 Tk 自带（理由见 docstring），跟日志区同一套算法
        font = tkfont.Font(font=t.cget("font"))
        cache: dict[str, float] = {}

        def measure(part: str) -> float:
            total = 0.0
            for ch in part:
                got = cache.get(ch)
                if got is None:
                    got = float(font.measure(ch))
                    cache[ch] = got
                total += got
            return total

        # 待排段落：[(要显示的正文, 是不是居中加粗小标题)]。formatted 时把 "@"
        # 行转标题、其余段首加"　　"（首行缩进两个汉字——折行第一行自动少算
        # 这两个字的宽）。
        # 段落种类：head＝居中加粗小标题（"@"开头）；right＝靠右对齐的落款（">"开头，
        # 不缩进）；body＝普通自然段。
        def split_bold(p: str) -> tuple[str, list[tuple[int, int]]]:
            """`**…**` 标出的加粗段：返回去掉标记后的文字和各段 (起, 止) 字符位置。"""
            if not markup or "**" not in p:
                return p, []
            plain: list[str] = []
            spans: list[tuple[int, int]] = []
            pos = 0
            for i, seg in enumerate(p.split("**")):
                if i % 2 == 1 and seg:
                    spans.append((pos, pos + len(seg)))
                plain.append(seg)
                pos += len(seg)
            return "".join(plain), spans

        paras: list[tuple[str, str, list[tuple[int, int]]]] = []
        for p in text.split("\n"):
            if formatted:
                if p.startswith("@"):
                    paras.append((p[1:], "head", []))
                elif p.startswith(">"):
                    paras.append((p[1:], "right", []))
                else:
                    paras.append((("　　" + p) if p else "", "body", []))
            elif markup and p.startswith("@"):
                paras.append((p[1:], "head", []))
            else:
                plain, spans = split_bold(p)
                paras.append((plain, "body", spans))
        t.tag_configure("devr", justify="right", spacing1=6)

        def rewrap(avail: float) -> None:
            if avail <= 0:
                return
            first = t.yview()[0]
            lines: list[str] = []
            tagged: list[tuple[str, int, int]] = []      # (tag, 起始行号 0 基, 行数)
            bolds: list[tuple[int, int, int]] = []       # (行号 0 基, 起列, 止列)
            for para, kind, spans in paras:
                wrapped = wrap_text(para, measure, avail) if para else [""]
                if spans:                                # 把整段里的加粗范围映射到折好的各行
                    cursor = 0
                    for li, ln in enumerate(wrapped):
                        st = para.find(ln, cursor) if ln else cursor
                        st = cursor if st < 0 else st
                        cursor = st + len(ln)
                        for lo0, hi0 in spans:
                            lo, hi = max(lo0, st), min(hi0, cursor)
                            if lo < hi:
                                bolds.append((len(lines) + li, lo - st, hi - st))
                if kind == "head":
                    tagged.append(("devh", len(lines), len(wrapped)))
                elif kind == "right":
                    tagged.append(("devr", len(lines), len(wrapped)))
                lines.extend(wrapped)
            t.configure(state="normal")
            t.delete("1.0", "end")
            t.insert("1.0", "\n".join(lines))
            for tag, start, count in tagged:               # 重排后 tag 要重打
                t.tag_add(tag, f"{start + 1}.0", f"{start + count}.end")
            for ln, c0, c1 in bolds:
                t.tag_add("bold", f"{ln + 1}.{c0}", f"{ln + 1}.{c1}")
            t.configure(state="disabled")
            t.yview_moveto(first)

        job: list[str | None] = [None]

        def render() -> None:
            job[0] = None
            rewrap(t.winfo_width() - LOG_WRAP_MARGIN)

        def on_text_configure(event) -> None:
            if event.widget is not t:
                return
            if job[0] is not None:
                try:
                    t.after_cancel(job[0])
                except Exception:              # noqa: BLE001
                    pass
            job[0] = t.after(80, render)

        t.bind("<Configure>", on_text_configure, add="+")
        # 首屏：窗口还没映射（winfo_width 恒为 1）时先按目标宽估一遍，
        # 映射后 <Configure> 到了再按真实宽度重排
        rewrap(width - 16 * 2 - 20 - LOG_WRAP_MARGIN)
        tk.Button(win, text="关闭", font=(FAMILY, 10), command=win.destroy
                 ).pack(pady=10)

    def _open_license_popup(self, name: str, license_name: str, text: str) -> None:
        self._open_text_popup(f"{name} —— {license_name}", text)

    def on_open_agreement(self) -> None:
        self._open_text_popup("关于使用本工具的协议", AGREEMENT_TEXT, markup=True)

    def on_open_logic(self) -> None:
        """打开「工具操作逻辑」说明，跟查阅《规范》原文一样是左树右文的结构。"""
        from essay_tool.gui.logic_viewer import open_logic_viewer
        open_logic_viewer(self, OPERATION_LOGIC_TEXT)

    def on_open_recognition_rules(self) -> None:
        """结构页底栏「查阅识别规则」——**跟关于页的「查阅本工具操作逻辑说明」
        同一个窗口**（用户 2026-09-15 要求："直接 link 到关于页'查阅操作逻辑说明'
        相同的窗口"）。

        用户 2026-09-17：“这个按钮自动跳到说明的 3.6”（3.6 识别，就是识别规则
        那三张表所在的小节）——`anchor="3.6"`；窗口已经开着时会在那个窗口里
        就跳过去（见 `open_logic_viewer`）。
        """
        from essay_tool.gui.logic_viewer import open_logic_viewer
        open_logic_viewer(self, OPERATION_LOGIC_TEXT, anchor="3.6")

    def on_open_spec(self) -> None:
        """「关于」里的《规范》原文浏览按钮（复用 spec_viewer，重复点会唤醒）。"""
        try:
            from essay_tool.gui.spec_viewer import open_spec_viewer
        except Exception as exc:
            self.status(f"规范浏览窗口打开失败：{exc}")
            return
        win = open_spec_viewer(self)
        if win is not None:
            return

    def on_save_official_spec(self) -> None:
        """「关于」里的「另存官方版」：把《规范》原始文件另存一份（不是从网上下载）。"""
        try:
            from essay_tool.gui.spec_viewer import find_default_spec
            spec = find_default_spec()
        except Exception:
            spec = None
        if spec is None or not spec.exists():
            # 2026-09-14 用户要求：跟规范窗口打开失败一个说法，把原因写在
            # 冒号后面，不再单独一句
            self.status("规范文件另存失败：reference 目录中未找到原始文件。")
            return
        try:
            verify_official(spec)
        except IntegrityError as exc:
            self.status(f"规范文件另存失败：{exc}")
            return
        out = filedialog.asksaveasfilename(
            title="另存官方版《规范》原文", defaultextension=spec.suffix,
            initialfile=spec.name)
        if not out:
            return
        try:
            shutil.copyfile(spec, out)
        except Exception as exc:
            self.status(f"另存失败：{exc}")
            return
        self.status(f"已另存：{pathlib.Path(out).name}")

    def on_open_star10_spec(self) -> None:
        """「关于」里学术十星那行的格式要求按钮：打开内置的原文浏览窗口。

        文本是开发期从官方 PDF 抽好的（`star10_spec_text.py`，运行时零依赖），
        跟《规范》原文浏览同一个窗口；PDF 原件另有「另存官方版」可拿。
        """
        try:
            from essay_tool.gui.spec_viewer import open_star10_viewer
        except Exception as exc:
            self.status(f"格式要求打开失败：{exc}")
            return
        open_star10_viewer(self)

    def on_save_official_star10(self) -> None:
        """「关于」里学术十星那行的「另存官方版」：把格式要求原件另存一份。"""
        try:
            from essay_tool.gui.spec_viewer import find_star10_spec
            spec = find_star10_spec()
        except Exception:
            spec = None
        if spec is None or not spec.exists():
            self.status("格式要求另存失败：reference 目录中未找到原始文件。")
            return
        try:
            verify_official(spec)
        except IntegrityError as exc:
            self.status(f"格式要求另存失败：{exc}")
            return
        out = filedialog.asksaveasfilename(
            title="另存官方版《学术十星格式要求》", defaultextension=spec.suffix,
            initialfile=spec.name)
        if not out:
            return
        try:
            shutil.copyfile(spec, out)
        except Exception as exc:
            self.status(f"另存失败：{exc}")
            return
        self.status(f"已另存：{pathlib.Path(out).name}")

    # ------------------------------------------------------------ 第一个页签

    def _build_tab_files(self, tab: tk.Frame) -> None:
        # 2026-09-14：本页也套滚动页。底部日志区展开后会占掉一块高度，
        # 把「导出论文到……」按钮压到可视区外面（用户实测反馈）。用户的意思
        # 是**不用**让其他部件上抬，允许挡住、能滚出来就行；其余五个页签
        # 早就套了滚动页，只剩这一页没有。
        page = self._make_scroll_page(tab)
        col = tk.Frame(page, bg=C_PANEL)
        col.pack(fill="both", expand=True, padx=42, pady=18)
        # 纵向排版：四块内容 + 三条“弹性空白”（用户 2026-09-17：“现在第一页头重
        # 脚轻……把空白在各个部件之间均一下，尤其是最下面的空白太大了”）。每块
        # 单独装一个 Frame，块与块之间放一个 `weight=1` 的空白行 ⇒ 视口比内容
        # 高出来的那部分被**三等分**摊到三处（含最下面那条），不会全堆在底部。
        # 视口比内容矮时 `weight` 拿不到空间、三行空白自然压成 0，滚动照旧
        # （`_make_scroll_page` 会把内层撑到视口高，空间就是这么来的）。
        col.grid_columnconfigure(0, weight=1)
        for _gap_row in (1, 3, 5):
            col.grid_rowconfigure(_gap_row, weight=1)
        blk_import = tk.Frame(col, bg=C_PANEL)
        blk_preset = tk.Frame(col, bg=C_PANEL)
        blk_backup = tk.Frame(col, bg=C_PANEL)
        blk_export = tk.Frame(col, bg=C_PANEL)
        # 四块和它们之间的三条空白都记下来（自检要量位置；也就四行，不包函数）
        self.lay_files_col = col
        self.lay_files_blocks = (blk_import, blk_preset, blk_backup, blk_export)
        for _index, _block in enumerate(self.lay_files_blocks):
            _block.grid(row=_index * 2, column=0, sticky="ew")

        # 导入文件
        row = tk.Frame(blk_import, bg=C_PANEL)
        row.pack(anchor="center", pady=(6, 6))
        self.btn_import = tk.Button(row, text="导入文件……", font=(FAMILY, 11),
                                    padx=16, pady=4, command=self.on_import_file)
        self.btn_import.pack(side="left")
        # 「取消导入」（用户 2026-09-17）：把界面退回"还没导入文件"的状态。
        # 没导入时置灰——没东西可取消，留着个能点的按钮只会让人以为点错了。
        self.btn_cancel_import = tk.Button(row, text="取消导入", font=(FAMILY, 11),
                                           padx=12, pady=4, state="disabled",
                                           command=self.on_cancel_import)
        self.btn_cancel_import.pack(side="left", padx=(10, 0))
        tk.Label(row, text="当前文件：", font=(FAMILY, 11, "bold"), bg=C_PANEL,
                 fg=C_TEXT).pack(side="left", padx=(18, 2))
        self.lbl_doc = tk.Label(row, text="（尚未选择）", font=(FAMILY, 10),
                                bg=C_PANEL, fg=C_HINT)
        self.lbl_doc.pack(side="left")
        # 原文的标题编号方案（用户 2026-09-19：放在导入旁边，只有两项、强制单选，
        # 没选过默认中文）——识别不去猜，由用户明确告诉工具原文是哪一种。
        # 带 2px 的隐形边线：流程图点"选择序号体系"时靠边线变色闪一下（`_blink`）
        row_scheme = tk.Frame(blk_import, bg=C_PANEL, highlightthickness=2,
                              highlightbackground=C_PANEL)
        row_scheme.pack(anchor="center", pady=(0, 6))
        self.row_scheme = row_scheme
        row_scheme._blink_fill = True     # 闪的是整个方块（底色），不只是边线（`_blink`）
        # 样子跟主界面别的复选框一致（自造方块），但**强制单选**：真值在
        # `var_heading_scheme`，两个复选框只是它的镜像（见 `_sync_scheme_checks`）。
        tk.Label(row_scheme, text="原文标题序号为：", font=(FAMILY, 11, "bold"),
                 bg=C_PANEL, fg=C_TEXT).pack(side="left")
        self._scheme_checks: dict[str, tk.BooleanVar] = {}
        for value, text, example in (("chinese", "中文序号", "如“一、”“（一）”"),
                                     ("dotted", "阿拉伯数字序号", "如“1.1”“1.1.1”")):
            opt = tk.Frame(row_scheme, bg=C_PANEL)
            opt.pack(side="left", padx=(18, 0))
            flag = tk.BooleanVar(value=self.var_heading_scheme.get() == value)
            self._scheme_checks[value] = flag
            self._checkbutton(opt, text, variable=flag, font_size=11,
                              command=lambda v=value: self._on_pick_heading_scheme(v)
                              ).pack(anchor="w")
            tk.Label(opt, text=example, font=(FAMILY, 9), bg=C_PANEL,
                     fg=C_HINT).pack(anchor="w", padx=(26, 0))
        self.var_heading_scheme.trace_add("write", lambda *_: self._sync_scheme_checks())
        # 这里原来还留着一个 24px 的空 Frame（补被删掉的那条灰字提示的位置）；
        # 2026-09-17 起块与块之间有三条弹性空白，它就没存在的必要了。

        # 预设和配置（用户 2026-09-17：“配置文件现在改叫『预设和配置』”——跟
        # 《工具操作逻辑》窗口里的节标题（2. 预设和配置文件）对齐）
        tk.Label(blk_preset, text="预设和配置", font=(FAMILY, 13, "bold"),
                 bg=C_PANEL, fg=C_TEXT).pack(anchor="w", pady=(0, 8))
        row2 = tk.Frame(blk_preset, bg=C_PANEL)
        row2.pack(anchor="w", pady=(0, 8))
        self.btn_save_preset = tk.Button(row2, text="保存当前设置为新配置",
                                         font=(FAMILY, 10), padx=10, pady=3,
                                         command=self.on_save_preset)
        self.btn_save_preset.pack(side="left")
        self.btn_import_preset = tk.Button(row2, text="从Excel导入配置文件……",
                                           font=(FAMILY, 10), padx=10, pady=3,
                                           command=self.on_import_preset)
        self.btn_import_preset.pack(side="left", padx=(10, 0))

        self.box_presets = tk.Frame(blk_preset, bg=C_PANEL,
                                    highlightbackground=C_LINE, highlightthickness=1)
        # ！这里**不能** expand，见 PRESET_LIST_MAX_H 处的说明：这一格原来会
        # 跟着内容一起长，把下面的协议/导出按钮挤没。
        self.box_presets.pack(fill="x", pady=(0, 6))
        listbox = self.box_presets
        # 配置列表套一层定高滚动区（PRESET_LIST_MIN_H~MAX_H）
        self.preset_canvas = tk.Canvas(listbox, bg=C_PANEL, highlightthickness=0,
                                       height=PRESET_LIST_MIN_H)
        self.preset_sb = ttk.Scrollbar(listbox, orient="vertical",
                                       command=self.preset_canvas.yview)
        self.preset_canvas.configure(yscrollcommand=self.preset_sb.set)
        # 打包顺序：滚动条先占右边，再让画布铺满剩下的（反了会盖住滚动条）
        self.preset_sb.pack(side="right", fill="y")
        self.preset_canvas.pack(side="top", fill="both", expand=True)
        self.rows_box = tk.Frame(self.preset_canvas, bg=C_PANEL)
        self.preset_window = self.preset_canvas.create_window(
            (0, 0), window=self.rows_box, anchor="nw")
        # 画布宽度变了就把内层拉到同宽（横向不出滚动条）
        self.preset_canvas.bind(
            "<Configure>",
            lambda e: self.preset_canvas.itemconfigure(self.preset_window,
                                                       width=e.width))
        self._render_preset_rows()
        tk.Label(listbox, text="应用配置后仍然可以基于配置自定义修改",
                 font=(FAMILY, 9), bg=C_PANEL, fg=C_HINT).pack(side="bottom", pady=6)

        # 备份设置（用户 2026-09-17：挪到“预设和配置”之下，且整块调小——它是
        # 安全网/次要设置，不该跟“导入文件 / 预设 / 导出”那几块争视线）
        self._build_backup_settings(blk_backup)

        # 协议 + 导出
        agree_row = tk.Frame(blk_export, bg=C_PANEL)
        agree_row.pack(anchor="center", pady=(0, 8))
        self.var_agree = tk.BooleanVar(value=False)
        self._checkbutton(agree_row, "我已阅读并同意", variable=self.var_agree,
                         command=self._on_agree, font_size=13).pack(side="left")
        # 兜底：不管谁改了 var_agree，导出按钮都跟着它走（状态只有这一个来源）
        self.var_agree.trace_add("write", lambda *_a: self._on_agree())
        lnk = tk.Label(agree_row, text="《关于使用本工具的协议》", font=(FAMILY, 11, "underline"),
                      bg=C_PANEL, fg=C_LINK, cursor="hand2")
        lnk.pack(side="left")
        lnk.bind("<Button-1>", lambda e: self.on_open_agreement())
        self.btn_exp_blank = tk.Button(blk_export, text="导出空白模板到……",
                                       font=(FAMILY, 11), width=30, pady=5,
                                       command=self.on_export_blank)
        self.btn_exp_blank.pack(anchor="center", pady=(4, 8))
        self.btn_exp_doc = tk.Button(blk_export, text="导出论文到……",
                                     font=(FAMILY, 11), width=30, pady=5,
                                     command=self.on_export_doc)
        self.btn_exp_doc.pack(anchor="center")

        wire = getattr(page, "_page_wire", None)
        if wire:
            wire()

    def _build_backup_settings(self, col: tk.Frame) -> None:
        """导入文件时的自动备份——开关/路径/自动清理，见协议一.3 和跟用户的
        讨论：这是"手滑删了原文件"时的安全网，不是正常处理流程的一部分。

        用户 2026-09-17：整块**调小一号**（字号与内距都缩，见上面那四个
        `BACKUP_*_FONT` 常量），并且挪到「预设和配置」之下。
        """
        self.lbl_backup_head = tk.Label(col, text="备份设置",
                                       font=(FAMILY, BACKUP_HEAD_FONT, "bold"),
                                       bg=C_PANEL, fg=C_TEXT)
        self.lbl_backup_head.pack(anchor="w", pady=(0, 5))

        row1 = tk.Frame(col, bg=C_PANEL)
        row1.pack(anchor="w", fill="x")
        self.var_backup_enabled = tk.BooleanVar(value=self.gui_settings.backup_enabled)
        self.cb_backup_enabled = self._checkbutton(
            row1, "自动备份原文件", variable=self.var_backup_enabled,
            command=self._on_backup_setting_changed, font_size=BACKUP_TEXT_FONT)
        self.cb_backup_enabled.pack(anchor="w")
        tk.Label(col, text="可以避免在误删源文件且程序崩溃之极端情况下丢失数据",
                 font=(FAMILY, BACKUP_HINT_FONT), bg=C_PANEL, fg=C_HINT
                 ).pack(anchor="w", pady=(0, 5))

        row2 = tk.Frame(col, bg=C_PANEL)
        row2.pack(anchor="w", pady=(0, 1))
        self.btn_set_backup_dir = tk.Button(
            row2, text="设置备份路径……", font=(FAMILY, BACKUP_BTN_FONT),
            padx=8, pady=1, command=self.on_set_backup_dir)
        self.btn_set_backup_dir.pack(side="left")
        self.btn_open_backup_dir = tk.Button(
            row2, text="打开备份路径", font=(FAMILY, BACKUP_BTN_FONT),
            padx=8, pady=1, command=self.on_open_backup_dir)
        self.btn_open_backup_dir.pack(side="left", padx=(6, 0))
        self.lbl_backup_dir = tk.Label(col, text="", font=(FAMILY, BACKUP_HINT_FONT),
                                       bg=C_PANEL, fg=C_HINT, anchor="w",
                                       wraplength=520, justify="left")
        self.lbl_backup_dir.pack(anchor="w", pady=(1, 6))

        row3 = tk.Frame(col, bg=C_PANEL)
        row3.pack(anchor="w", pady=(0, 6))
        self.var_backup_prune = tk.BooleanVar(value=self.gui_settings.backup_auto_prune)
        self.cb_backup_prune = self._checkbutton(
            row3, "自动删除过早备份", variable=self.var_backup_prune,
            command=self._on_backup_setting_changed, font_size=BACKUP_TEXT_FONT)
        self.cb_backup_prune.pack(side="left")
        self.var_backup_keep = tk.StringVar(value=str(self.gui_settings.backup_keep_count))
        self.ent_backup_keep = tk.Entry(row3, textvariable=self.var_backup_keep, width=4,
                                        font=(FAMILY, BACKUP_TEXT_FONT), justify="center")
        self.ent_backup_keep.pack(side="left", padx=(6, 2))
        self.ent_backup_keep.bind("<FocusOut>",
                                  lambda e: self._on_backup_keep_count_changed())
        self.ent_backup_keep.bind("<Return>",
                                  lambda e: self._on_backup_keep_count_changed())
        self.var_backup_keep_unit = tk.StringVar(value=self.gui_settings.backup_keep_unit)
        self.cb_backup_unit = ttk.Combobox(row3, textvariable=self.var_backup_keep_unit,
                                           width=3, state="readonly",
                                           font=(FAMILY, BACKUP_TEXT_FONT),
                                           values=["条", "天", "周", "月"])
        self.cb_backup_unit.pack(side="left", padx=(2, 0))
        self.cb_backup_unit.bind("<<ComboboxSelected>>",
                                 lambda e: self._on_backup_keep_unit_changed())
        self._lock_wheel(self.cb_backup_unit)
        self._refresh_prune_controls()

        # 原来这里有一条灰字提示“《条》按份数保留最近的备份；《天/周/月》…”，
        # 用户 2026-09-16 点名删掉。

        self._refresh_backup_dir_label()

    def _effective_backup_dir(self) -> pathlib.Path:
        """当前实际生效的备份目录：用户设置过就用那个，没设置就用默认位置。"""
        override = self.gui_settings.backup_dir
        return pathlib.Path(override) if override else import_backup_dir()

    def _refresh_backup_dir_label(self) -> None:
        self.lbl_backup_dir.config(text=f"当前备份路径：{self._effective_backup_dir()}")

    def _refresh_prune_controls(self) -> None:
        """「自动删除过早备份」没勾时，旁边的份数/单位跟着置灰。

        这两项只对这个开关负责：摆在可编辑状态会让人以为"不勾也会清理"，
        而且改完还会写进 settings.json（属于"看着能设、其实不生效"的控件）。
        """
        on = bool(self.var_backup_prune.get())
        self.ent_backup_keep.config(state="normal" if on else "disabled")
        self.cb_backup_unit.config(state="readonly" if on else "disabled")

    def _save_gui_settings(self) -> None:
        try:
            save_settings(self.gui_settings)
        except OSError as exc:
            self.status(f"备份设置保存失败（本次会话内仍生效）：{exc}")

    def _on_backup_setting_changed(self) -> None:
        self.gui_settings.backup_enabled = bool(self.var_backup_enabled.get())
        self.gui_settings.backup_auto_prune = bool(self.var_backup_prune.get())
        self._refresh_prune_controls()
        self._save_gui_settings()

    def _on_backup_keep_count_changed(self) -> None:
        raw = self.var_backup_keep.get().strip()
        try:
            n = int(raw)
            if n <= 0:
                raise ValueError
        except ValueError:
            self.status(f"自动删除过早备份的数量须为正整数；{raw} 无法识别，"
                        f"已恢复为 {self.gui_settings.backup_keep_count}。")
            self.var_backup_keep.set(str(self.gui_settings.backup_keep_count))
            return
        self.gui_settings.backup_keep_count = n
        self._save_gui_settings()

    def _on_backup_keep_unit_changed(self) -> None:
        self.gui_settings.backup_keep_unit = self.var_backup_keep_unit.get()
        self._save_gui_settings()

    def on_set_backup_dir(self) -> None:
        chosen = filedialog.askdirectory(
            title="设置备份路径", initialdir=str(self._effective_backup_dir()))
        if not chosen:
            return
        self.gui_settings.backup_dir = chosen
        self._save_gui_settings()
        self._refresh_backup_dir_label()
        self.status(f"备份路径已设置为：{chosen}")

    def on_open_backup_dir(self) -> None:
        path = self._effective_backup_dir()
        try:
            path.mkdir(parents=True, exist_ok=True)
            os.startfile(str(path))  # noqa: S606 —— Windows 专用工具，就是要用这个
        except OSError as exc:
            self.status(f"打开备份路径失败：{exc}")

    # ------------------------------------------------------------ 页面页签

    def _fmt_cm(self, length) -> str:
        """把 Emu 长度显示成厘米字符串（最多两位小数，去尾零）。"""
        try:
            v = length.cm
        except Exception:
            try:
                v = length / 360000
            except Exception:
                return ""
        return f"{v:.2f}".rstrip("0").rstrip(".")

    def _build_tab_page(self, tab: tk.Frame) -> None:
        page = self._make_scroll_page(tab)   # 窗口小的时候整页可滚
        col = tk.Frame(page, bg=C_PANEL)
        col.pack(fill="both", expand=True, padx=46, pady=(20, 14))

        top = tk.Frame(col, bg=C_PANEL)
        top.pack(fill="x")
        tk.Label(top, text="页面设置", font=(FAMILY, 13, "bold"), bg=C_PANEL,
                 fg=C_TEXT).pack(side="left")
        self.lbl_page_src = tk.Label(top, text="（未应用配置）", font=(FAMILY, 9),
                                     bg=C_PANEL, fg=C_HINT)
        self.lbl_page_src.pack(side="left", padx=(10, 0))

        self.pv: dict[str, tk.Variable] = {}

        def entry(parent, name: str, width: int = 7) -> tk.Entry:
            v = tk.StringVar()
            self.pv[name] = v
            return tk.Entry(parent, textvariable=v, width=width, font=(FAMILY, 10),
                            justify="center")

        def lab(parent, text: str, hint: bool = False) -> tk.Label:
            return tk.Label(parent, text=text, font=(FAMILY, 10), bg=C_PANEL,
                            fg=C_HINT if hint else C_TEXT)

        def group(title: str) -> tk.Frame:
            tk.Label(col, text=title, font=(FAMILY, 11, "bold"), bg=C_PANEL,
                     fg=C_TEXT).pack(anchor="w", pady=(26, 10))
            f = tk.Frame(col, bg=C_PANEL)
            f.pack(anchor="w")
            return f

        # 纸张与方向
        g1 = group("纸张与方向")
        lab(g1, "纸张：").pack(side="left")
        self.cb_paper = ttk.Combobox(g1, width=7, state="readonly",
                                     values=["A4", "A3", "B5", "自定义"],
                                     font=(FAMILY, 10))
        self.cb_paper.pack(side="left")
        self._lock_wheel(self.cb_paper)
        self.cb_paper.bind("<<ComboboxSelected>>", lambda e: self._toggle_custom_size())
        lab(g1, "　方向：").pack(side="left")
        self.cb_orient = ttk.Combobox(g1, width=5, state="readonly",
                                      values=["纵向", "横向"], font=(FAMILY, 10))
        self.cb_orient.pack(side="left")
        self._lock_wheel(self.cb_orient)
        lab(g1, "　装订线：").pack(side="left")
        entry(g1, "gutter").pack(side="left")
        lab(g1, " 厘米", hint=True).pack(side="left")

        g1b = tk.Frame(col, bg=C_PANEL)
        g1b.pack(anchor="w", pady=(10, 0))
        lab(g1b, "自定义尺寸：宽 ").pack(side="left")
        self.e_pw = entry(g1b, "pw")
        self.e_pw.pack(side="left")
        lab(g1b, " × 高 ", hint=True).pack(side="left")
        self.e_ph = entry(g1b, "ph")
        self.e_ph.pack(side="left")
        lab(g1b, " 厘米", hint=True).pack(side="left")

        # 页边距与距边
        group("页边距与距边（厘米）")
        grid = tk.Frame(col, bg=C_PANEL)
        grid.pack(anchor="w")
        specs = (("上边距", "mt"), ("下边距", "mb"), ("左边距", "ml"),
                 ("右边距", "mr"), ("页眉距边", "hd"), ("页脚距边", "fd"))
        for i, (label, key) in enumerate(specs):
            cell = tk.Frame(grid, bg=C_PANEL)
            cell.grid(row=i // 3, column=i % 3, padx=(0, 34), pady=7, sticky="w")
            lab(cell, label + "：").pack(side="left")
            entry(cell, key, width=6).pack(side="left")

        # 页眉（按节）
        group("页眉")
        self.hdr_box = tk.Frame(col, bg=C_PANEL)
        self.hdr_box.pack(anchor="w", fill="x")
        self.hdr_vars: list[dict] = []
        # 不套用任何预设时的默认值：页眉装饰线选否（用户 2026-09-14 定）。
        # 应用预设时会被预设里的值覆写（见 _load_page_from）。
        self.v_headerline = tk.BooleanVar(value=False)
        self._checkbutton(col, "为除封面外的页眉应用装饰线", variable=self.v_headerline
                         ).pack(anchor="w", pady=(10, 0))
        # 装饰线左右两端各向正文栏外延伸多少厘米（0＝与正文栏等宽）。没勾装饰线时
        # 这两格置灰——线都没有，谈不上延伸。
        g_ext = tk.Frame(col, bg=C_PANEL)
        g_ext.pack(anchor="w", pady=(8, 0))
        lab(g_ext, "装饰线向外延伸：左 ").pack(side="left")
        self.e_hl = entry(g_ext, "hl", width=5)
        self.e_hl.pack(side="left")
        lab(g_ext, " 厘米　右 ").pack(side="left")
        self.e_hr = entry(g_ext, "hr", width=5)
        self.e_hr.pack(side="left")
        lab(g_ext, " 厘米　（0＝与正文栏等宽）", hint=True).pack(side="left")

        def _sync_ext_state(*_a) -> None:
            state = "normal" if self.v_headerline.get() else "disabled"
            self.e_hl.config(state=state)
            self.e_hr.config(state=state)

        self.v_headerline.trace_add("write", _sync_ext_state)
        self.pv["hl"].set("0")
        self.pv["hr"].set("0")
        _sync_ext_state()

        # 页脚（页码）
        group("页脚")
        self.ftr_box = tk.Frame(col, bg=C_PANEL)
        self.ftr_box.pack(anchor="w", fill="x")
        self.ftr_vars: list[dict] = []

        # 原来这里有一条灰字提示“这里改的值只记在本次会话；去文件页签点…”，
        # 用户 2026-09-16 点名删掉（原来的段前 28px 间隔搬到一个空 Frames 上）。
        tk.Frame(col, bg=C_PANEL, height=28).pack()

        self._toggle_custom_size()
        self._rebuild_hdr_rows([])
        self._rebuild_ftr_rows([])
        wire = getattr(page, "_page_wire", None)
        if wire:
            wire()

    def _toggle_custom_size(self) -> None:
        custom = self.cb_paper.get() == "自定义"
        state = "normal" if custom else "disabled"
        self.e_pw.config(state=state)
        self.e_ph.config(state=state)

    def _rebuild_hdr_rows(self, rules: list) -> None:
        for w in self.hdr_box.winfo_children():
            w.destroy()
        self.hdr_vars = []
        if not rules:
            tk.Label(self.hdr_box, text="（还没读到节信息——先应用一个配置）",
                     font=(FAMILY, 9), bg=C_PANEL, fg=C_HINT).pack(anchor="w")
            return
        for rule in rules:
            rv: dict = {"section": getattr(rule, "section_name", "?")}
            row = tk.Frame(self.hdr_box, bg=C_PANEL)
            row.pack(anchor="w", pady=5)
            tk.Label(row, text=getattr(rule, "section_name", "?"), font=(FAMILY, 10),
                     bg=C_PANEL, fg=C_TEXT, width=20, anchor="w").pack(side="left")
            tk.Label(row, text="页眉文字：", font=(FAMILY, 10), bg=C_PANEL,
                     fg=C_TEXT).pack(side="left")
            rv["header"] = tk.StringVar(value=getattr(rule, "header_text", ""))
            tk.Entry(row, textvariable=rv["header"], width=32,
                     font=(FAMILY, 10)).pack(side="left")
            self.hdr_vars.append(rv)

    def _rebuild_ftr_rows(self, rules: list) -> None:
        for w in self.ftr_box.winfo_children():
            w.destroy()
        self.ftr_vars = []
        if not rules:
            tk.Label(self.ftr_box, text="（还没读到节信息——先应用一个配置）",
                     font=(FAMILY, 9), bg=C_PANEL, fg=C_HINT).pack(anchor="w")
            return
        for rule in rules:
            rv: dict = {"section": getattr(rule, "section_name", "?")}
            row = tk.Frame(self.ftr_box, bg=C_PANEL)
            row.pack(anchor="w", pady=5)
            tk.Label(row, text=getattr(rule, "section_name", "?"), font=(FAMILY, 10),
                     bg=C_PANEL, fg=C_TEXT, width=20, anchor="w").pack(side="left")
            rv["num"] = tk.BooleanVar(value=bool(getattr(rule, "page_number_enabled",
                                                          False)))
            self._checkbutton(row, "页码", variable=rv["num"],
                              command=lambda v=rv: self._refresh_ftr_row(v)
                              ).pack(side="left")
            tk.Label(row, text="　格式：", font=(FAMILY, 10), bg=C_PANEL,
                     fg=C_TEXT).pack(side="left")
            fmt = getattr(rule, "page_number_format", None)
            rv["fmt"] = tk.StringVar(value=PAGE_FMT_LABEL.get(fmt, fmt or "阿拉伯数字"))
            cb = ttk.Combobox(row, textvariable=rv["fmt"], width=10, state="readonly",
                              values=list(PAGE_FMT_LABEL.values()), font=(FAMILY, 9))
            self._lock_wheel(cb)
            cb.pack(side="left")
            rv["fmt_cb"] = cb
            tk.Label(row, text="　前缀：", font=(FAMILY, 10), bg=C_PANEL,
                     fg=C_TEXT).pack(side="left")
            rv["pre"] = tk.StringVar(value=getattr(rule, "page_number_prefix", ""))
            rv["pre_ent"] = tk.Entry(row, textvariable=rv["pre"], width=4,
                                     font=(FAMILY, 10), justify="center")
            rv["pre_ent"].pack(side="left")
            tk.Label(row, text="　后缀：", font=(FAMILY, 10), bg=C_PANEL,
                     fg=C_TEXT).pack(side="left")
            rv["suf"] = tk.StringVar(value=getattr(rule, "page_number_suffix", ""))
            rv["suf_ent"] = tk.Entry(row, textvariable=rv["suf"], width=4,
                                     font=(FAMILY, 10), justify="center")
            rv["suf_ent"].pack(side="left")
            self._refresh_ftr_row(rv)
            self.ftr_vars.append(rv)

    def _refresh_ftr_row(self, rv: dict) -> None:
        """没勾"页码"的节，格式/前缀/后缀置灰。

        这三项只在有页码时有意义：`page_number_enabled=False` 时收集配置会把
        `page_number_format` 直接写成 None（见 `_collect_current_preset_config`），
        改了半天保存时被静默丢掉。
        """
        on = bool(rv["num"].get())
        rv["fmt_cb"].config(state="readonly" if on else "disabled")
        rv["pre_ent"].config(state="normal" if on else "disabled")
        rv["suf_ent"].config(state="normal" if on else "disabled")

    def _load_page_from(self, cfg, label: str) -> None:
        p = cfg.page
        # 这几项预设里有、但界面上没有控件能编辑（如保留字符格式）——
        # 先记住，「保存当前设置为新配置」的时候原样带回去，免得应用一次配置
        # 就把它们丢了。（2026-09-14 起预设里不再有"打印方式"，所以那份
        # 搬运也一起删了。）
        self._preset_extra = {
            "keep_run_formatting": cfg.keep_run_formatting,
        }
        self.lbl_page_src.config(text=f"当前配置：{label}")
        for cb, val in ((self.cb_paper, p.paper_size), (self.cb_orient, p.orientation)):
            vals = list(cb.cget("values"))
            if val and val not in vals:
                cb.config(values=vals + [val])
            cb.set(val or "")
        for key, val in (("pw", p.page_width), ("ph", p.page_height),
                         ("gutter", p.gutter), ("mt", p.margin_top),
                         ("mb", p.margin_bottom), ("ml", p.margin_left),
                         ("mr", p.margin_right), ("hd", p.header_distance),
                         ("fd", p.footer_distance)):
            self.pv[key].set(self._fmt_cm(val))
        self.pv["hl"].set(f"{cfg.header_line_extend_left_cm:g}")
        self.pv["hr"].set(f"{cfg.header_line_extend_right_cm:g}")
        self.v_headerline.set(bool(cfg.header_bottom_line))
        rules = list(cfg.header_footer or [])
        self._rebuild_hdr_rows(rules)
        self._rebuild_ftr_rows(rules)
        self._toggle_custom_size()

    def _load_preset_defaults(self) -> None:
        """启动时拿配置列表里第一条（内置默认排前面）把「页面」「格式」「结构」
        「封面」填上（读得到就填；这一步只填表单，不算「已应用」——导出按钮
        仍然要求用户自己在列表里点一次「应用」）。

        填完再盖一层「不套用任何预设时的默认值」：先填内置毕业论文预设，
        相当于用户说的"其他未尽事宜参照毕业论文格式"；再按他点名的几条
        覆盖，见 `_apply_no_preset_baseline`。
        """
        preset = next((r["path"] for r in self._preset_rows
                       if not r.get("pending") and r.get("path") is not None), None)
        if preset is None:
            return
        try:
            from essay_tool.config.preset_io import load_preset
            cfg = load_preset(preset)
        except Exception as exc:
            self.lbl_page_src.config(text=f"（预设读取失败：{exc}）")
            self.lbl_fmt_src.config(text=f"（预设读取失败：{exc}）")
            return
        self._load_page_from(cfg, preset.stem)
        self._load_formats_from(cfg, preset.stem)
        self._load_cover_from(cfg, preset.stem)
        # 「结构」不在这里填——那是文档自己的事，见 _rebuild_struct_adjust。
        self._apply_no_preset_baseline()

    def _apply_no_preset_baseline(self) -> None:
        """「不套用任何预设」时的默认值（用户 2026-09-14 定稿）。

        用户原话：「不套用任何预设的默认值按照以下办法设置：1. 页面设置，
        页眉装饰线选否；2. 页眉文字为空；3. 结构部件，题目开、副标题开、
        摘要开、目录开、引言开、结论开；4. 图表自动编号开、保留字符格式开；
        5. 不使用封面。除了部件结构启用以外的其他未尽事宜，参照毕业论文格式。」

        落地方式：启动时先把内置毕业论文预设填进表单（＝"参照毕业论文
        格式"），再用这个方法覆盖下面的几条。第 3 条的真身在
        `STRUCTURE_PARTS` 的"默认启用"列和 `adjust_preview.ENABLED` 里
        （那张表格导入文档后会按"预设 ∪ 识别到的部件"重算）；第 4 条里
        "保留字符格式"见 `cleanup.DEFAULT_KEEP_RUN_FORMATS`（预设里没配
        ＝保留 加粗/斜体/下划线/上下标）。

        ⚠️ "图表自动编号"那一项 **2026-09-16 已经取消**（用户定：我们只管
        样式、编号与我们无关）——预设里不再有这个键，这里的第 4 条只剩
        保留字符格式。
        """
        self.v_headerline.set(False)          # ① 页眉装饰线：否
        self.pv["hl"].set("0")                #    延伸长度归零
        self.pv["hr"].set("0")
        for hv in self.hdr_vars:              # ② 页眉文字：空
            hv["header"].set("")
        self.cv_cover["on"].set(False)        # ⑤ 不使用封面
        self._cover_power()                   #    封面整块随之变灰

    # ------------------------------------------------------------ 格式页签

    @property
    def fmt_rows(self) -> list[dict]:
        """格式表每行的变量字典（导出/保存预设读它；表格本体见 `fmt_grid.FmtGrid`）。"""
        return self.fmt_grid.rows

    def _build_tab_fmt(self, tab: tk.Frame) -> None:
        page = self._make_scroll_page(tab)   # 窗口很小时整页可滚
        col = tk.Frame(page, bg=C_PANEL)
        col.pack(fill="both", expand=True, padx=20, pady=(20, 12))

        top = tk.Frame(col, bg=C_PANEL)
        top.pack(fill="x")
        tk.Label(top, text="格式和引注", font=(FAMILY, 13, "bold"),
                 bg=C_PANEL, fg=C_TEXT).pack(side="left")
        self.lbl_fmt_src = tk.Label(top, text="（未应用配置）", font=(FAMILY, 9),
                                    bg=C_PANEL, fg=C_HINT)
        self.lbl_fmt_src.pack(side="left", padx=(10, 0))
        # 原来这里还有一条灰字提示（“行为样式、列为设置项；第一行/第一列固定…”，
        # 里面带着两个《》角括号），用户 2026-09-16 点名删掉。

        # 自绘表格（`fmt_grid.py`）：四象限、可折叠分组，点哪一格才临时放一个真编辑框。
        # 旧版一格一个真控件（约 700 个），展开/拖分隔条都卡（用户 2026-09-20）。
        from essay_tool.engine.styles import DIRECT_FORMAT_IDS
        self._fmt_open: dict[str, bool] = {}   # 块的展开状态（默认仅正文部分展开）
        self.fmt_grid = FmtGrid(
            col, theme=dict(panel=C_PANEL, text=C_TEXT, hint=C_HINT, line=C_LINE),
            direct_ids=DIRECT_FORMAT_IDS, open_state=self._fmt_open)
        self.fmt_grid.pack(fill="both", expand=True)
        self.fmt_wrap = self.fmt_grid
        self.fmt_grid.set_groups([])

        # ＊ 的说明放在表格下面（用户 2026-09-13 要求；文案 2026-09-16 换成
        # 用户自己给的这一句）。
        # 注：文本比栏目宽，必须 wraplength + justify=left，否则居中裁两侧
        legend = "＊ 这些条目将不会在Word中创建样式，而是直接在段落上调整格式。"
        tk.Label(col, text=legend, font=(FAMILY, 9), bg=C_PANEL, fg=C_HINT,
                 wraplength=740, justify="left").pack(anchor="w", pady=(6, 0))

        # 引注块：从大表格里摘出来的设置（2026-09-13 用户要求）
        foot = tk.Frame(col, bg=C_PANEL)
        foot.pack(fill="x", pady=(14, 0))
        tk.Label(foot, text="引注", font=(FAMILY, 11, "bold"), bg=C_PANEL,
                 fg=C_TEXT).pack(anchor="w")
        # 原来这里有一句灰字说明（“只放表格没涵盖到的：……”），用户 2026-09-16
        # 点名删掉。
        self._build_fn_box(foot)
        # 标签写法块：预设的一部分（随预设保存/导出），所以放在格式页而不是结构页
        tk.Label(foot, text="标签写法", font=(FAMILY, 11, "bold"), bg=C_PANEL,
                 fg=C_TEXT).pack(anchor="w", pady=(16, 0))
        self._build_label_box(foot)
        # 清理规则块：同样是预设的一部分（用户 2026-09-20）
        tk.Label(foot, text="清理规则", font=(FAMILY, 11, "bold"), bg=C_PANEL,
                 fg=C_TEXT).pack(anchor="w", pady=(16, 0))
        self._build_cleanup_box(foot)
        wire = getattr(page, "_page_wire", None)
        if wire:
            wire()

    @staticmethod
    def _sb_state(sb, first, last) -> None:
        """内容装得下就把滚动条置灰（用户 2026-09-13：没超过一页就不该能滚）。"""
        fit = float(first) <= 0.0 and float(last) >= 1.0
        sb.state(["disabled"] if fit else ["!disabled"])

    @staticmethod
    def _scroll_canvas(cv: tk.Canvas, units: int, vertical: bool) -> None:
        """按格滚动画布，但**夹在内容范围内**。

        Tk 画布向上（左）滚没有下限，会滚进"负区域"：内容被整体推下去、
        而 yview 分数不变——同步回调因此不触发，另一侧就不动（用户
        2026-09-13 抓到：内容装得下也能滚、右边动左边不动）。这里统一夹好。
        """
        bb = cv.bbox("all")
        if not bb or not units:
            return
        if vertical:
            total, view, top = bb[3], cv.winfo_height(), cv.canvasy(0)
        else:
            total, view, top = bb[2], cv.winfo_width(), cv.canvasx(0)
        step = max(20, int(view * 0.1))   # 一格≈视口的 1/10（和 Tk 的 unit 一致）
        new_top = max(0.0, min(top + units * step, max(0.0, total - view)))
        frac = new_top / total if total else 0.0
        if vertical:
            cv.yview_moveto(frac)
        else:
            cv.xview_moveto(frac)

    def _make_scroll_page(self, tab: tk.Frame) -> tk.Frame:
        """给页签套一层"滚动页"：窗口拉得很小时，被挡住的选项能滚出来
        （用户 2026-09-13：以前只能把窗口拉大）。

        返回内层 Frame——页签原来的内容 pack 到它里面；内容建好后调一次
        `inner._page_wire()` 把滚轮接上（下拉框仍锁滚轮）。窗口够大时两条
        滚动条自动置灰不碍事。
        """
        cv = tk.Canvas(tab, bg=C_PANEL, highlightthickness=0)
        vsb = ttk.Scrollbar(tab, orient="vertical", command=cv.yview)
        hsb = ttk.Scrollbar(tab, orient="horizontal", command=cv.xview)

        def sync_v(f, l) -> None:
            vsb.set(f, l)          # 滑块本身的位置/长度——漏了这行滑块就不会动
            self._sb_state(vsb, f, l)

        def sync_h(f, l) -> None:
            hsb.set(f, l)
            self._sb_state(hsb, f, l)

        cv.configure(yscrollcommand=sync_v, xscrollcommand=sync_h)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        cv.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(cv, bg=C_PANEL)
        win = cv.create_window((0, 0), window=inner, anchor="nw")

        refit_job = {"id": None}

        def refit_now() -> None:
            refit_job["id"] = None
            # 视口大就撑满（观感不变）；视口比内容小就保持原尺寸 → 出滚动条
            # （用户 2026-09-13：窗口很小时选项被挡住、又滚不动）
            cv.itemconfigure(win,
                             width=max(cv.winfo_width(),
                                       inner.winfo_reqwidth()),
                             height=max(cv.winfo_height(),
                                        inner.winfo_reqheight()))
            cv.configure(scrollregion=cv.bbox("all"))
            sync_v(*cv.yview())
            sync_h(*cv.xview())

        def refit(_e=None) -> None:
            """视口尺寸变了：先记下，停手 90ms 再真正重排一次。

            原来这里是每帧都算的（`cv` 和 `inner` 各绑一次 Configure，拖一下
            窗口就跑两遍，每遍都要量整页内容的请求尺寸 + 改 window item 的
            尺寸 ⇒ 整页重排）。拖动窗口时几十帧连发，主线程全花在这上面，
            Windows 来不及重绘窗口，屏幕上一块块露出旧内容（用户 2026-09-16
            反馈的"乱码色块"）。

            而"视口撑满 / 出滚动条"这件事**滞后 90ms 完全看不出来**，所以
            防抖是白捐的。两个绑定共用同一个 job，顺带把"每帧跑两遍"也省了。
            """
            if refit_job["id"] is not None:
                try:
                    cv.after_cancel(refit_job["id"])
                except Exception:                      # noqa: BLE001
                    pass
            refit_job["id"] = cv.after(90, refit_now)

        cv.bind("<Configure>", refit)
        inner.bind("<Configure>", refit)

        def wheel(event, vertical=True):
            notch = wheel_notch(event)
            self._scroll_canvas(cv, -2 * notch, vertical)
            return "break"

        def wire(widget=None) -> None:
            # 下拉框锁滚轮；画布交给它们自己的滚轮逻辑（比如格式表）
            widget = inner if widget is None else widget
            if not isinstance(widget, (ttk.Combobox, tk.Canvas)):
                widget.bind("<MouseWheel>", lambda e: wheel(e, True))
                widget.bind("<Shift-MouseWheel>", lambda e: wheel(e, False))
            for ch in widget.winfo_children():
                wire(ch)

        inner._page_wire = wire
        self.after_idle(refit_now)
        return inner

    def _lock_wheel(self, widget) -> None:
        """拦掉 Combobox 的滚轮改值（返回 "break"，不让类绑定处理）。"""
        widget.bind("<MouseWheel>", lambda e: "break")
        return widget

    @staticmethod
    def _group_fmt_specs(specs: list) -> list[tuple[str, list]]:
        """按四大块分组（用户 2026-09-13 定的顺序：整表按论文出现顺序）。

        组内按 FMT_GROUPS 里的名单排；名单外的样式兜底进「其它」组。空组不显示。
        """
        by_name = {s.name: s for s in specs}
        groups: list[tuple[str, list]] = []
        used: set[str] = set()
        for title, names in FMT_GROUPS:
            members = [by_name[n] for n in names if n in by_name]
            used.update(n for n in names if n in by_name)
            if members:
                groups.append((title, members))
        rest = [s for s in specs if s.name not in used]
        if rest:
            groups.append(("其它（不在名单里）", rest))
        return groups

    def _load_formats_from(self, cfg, label: str) -> None:
        self.lbl_fmt_src.config(text=f"当前配置：{label}")
        self.fmt_grid.set_groups(self._group_fmt_specs(list(cfg.styles.values())))
        for rv in self.fmt_grid.rows:                      # 样式表"段前分页"一改，没动过的标签复选框跟着变
            rv["pb"].trace_add("write", lambda *_a: self._sync_label_pb_defaults())
        self._load_citation_from(cfg)
        self._load_labels_from(cfg)
        self._load_cleanup_from(cfg)

    def _build_fn_box(self, parent: tk.Frame) -> None:
        """引注块：脚注的编号设置 + 尾注（文末参考文献列表）的条目编号。"""
        box = tk.Frame(parent, bg=C_PANEL, highlightbackground=C_LINE,
                       highlightthickness=1)
        box.pack(fill="x")
        self.cv: dict[str, tk.Variable] = {}    # 编号设置的值
        self.cv_cb: dict[str, ttk.Combobox] = {}

        def newrow(pady=(4, 0)) -> tk.Frame:
            f = tk.Frame(box, bg=C_PANEL)
            f.pack(fill="x", padx=12, pady=pady)
            return f

        def lab(par, text, padx=(10, 3), hint=False, bold=False):
            tk.Label(par, text=text,
                     font=(FAMILY, 10, "bold" if bold else "normal"),
                     bg=C_PANEL, fg=C_HINT if hint else C_TEXT
                     ).pack(side="left", padx=padx)

        def spacer(par, w=62):
            tk.Frame(par, bg=C_PANEL, width=w, height=1).pack(side="left")

        def combo(par, key, values, width, store, registry):
            store[key] = tk.StringVar()
            cb = ttk.Combobox(par, textvariable=store[key], values=values,
                              width=width, font=(FAMILY, 10), state="readonly")
            cb.pack(side="left")
            self._lock_wheel(cb)
            registry[key] = cb
            return cb

        def check(par, key, text, store):
            store[key] = tk.BooleanVar()
            self._checkbutton(par, text, variable=store[key]
                             ).pack(side="left", padx=(12, 0))

        # —— 脚注（跨页注）
        r1 = newrow(pady=(10, 0))
        lab(r1, "脚注", padx=(0, 3), bold=True)
        lab(r1, "序号样式")
        combo(r1, "fn_style", ["①②③", "1.2.3."], 10, self.cv,
              self.cv_cb)
        lab(r1, "序号字体")
        combo(r1, "fn_font", ["正文(非上标)", "上标"], 11, self.cv, self.cv_cb)

        r2 = newrow()
        spacer(r2)
        lab(r2, "序号后空白", padx=(0, 3))
        combo(r2, "fn_space", ["半角空格", "全角空格", "不空格"], 7, self.cv,
              self.cv_cb)
        check(r2, "fn_restart", "每页重新编号", self.cv)

        tk.Frame(box, bg=CELL_LINE, height=1).pack(fill="x", padx=12, pady=(8, 0))

        # —— 尾注（用户口径：文末参考文献列表算尾注）
        r3 = newrow(pady=(8, 10))
        lab(r3, "尾注", padx=(0, 3), bold=True)
        lab(r3, "条目编号")
        combo(r3, "ref_style", ["[1][2][3]", "1.2.3.", "(1)(2)(3)"], 10,
              self.cv, self.cv_cb)
        lab(r3, "（文末参考文献列表里每条前面的编号）", padx=(8, 0), hint=True)

    def _build_label_box(self, parent: tk.Frame) -> None:
        """标签写法块：每个标签类部件一格（摘要、关键词、目录……），下面另有一行
        "引言、结论两字之间空几格"。

        2026-09-19 起标签写法允许自定义（十星的"摘要""目录"没有【】）。格子里写
        的就是成品里的写法；**留空＝保持文档里的原文**。题目/副标题/英文题目/
        英文副标题是"内容占位语"（工具行为），不许自定义，所以不在这里。

        引言/结论**不给文字格、只给一个数字**（用户 2026-09-19）：规范原话是
        "引言、结论两字之间空两格"，而引言家族认出来后本来就保留作者写的词
        （"绪论"仍是"绪论"），格子里填什么词都会被忽略，容易误会。数字在预设里
        仍存成"引  言"这样的模板（`schema.heading_label_with_gap`）。
        """
        box = tk.Frame(parent, bg=C_PANEL, highlightbackground=C_LINE,
                       highlightthickness=1)
        box.pack(fill="x", pady=(4, 0))
        grid = tk.Frame(box, bg=C_PANEL)
        grid.pack(fill="x", padx=12, pady=(10, 4))
        self.label_vars: dict[str, tk.StringVar] = {}
        # "标签前另起一页"（2026-09-20）：每个标签部件一个复选框，覆盖样式表里的"段前分页"。
        # 没动过的（`_label_pb_touched` 为假）跟随样式；动过且与样式不同才写进预设。
        self.label_pb_vars: dict[str, tk.BooleanVar] = {}
        self._label_pb_touched: dict[str, bool] = {}
        defaults = {n: txt for n, _on, txt, _note in STRUCTURE_PARTS}
        for i, name in enumerate(LABEL_PART_NAMES):
            r, c = divmod(i, 2)
            tk.Label(grid, text=name, font=(FAMILY, 10), bg=C_PANEL, fg=C_TEXT,
                     width=13, anchor="w").grid(row=r, column=c * 3, sticky="w",
                                                pady=2, padx=(0 if c == 0 else 24, 0))
            var = tk.StringVar(value=defaults.get(name, ""))
            self.label_vars[name] = var
            var.trace_add("write", lambda *_: self._schedule_label_refresh())
            tk.Entry(grid, textvariable=var, font=(FAMILY, 10), width=20
                     ).grid(row=r, column=c * 3 + 1, sticky="w", pady=2)
            self._label_pb_checkbox(grid, name).grid(row=r, column=c * 3 + 2, sticky="w",
                                                     pady=2, padx=(8, 0))
        tk.Label(box, text="用于设置上述部件标题处应当出现的文字。例如，对关键词设置为“【关键词】”"
                           "将会把所有识别到的关键词标题（如“[关键词]”“关键词”）修正为“【关键词】”。"
                           "但是，这一设置不会影响识别逻辑。",
                 font=(FAMILY, 9), bg=C_PANEL, fg=C_HINT, wraplength=700,
                 justify="left").pack(anchor="w", padx=12, pady=(0, 8))

        # —— 引言、结论：两字之间空几格（只收 0~9 的一位数字；空＝保持文档原文）
        gaps = tk.Frame(box, bg=C_PANEL)
        gaps.pack(fill="x", padx=12, pady=(0, 2))
        self.heading_gap_vars: dict[str, tk.StringVar] = {}
        self._heading_raw: dict[str, tuple[str | None, int | None]] = {}
        self._heading_touched: dict[str, bool] = {}
        one_digit = (self.register(lambda p: p == "" or (p.isdigit() and len(p) == 1)), "%P")
        for i, name in enumerate(HEADING_LABEL_PARTS):
            var = tk.StringVar(value="")
            self.heading_gap_vars[name] = var
            self._heading_touched[name] = False
            var.trace_add("write", lambda *_a, n=name: self._on_heading_gap_edit(n))
            tk.Label(gaps, text=f"{name}两字之间空", font=(FAMILY, 10), bg=C_PANEL,
                     fg=C_TEXT).pack(side="left", padx=(0 if i == 0 else 28, 4))
            tk.Entry(gaps, textvariable=var, font=(FAMILY, 10), width=3, justify="center",
                     validate="key", validatecommand=one_digit).pack(side="left")
            tk.Label(gaps, text="格", font=(FAMILY, 10), bg=C_PANEL,
                     fg=C_TEXT).pack(side="left", padx=(4, 0))
            self._label_pb_checkbox(gaps, name).pack(side="left", padx=(8, 0))
        tk.Frame(box, bg=C_PANEL, height=10).pack()

    # ------------------------------------------------------------ 标签前另起一页
    def _label_pb_checkbox(self, parent, name: str):
        var = tk.BooleanVar(value=False)
        self.label_pb_vars[name] = var
        self._label_pb_touched[name] = False

        def clicked(n=name) -> None:
            self._label_pb_touched[n] = True          # 用户亲手改过 → 不再跟随样式
        return self._checkbutton(parent, "另起一页", variable=var, command=clicked, font_size=10)

    def _label_style_default_pb(self, part: str) -> bool:
        """这个标签套的那个样式，在格式表里"段前分页"现在勾没勾。"""
        from essay_tool.engine.apply_structure import label_style_id
        sid = label_style_id(part)
        for rv in getattr(self.fmt_grid, "rows", []):
            if rv.get("id") == sid:
                return bool(rv["pb"].get())
        return False

    def _sync_label_pb_defaults(self) -> None:
        """没被用户动过的标签，"另起一页"复选框跟随样式表（样式表那一格改了这里也跟着变）。"""
        for name, var in getattr(self, "label_pb_vars", {}).items():
            if not self._label_pb_touched.get(name):
                var.set(self._label_style_default_pb(name))

    def _collect_label_page_break(self) -> dict[str, bool]:
        """写进预设的"标签前另起一页"覆盖：只记动过、且跟样式不一样的（其余跟随样式）。"""
        out: dict[str, bool] = {}
        for name, var in getattr(self, "label_pb_vars", {}).items():
            if self._label_pb_touched.get(name) and bool(var.get()) != self._label_style_default_pb(name):
                out[name] = bool(var.get())
        return out

    # ------------------------------------------------------------ 清理规则块
    KEEP_FORMAT_CHOICES = (("加粗", "b"), ("斜体", "i"), ("下划线", "u"),
                           ("上标下标", "vertAlign"), ("颜色", "color"))

    def _build_cleanup_box(self, parent: tk.Frame) -> None:
        """清理规则块（用户 2026-09-20）：清除手打分页符／清除空行／正文保留哪些字符格式。
        **都不作用于封面范围**。随预设保存（JSON/Excel 都认，见 `PresetConfig`）。"""
        from essay_tool.engine.cleanup import DEFAULT_KEEP_RUN_FORMATS
        box = tk.Frame(parent, bg=C_PANEL, highlightbackground=C_LINE, highlightthickness=1)
        box.pack(fill="x", pady=(4, 0))
        inner = tk.Frame(box, bg=C_PANEL)
        inner.pack(fill="x", padx=12, pady=(10, 6))

        self.v_clear_pb = tk.BooleanVar(value=False)
        self.v_clear_blank = tk.BooleanVar(value=False)
        self._checkbutton(inner, "清除文档中所有的手打分页符", variable=self.v_clear_pb,
                          font_size=11).pack(anchor="w", pady=(0, 2))
        self._checkbutton(inner, "清除文档中所有的空行", variable=self.v_clear_blank,
                          font_size=11).pack(anchor="w", pady=(0, 2))
        tk.Label(inner, text="分页符指手动插入的分页符（Ctrl+Enter），不含分节符；空行指完全没有内容的段落"
                             "（含有空格、图片、编号等的段落不算）。封面范围内的内容不受影响。",
                 font=(FAMILY, 9), bg=C_PANEL, fg=C_HINT, wraplength=700,
                 justify="left").pack(anchor="w", pady=(2, 8))

        row = tk.Frame(inner, bg=C_PANEL)
        row.pack(fill="x")
        tk.Label(row, text="正文保留的字符格式", font=(FAMILY, 10), bg=C_PANEL,
                 fg=C_TEXT).pack(side="left", padx=(0, 12))
        self.keep_vars: dict[str, tk.BooleanVar] = {}
        for label, code in self.KEEP_FORMAT_CHOICES:
            self.keep_vars[code] = tk.BooleanVar(value=code in DEFAULT_KEEP_RUN_FORMATS)
            self._checkbutton(row, label, variable=self.keep_vars[code],
                              font_size=10).pack(side="left", padx=(0, 14))
        self._keep_extra: frozenset[str] = frozenset()    # 预设里有、但这里没有复选框的（如删除线）
        tk.Label(inner, text="没勾选的字符格式会被清除，改由样式决定。一至四级标题不保留任何字符格式；"
                             "参考文献里的链接文字始终保留颜色与划线。",
                 font=(FAMILY, 9), bg=C_PANEL, fg=C_HINT, wraplength=700,
                 justify="left").pack(anchor="w", pady=(6, 8))

    def _load_cleanup_from(self, cfg) -> None:
        """预设的清理规则 → 清理规则块。"""
        if not hasattr(self, "keep_vars"):
            return
        from essay_tool.engine.cleanup import DEFAULT_KEEP_RUN_FORMATS
        self.v_clear_pb.set(bool(getattr(cfg, "clear_page_breaks", False)))
        self.v_clear_blank.set(bool(getattr(cfg, "clear_blank_lines", False)))
        keep = getattr(cfg, "keep_run_formatting", None)
        keep = frozenset(DEFAULT_KEEP_RUN_FORMATS if keep is None else keep)
        for code, var in self.keep_vars.items():
            var.set(code in keep)
        self._keep_extra = keep - set(self.keep_vars)

    def _collect_keep_run_formatting(self, fallback):
        if not hasattr(self, "keep_vars"):
            return fallback
        return frozenset({code for code, var in self.keep_vars.items() if var.get()}
                         | self._keep_extra)

    def _on_heading_gap_edit(self, name: str) -> None:
        self._heading_touched[name] = True
        self._schedule_label_refresh()

    def _heading_label_text(self, name: str) -> str:
        """引言/结论写进预设的模板；空串＝不写（保持文档原文）。

        数字没动过、预设里读入的原模板就原样带回（保住全角空格等写法）；预设里带着
        一个数字表达不了的写法（如"【引言】"）且这格没被用户动过，也原样带着走，
        不能静默丢。"""
        from essay_tool.config.schema import heading_label_with_gap
        text = self.heading_gap_vars[name].get().strip()
        raw, raw_gap = self._heading_raw.get(name, (None, None))
        if text == "":
            untouched_odd = raw and raw_gap is None and not self._heading_touched.get(name)
            return raw if untouched_odd else ""
        gap = int(text)
        return raw if (raw and raw_gap == gap) else heading_label_with_gap(name, gap)

    def _schedule_label_refresh(self) -> None:
        """标签写法改了：结构页显示的是"导出时的写法"，要跟着重铺（合并成一次，
        应用预设时一口气改十几格也只重铺一遍）。"""
        if getattr(self, "_label_refresh_job", None) is not None:
            return
        self._label_refresh_job = self.after_idle(self._refresh_structure_labels)

    def _refresh_structure_labels(self) -> None:
        self._label_refresh_job = None
        inst = getattr(self, "_adjust_instance", None)
        if inst is not None:
            try:
                inst.set_insert_text(self._structure_insert_text())
            except Exception:                       # noqa: BLE001
                pass                                # 结构页已关/在重建：下次打开自然是新的

    def _load_labels_from(self, cfg) -> None:
        """预设的"插入文字" → 标签写法块。预设里**没有这个键**的部件用默认写法
        （老预设缺英文关键词那一格之类）；键在但值为空＝用户特意清空，就留空。"""
        if not hasattr(self, "label_vars"):
            return
        defaults = {n: txt for n, _on, txt, _note in STRUCTURE_PARTS}
        saved = getattr(cfg, "structure_insert_text", None) or {}
        for name, var in self.label_vars.items():
            var.set(saved[name] if name in saved else defaults.get(name, ""))
        from essay_tool.config.schema import heading_label_gap
        for name, var in self.heading_gap_vars.items():
            raw = saved.get(name) or None
            gap = heading_label_gap(raw) if raw else None
            var.set("" if gap is None else str(gap))
            self._heading_raw[name] = (raw, gap)
            self._heading_touched[name] = False       # 上面 set 会触发"被动过"，这里复位
        # "标签前另起一页"：预设里写了的用预设的（视为动过），没写的跟随样式表
        override = getattr(cfg, "label_page_break", None) or {}
        for name, var in getattr(self, "label_pb_vars", {}).items():
            if name in override:
                var.set(bool(override[name]))
                self._label_pb_touched[name] = True
            else:
                self._label_pb_touched[name] = False
        self._sync_label_pb_defaults()

    def _load_citation_from(self, cfg) -> None:
        """把预设「引注」小节填进脚注块的编号区（写法不在候选里就补进去）。"""
        cit = getattr(cfg, "citation", None) if cfg is not None else None
        if cit is None:
            return
        for key, val in (("fn_style", cit.footnote_marker_style),
                         ("fn_font", cit.footnote_marker_font),
                         ("fn_space", cit.footnote_marker_trailing_space),
                         ("ref_style", cit.reference_numbering_style)):
            cb = self.cv_cb[key]
            vals = list(cb.cget("values"))
            if val and val not in vals:
                cb.config(values=vals + [val])
            if val is not None:
                self.cv[key].set(val)
        self.cv["fn_restart"].set(bool(cit.footnote_restart_each_page))

    # ------------------------------------------------------------ 收集当前设置

    def _collect_style_spec(self, rv: dict):
        """格式表格里一行控件的当前值 → ParagraphStyleSpec。

        直接复用 `units.py` 的 parse_* 函数：表格里显示的字符串（"小四"、
        "居中"、"固定值（磅）"……）跟 Excel 认的是同一套写法，这样"保存当前
        设置"和"读 Excel"两条路径共用同一份校验逻辑，不用维护两份。
        """
        from essay_tool.config import units
        from essay_tool.config.schema import ParagraphStyleSpec

        line_rule = units.parse_line_rule(rv["lr"].get())
        line_value = units.parse_line_value(line_rule, rv["lv"].get())
        return ParagraphStyleSpec(
            id=rv["id"],
            name=rv["name"],
            font_east_asian=rv["cn_font"].get().strip(),
            font_ascii=rv["en_font"].get().strip(),
            size=units.parse_font_size(rv["size"].get()),
            bold=bool(rv["bold"].get()),
            alignment=units.parse_alignment(rv["align"].get()),
            first_line_indent=units.parse_indent_pair(rv["fl_unit"].get(), rv["fl_val"].get()),
            left_indent=units.parse_indent_pair(rv["li_unit"].get(), rv["li_val"].get()),
            space_before=units.parse_optional_length_pair("磅", rv["sb"].get()),
            space_after=units.parse_optional_length_pair("磅", rv["sa"].get()),
            line_rule=line_rule,
            line_value=line_value,
            page_break_before=bool(rv["pb"].get()),
        )

    def _label_texts(self) -> dict[str, str]:
        """预设「结构部件」里的"插入文字"（标签写法）＝格式页「标签写法」块当前
        的值；块还没建好时用 `STRUCTURE_PARTS` 的默认写法。

        **内容占位语不许自定义**：题目/副标题/英文题目/英文副标题这四格恒为空。
        引言/结论没有"启用"开关，界面上只有"两字之间空几格"一个数字，由它生成
        写法模板；留空就不写进预设（＝保持原文）。
        """
        texts = {n: txt for n, _on, txt, _note in STRUCTURE_PARTS}
        for name, var in getattr(self, "label_vars", {}).items():
            texts[name] = var.get()
        for name in getattr(self, "heading_gap_vars", {}):
            text = self._heading_label_text(name)
            if text:
                texts[name] = text
        return texts

    def _header_line_extend(self, key: str, side: str) -> float:
        """页面页签里"装饰线向外延伸"的一格 → 厘米数（空＝0；不许负数）。"""
        from essay_tool.config import units

        text = self.pv[key].get().strip()
        if not text:
            return 0.0
        value = round(units.parse_cm(text).cm, 2)
        if value < 0:
            raise ValueError(f"页眉装饰线{side}延伸不能是负数：{text}")
        return value

    def _collect_current_preset_config(self):
        """把「页面」「格式」「结构」「封面」几个页签当前的表单值收成一份
        PresetConfig——「保存当前设置为新配置」用这个。

        「引注」块（脚注/参考文献编号设置）和结构部件表都各自在自己的页签
        里管，这里一并收进来；`_preset_extra` 里记的是表单管不到的几项
        （清理规则要保留的字符格式）——没有配置来源时
        （比如还没应用过任何预设就点了保存）就用代码默认值兜底。
        """
        from essay_tool.config import units
        from essay_tool.config.schema import (
            CitationConfig,
            HeaderFooterRule,
            PageSetup,
            PresetConfig,
        )

        paper_choice = self.cb_paper.get().strip()
        if paper_choice == "自定义":
            page_width = units.parse_cm(self.pv["pw"].get())
            page_height = units.parse_cm(self.pv["ph"].get())
        elif paper_choice in units.NAMED_PAPER_SIZES:
            page_width, page_height = units.NAMED_PAPER_SIZES[paper_choice]
        else:
            raise ValueError(f'无法识别的纸张规格："{paper_choice}"')
        extra = self._preset_extra
        page = PageSetup(
            paper_size=paper_choice,
            page_width=page_width,
            page_height=page_height,
            orientation=self.cb_orient.get().strip(),
            margin_top=units.parse_cm(self.pv["mt"].get()),
            margin_bottom=units.parse_cm(self.pv["mb"].get()),
            margin_left=units.parse_cm(self.pv["ml"].get()),
            margin_right=units.parse_cm(self.pv["mr"].get()),
            gutter=units.parse_cm(self.pv["gutter"].get()),
            header_distance=units.parse_cm(self.pv["hd"].get()),
            footer_distance=units.parse_cm(self.pv["fd"].get()),
        )

        styles: dict = {}
        styles_by_name: dict = {}
        for rv in self.fmt_rows:
            spec = self._collect_style_spec(rv)
            styles[spec.id] = spec
            styles_by_name[spec.name] = spec

        header_footer: list[HeaderFooterRule] = []
        for i, hv in enumerate(self.hdr_vars):
            fv = self.ftr_vars[i] if i < len(self.ftr_vars) else {}
            enabled = bool(fv["num"].get()) if fv else False
            fmt = None
            if enabled:
                fmt_label = fv["fmt"].get()
                fmt = PAGE_FMT_LABEL_REV.get(fmt_label, fmt_label)
            header_footer.append(HeaderFooterRule(
                section_name=hv.get("section", "?"),
                header_text=hv["header"].get(),
                page_number_enabled=enabled,
                page_number_format=fmt,
                page_number_prefix=fv["pre"].get() if fv else "",
                page_number_suffix=fv["suf"].get() if fv else "",
            ))

        citation = CitationConfig(
            footnote_marker_style=self.cv["fn_style"].get(),
            footnote_restart_each_page=bool(self.cv["fn_restart"].get()),
            footnote_marker_font=self.cv["fn_font"].get(),
            footnote_marker_trailing_space=self.cv["fn_space"].get(),
            reference_numbering_style=self.cv["ref_style"].get(),
        )

        # 「结构」不再有自己的表单——启用状态直接读调整表格当前的 ENABLED
        # （STRUCTURE_PARTS 的默认值垫底，覆盖引言/结论这类不在 ENABLED 里的
        # 部件）；标签写法（"插入文字"）见 `_label_texts`。
        from essay_tool.gui.adjust_preview import ENABLED
        structure_required: dict[str, bool] = {n: on for n, on, _txt, _note in STRUCTURE_PARTS}
        structure_required.update(ENABLED)
        structure_insert_text = self._label_texts()

        return PresetConfig(
            page=page,
            styles=styles,
            styles_by_name=styles_by_name,
            header_footer=header_footer,
            citation=citation,
            structure_required=structure_required,
            structure_insert_text=structure_insert_text,
            keep_run_formatting=self._collect_keep_run_formatting(extra.get("keep_run_formatting")),
            clear_page_breaks=bool(getattr(self, "v_clear_pb", None) and self.v_clear_pb.get()),
            clear_blank_lines=bool(getattr(self, "v_clear_blank", None) and self.v_clear_blank.get()),
            label_page_break=self._collect_label_page_break(),
            header_bottom_line=bool(self.v_headerline.get()),
            header_line_extend_left_cm=self._header_line_extend("hl", "左"),
            header_line_extend_right_cm=self._header_line_extend("hr", "右"),
            cover_hanging_indent=bool(self.cv_cover["hanging"].get()),
            # 封面开关与形态跟着预设存（2026-09-14）；填写内容不存。
            cover_enabled=bool(self.cv_cover["on"].get()),
            cover_form=self.cv_cover["form"].get(),
        )

    # ------------------------------------------------------------ 配置文件列表

    def _load_preset_rows(self) -> list[dict]:
        """配置列表：内置只读预设（随程序分发，presets/*.json）在前，用户自己
        保存/从 Excel 导入的预设（%APPDATA%，可删）在后——JSON 是默认存盘
        格式，Excel 只在"导出/导入"两个动作里出现，不参与这份列表。
        """
        rows: list[dict] = []
        seen: set[str] = set()
        for d, builtin in ((builtin_presets_dir(), True), (user_presets_dir(), False)):
            if not d.exists():
                continue
            for p in sorted(d.glob("*.json"), key=_preset_sort_key):
                if p.stem in seen:
                    continue
                seen.add(p.stem)
                rows.append({"path": p, "builtin": builtin})
        if not any(r["path"].stem == BUILTIN_COURSE for r in rows):
            rows.append({"path": None, "builtin": True, "pending": True,
                         "label": BUILTIN_COURSE + "（待创建）"})
        return rows

    def _render_preset_rows(self) -> None:
        for child in self.rows_box.winfo_children():
            child.destroy()
        self._apply_btns: list[tk.Button] = []   # 流程图「选择预设」闪的就是它们
        for i, row in enumerate(self._preset_rows, start=1):
            rf = tk.Frame(self.rows_box, bg=C_PANEL)
            rf.pack(fill="x", pady=2)
            tk.Label(rf, text=str(i), width=3, font=(FAMILY, 10), bg=C_PANEL,
                     fg=C_TEXT).pack(side="left")
            if row.get("pending"):
                name = row["label"]
            else:
                name = row["path"].stem
                if row["path"] == self.applied_preset:
                    name += "（已应用）"
            tk.Label(rf, text=name, font=(FAMILY, 10, "bold" if
                     row["path"] is not None and row["path"] == self.applied_preset
                     else "normal"),
                     bg=C_PANEL, fg=C_HINT if row.get("pending") else C_TEXT,
                     anchor="w").pack(side="left", fill="x", expand=True, padx=(2, 8))
            for text, cmd, on in (
                ("导出为Excel", lambda r=row: self.on_export_preset(r), not row.get("pending")),
                ("删　除", lambda r=row: self.on_delete_preset(r),
                 not row.get("pending") and not row.get("builtin")),
                ("应　用", lambda r=row: self.on_apply_preset(r), not row.get("pending")),
            ):
                b = tk.Button(rf, text=text, font=(FAMILY, 9), padx=6,
                              command=cmd if on else None)
                if not on:
                    b.config(state="disabled")
                b.pack(side="left", padx=2)
                if text == "应　用" and on:
                    self._apply_btns.append(b)
        self._fit_preset_box()

    def _fit_preset_box(self) -> None:
        """按行数定配置列表的高度：最多 PRESET_LIST_MAX_H，超了出滚动条。

        用户 2026-09-16：「添加超过空位的预设时，将会将窗口无限拉长，直到将
        下面的按钮挤没。」每次重画行之后都要重新量一次。
        """
        canvas = getattr(self, "preset_canvas", None)
        if canvas is None:
            return
        self.rows_box.update_idletasks()
        need = self.rows_box.winfo_reqheight()
        canvas.configure(height=max(PRESET_LIST_MIN_H, min(need, PRESET_LIST_MAX_H)),
                         scrollregion=(0, 0, canvas.winfo_width(), need))
        self._sb_state(self.preset_sb, *canvas.yview())
        self._bind_preset_wheel()

    def _bind_preset_wheel(self) -> None:
        """滚轮也能滚配置列表（装得下时 `_scroll_canvas` 自己就不动）。"""
        canvas = getattr(self, "preset_canvas", None)
        if canvas is None:
            return

        def wheel(event):
            notch = wheel_notch(event)
            self._scroll_canvas(canvas, -2 * notch, True)
            return "break"

        def wire(widget) -> None:
            widget.bind("<MouseWheel>", wheel)
            for child in widget.winfo_children():
                wire(child)

        wire(self.rows_box)

    # ------------------------------------------------------------ 动作

    def status(self, msg: str) -> None:
        """写一条提示：追加进下方**日志区**（可上下翻历史），行首自带 `[时:分:秒]`。

        全工具几十处提示都走这一个口子，所以日志区不用各调用点自己接。

        （2026-09-15：原来还有一个 `log=False` 的"瞬时提示"分支，审计核实全仓
        已经 **0 个调用点**——它原来只写日志区最下那条窄栏，而窄栏从 2026-09-15
        起只显示"日志折叠时的最近一条"（见 `_refresh_log_bar`），不再收临时
        提示。参数既然没人用、也没有去处，一并删掉，免得以后有人以为写了会有用。）
        """
        self._log_append(msg)

    def _refresh_log_bar(self) -> None:
        """刷新窄栏：**日志展开时留空，折叠时显示最近一条日志**。

        用户 2026-09-15 重设了这个位置的功能：它不再是"最新提示"栏，只是
        日志区折叠之后的一个"剥一眼"——想看全部就展开日志。
        """
        if self._log_open:
            self.lbl_status.config(text="")
        else:
            self.lbl_status.config(text=self._log_raw[-1] if self._log_raw else "")

    def _dialog_dir(self) -> str:
        """文件对话框的默认文件夹：上次用过的（导入后就是原文件所在处），没有就用"文档"。
        不再用程序自己的目录——打包后那里通常是只读的，也不是用户找文件的地方。"""
        last = getattr(self, "_last_dir", None)
        if last and pathlib.Path(last).is_dir():
            return last
        docs = pathlib.Path.home() / "Documents"
        return str(docs if docs.is_dir() else pathlib.Path.home())

    def on_import_file(self) -> None:
        path = filedialog.askopenfilename(
            title="导入论文（.docx / .txt）", initialdir=self._dialog_dir(),
            filetypes=[("Word 文档", "*.docx"), ("纯文本", "*.txt"),
                       ("全部文件", "*.*")])
        if not path:
            return
        self._import_path(pathlib.Path(path))

    def _on_files_dropped(self, paths: list[str]) -> None:
        """资源管理器拖了文件进来（回调在窗口过程里，弹窗/导入一律挪到主循环里做）。"""
        self.after(10, lambda: self._import_dropped(paths))

    def _import_dropped(self, paths: list[str]) -> None:
        picked = [pathlib.Path(p) for p in paths
                  if pathlib.Path(p).is_file() and pathlib.Path(p).suffix.lower() in (".docx", ".txt")]
        if not picked:
            self.status("拖入的文件不是 .docx 或 .txt，已忽略。")
            return
        if len(paths) > 1:
            self.status(f"一次只能导入一个文件，已取第一个可用的：{picked[0].name}")
        self._import_path(picked[0])

    def _import_path(self, chosen: pathlib.Path) -> None:
        """导入 `chosen`（按钮选的、拖进来的都走这里）；已经导入过文件时先确认。"""
        if self.doc_path is not None and not messagebox.askokcancel(
                "导入文件", "继续导入将覆盖已有文件及其编辑，确定继续吗？",
                parent=self, icon="warning"):
            return
        # 先记下"改之前"的样子——.txt 转换失败时要能整个退回去，不能让上一份
        # 文档凭空消失（2026-09-22 审计揪出的真 bug：以前失败时只把 `doc_path`
        # 置空，`src_path`／导入区文字都已经指向了这份读不了的新文件，上一份
        # 文档的引用彻底丢了，界面却还显示着它的文件名，状态对不上）。
        prev_doc_path, prev_src_path = self.doc_path, self.src_path
        prev_label_text = self.lbl_doc.cget("text")
        self._last_dir = str(chosen.parent)      # 之后的"导出"对话框默认落在原文件所在文件夹
        self.src_path = chosen
        self.doc_path = chosen
        if chosen.suffix.lower() == ".txt":
            # `.txt` 不能直接交给 python-docx 读（结构页/告警会当场抛异常，
            # 实测"识别失败"），而导出路径一直靠 loader 转成 docx。用户
            # 2026-09-16 裁决："让它真支持：导入时转成 docx 再识别"——
            # 这里转一次，之后结构页、告警、导出看到的是同一份 docx 副本。
            from essay_tool.engine.loader import LoadError, load_as_working_copy
            try:
                self.doc_path = load_as_working_copy(chosen, session_work_dir())
            except Exception as exc:
                # ⚠️ 以前只抓 `LoadError`——编码认不出来这类"意料之中"的失败
                # 确实是 `LoadError`，但文件被占用、权限不够、磁盘满这类意外
                # 会以别的异常类型（`OSError` 等）一路冒到兜底崩溃弹窗，点
                # "关闭"整个程序退出。这里统一接住，退回"改之前"那份文档
                # （没有就是真的退回"未导入"），不留半新不旧的状态。
                self.doc_path, self.src_path = prev_doc_path, prev_src_path
                self.lbl_doc.config(text=prev_label_text)
                detail = str(exc) if isinstance(exc, LoadError) else f"{type(exc).__name__}：{exc}"
                messagebox.showerror(
                    "导入失败", f"文本文件读取失败：{detail}", parent=self)
                return
        self.lbl_doc.config(text=chosen.name, fg=C_TEXT)
        self.btn_cancel_import.config(state="normal")
        # 换了一份新文档，之前留下的身份地图不再适用（段落下标对不上新
        # 文档）——重新按这份新文档建表，「论文结构」页签会自动刷新。
        self._rebuild_struct_adjust()
        self.status(f"已导入：{chosen.name}")
        note = self._backup_imported_file(chosen)
        if note:
            # 用户 2026-09-14：备份单独占一行（日志区里能上下翻，不会被导入
            # 那条盖掉）
            self.status(note)
        # 导入的时候不识别、不报告警（方案 3）；但如果用户本来就站在「结构」
        # 页签上，导入之后他就已经在看这一页了，那就当场识别一遍。
        self._on_tab_changed()

    def _on_pick_heading_scheme(self, value: str) -> None:
        """点了某个复选框：**强制单选**——点已选中的那个不会取消，只是保持选中。"""
        self.var_heading_scheme.set(value)
        self._sync_scheme_checks()          # 点已选中的那个时 var 没变、不触发 trace，手动同步
        self._on_heading_scheme_changed()

    def _sync_scheme_checks(self) -> None:
        """两个复选框镜像 `var_heading_scheme`（撤回、取消导入等任何改它的地方都走这里）。"""
        current = self.var_heading_scheme.get()
        for value, flag in getattr(self, "_scheme_checks", {}).items():
            flag.set(value == current)

    def _on_heading_scheme_changed(self) -> None:
        """导入区「原文标题序号」被点了：按新方案整份重新识别。

        跟结构页里的「重新识别」一样，这是一次**没确认的调整**：结构页里
        Ctrl+Z 可以撤回（单选会跟着拨回去，见 `_on_heading_scheme_restored`），
        点了「确认」才作数（用户 2026-09-19："只有点击确定才写入"）。
        """
        scheme = self.var_heading_scheme.get()
        if scheme == self._heading_scheme_applied:
            return
        self._heading_scheme_applied = scheme
        inst = self._adjust_instance
        if inst is not None and inst.doc_path is not None and inst.doc_path == self.doc_path:
            inst.set_heading_scheme(scheme)
        self.status("已切换原文标题序号："
                    + ("中文序号" if scheme == "chinese" else "阿拉伯数字序号"))
        # 本来就站在「结构」页签上的话，识别结果变了，告警也得重报一遍
        # （`_doc_fingerprint` 已经把方案算进去了）。
        if self.doc_path is not None:
            self._on_tab_changed()

    def _on_heading_scheme_restored(self, scheme: str | None) -> None:
        """结构页 Ctrl+Z 撤回了"切换标题编号方式"：把导入区的单选拨回去。"""
        scheme = scheme or DEFAULT_HEADING_SCHEME
        self._heading_scheme_applied = scheme
        self.var_heading_scheme.set(scheme)

    def on_cancel_import(self) -> None:
        """「取消导入」：把界面退回**尚未导入文件**的状态（用户 2026-09-17）。

        只清"**这份文档**带来的东西"：识别结果、结构页的调整、文件名、导出开关。
        **预设与格式页的设置不动**（那是"你要怎么排"，不属于这份稿子；一并清掉的话
        用户还得重新挑一遍预设，反而更麻烦）。封面填写区同理不动。
        """
        if self.doc_path is None:
            self.status("还没有导入文件。")
            return
        if not messagebox.askokcancel(
                "取消导入", "将清除已有排版。", parent=self, icon="warning"):
            return
        self.doc_path = None
        self.src_path = None
        self.lbl_doc.config(text="（尚未选择）", fg=C_HINT)
        # 结构页那两份"本次驻留报过/记过"的账要清掉：换一份文档就该重新报
        self._review_logged = None
        self._review_popped = None
        self._rebuild_struct_adjust()
        # 导出按钮的可用状态**只由「我已阅读并同意」勾选框决定**（`_on_agree`）。原来这里
        # 顺手 `_set_exports_enabled(False)`：勾选框还打着勾、按钮却灰了，用户得先取消
        # 勾选再重新勾一次才恢复（2026-09-20 用户反馈"偶现导出按钮变灰、要重新勾选协议"）。
        self.btn_cancel_import.config(state="disabled")
        self.status("已取消导入：识别结果与结构调整已清除。")

    def _backup_imported_file(self, path: pathlib.Path) -> str:
        """按「备份设置」区当前的开关，把导入的文件复制一份到备份路径。

        这是用户误删自己原始文件时的安全网，不是正常处理流程的一部分——
        工具本身从不修改用户的原始文件（协议一.3），这份备份纯粹防手滑。
        备份/清理失败（比如磁盘满了）不应该拦住导入本身，所以单独
        try/except，出错只在状态栏里带一句，不弹窗、不中断。
        """
        if not self.gui_settings.backup_enabled:
            return ""
        try:
            from essay_tool.engine.loader import backup_import, prune_backups
            backup_dir = self._effective_backup_dir()
            dest = backup_import(path, backup_dir)
            note = f"已经备份原文件到 {dest}"
            if self.gui_settings.backup_auto_prune:
                removed = prune_backups(backup_dir, self.gui_settings.backup_keep_count,
                                        self.gui_settings.backup_keep_unit)
                if removed:
                    note += f"（已清理 {removed} 份过旧备份）"
            return note
        except OSError as exc:
            return f"备份原文件失败：{exc}（不影响导入）"

    def _builtin_preset_names(self) -> set[str]:
        """内置配置的名字（含还没落盘的「课程考核（待创建）」）。

        用户 2026-09-16：「虽然新增预设之间不能重名，但新增预设和内置预设
        （论文和考核）之间却可以重名，这不好。」内置那份在 presets/ 里、是
        随程序分发的只读文件，根本盖不掉；而且 _load_preset_rows 遇到同名
        会跳过用户那份 ⇒ 存下去的结果是"列表里根本看不到它"。所以直接不让
        起这个名字。
        """
        names: set[str] = set()
        for row in self._preset_rows:
            if not row.get("builtin"):
                continue
            if row.get("path") is not None:
                names.add(row["path"].stem)
            else:                               # 「……（待创建）」那种占位行
                names.add(str(row.get("label", "")).split("（")[0])
        names.add(BUILTIN_COURSE)
        return names

    def on_save_preset(self) -> None:
        """把当前表单里的设置存成一份新的 JSON 配置（本工具默认的存盘格式）。"""
        try:
            cfg = self._collect_current_preset_config()
        except Exception as exc:
            self.status(f"保存失败：设置内容有误（{exc}）")
            return
        default_name = self.applied_preset.stem if self.applied_preset else "我的配置"
        name = simpledialog.askstring("保存配置", "给这份配置起个名字：",
                                      initialvalue=default_name, parent=self)
        if name is None:
            return
        safe_name = "".join(c for c in name.strip() if c not in '\\/:*?"<>|')
        if not safe_name:
            self.status("名称不能为空，且不得包含 \\ / : * ? \" < > | 等字符。")
            return
        if safe_name in self._builtin_preset_names():
            self.status(f"「{safe_name}」是内置配置的名字，内置配置不能被覆盖，请换个名字。")
            return
        out_path = user_presets_dir() / f"{safe_name}.json"
        if out_path.exists() and not messagebox.askyesno(
                "覆盖确认", f"“{safe_name}”已经存在，要覆盖它吗？"):
            return
        try:
            from essay_tool.config.preset_io import save_preset_json
            save_preset_json(cfg, out_path)
        except Exception as exc:
            self.status(f"保存失败：{exc}")
            return
        if not any(r.get("path") == out_path for r in self._preset_rows):
            self._preset_rows.append({"path": out_path, "builtin": False})
        self.applied_preset = out_path
        self._preset_extra = {
            key: self._preset_extra.get(key) for key in ("keep_run_formatting",)
        }
        self._render_preset_rows()
        self.status(f"配置已保存：{safe_name}（{out_path}）")

    def on_import_preset(self) -> None:
        """从 Excel 导入配置文件——Excel 是边缘产品，读进来之后转存成 JSON
        落盘（跟「保存当前设置为新配置」走的是同一套本地存储）。"""
        path = filedialog.askopenfilename(
            title="从 Excel 导入配置文件", initialdir=str(PRESET_DIR),
            filetypes=[("Excel 工作簿", "*.xlsx")])
        if not path:
            return
        xlsx_path = pathlib.Path(path)
        name = simpledialog.askstring(
            "导入配置", "给这份配置起个名字：",
            initialvalue=xlsx_path.stem, parent=self)
        if name is None:
            return
        safe_name = "".join(c for c in name.strip() if c not in '\\/:*?"<>|')
        if not safe_name:
            self.status("名称不能为空，且不得包含 \\ / : * ? \" < > | 等字符。")
            return
        if safe_name in self._builtin_preset_names():
            self.status(f"「{safe_name}」是内置配置的名字，内置配置不能被覆盖，请换个名字。")
            return
        out_path = user_presets_dir() / f"{safe_name}.json"
        if out_path.exists() and not messagebox.askyesno(
                "覆盖确认", f"“{safe_name}”已经存在，要覆盖它吗？"):
            return
        try:
            from essay_tool.config.preset_io import import_from_excel
            import_from_excel(xlsx_path, out_path)
        except Exception as exc:
            self.status(f"导入失败：{exc}")
            return
        if not any(r.get("path") == out_path for r in self._preset_rows):
            self._preset_rows.append({"path": out_path, "builtin": False})
            self._render_preset_rows()
        self.status(f"配置已导入：{safe_name}（{xlsx_path.name} → {out_path.name}）")

    def on_apply_preset(self, row: dict) -> None:
        if row.get("pending") or row["path"] is None:
            return                      # 占位行本来没按钮；真走到这里就静默退出
        try:
            # 这一行的 import 2026-09-14 被误删过一次（删"流程图进度"那几行时
            # 连带删掉的），结果点「应　用」一律报"参数读取失败"——留在方法
            # 里重新补上，别再动。
            from essay_tool.config.preset_io import load_preset
            cfg = load_preset(row["path"])
            self._load_page_from(cfg, row["path"].stem)
            self._load_formats_from(cfg, row["path"].stem)
            self._load_cover_from(cfg, row["path"].stem)
            # 「结构」不属于"应用预设"这一步——那是文档自己的事，导入的时候
            # 已经按"预设∪识别到的部件"定过初始值了，重新应用预设不能把
            # 用户已经做的调整冲掉。但很多人是先导入、后选预设，那时候还
            # 没有预设可参考，这里补一次"只点亮、不熄灭"的合并，让预设的
            # 结构设定不至于被完全无视（见 merge_enabled_from）。
            if self._adjust_instance is not None and self._adjust_instance.doc_path == self.doc_path:
                from essay_tool.gui.adjust_preview import merge_enabled_from
                merge_enabled_from(cfg.structure_required)
                # 预设里「插入文字」那一格也换了部件写法（摘要→【摘要】之类），
                # 结构页显示的是"导出时的写法"，所以得跟着重铺一遍
                # （用户 2026-09-15）。
                self._adjust_instance.set_insert_text(cfg.structure_insert_text)
                self._adjust_instance.draw()
            # "已应用"的状态**读预设成功之后**才记（2026-09-15 审计核实：以前
            # `applied_preset` 在 `try` 之后、列表重绘也排在前面，读预设
            # 抛异常时列表已经显示"（已应用）"、而几个页签的表单可能只加载
            # 了一半——界面说成功了、实际是半成品）。
            self.applied_preset = row["path"]
            self._render_preset_rows()
            self.status("配置应用成功！")
        except Exception as exc:
            self.status(f"配置应用失败：{row['path'].stem}，参数读取失败（{exc}）")

    def on_delete_preset(self, row: dict) -> None:
        if row.get("builtin"):
            self.status("内置配置不可删除。")
            return
        path = row["path"]
        if not messagebox.askyesno("删除配置", f"确定删除“{path.stem}”？此操作不可撤销。"):
            return
        try:
            path.unlink(missing_ok=True)
        except OSError as exc:
            self.status(f"删除失败：{exc}")
            return
        self._preset_rows = [r for r in self._preset_rows if r.get("path") != path]
        if self.applied_preset == path:
            self.applied_preset = None
        self._render_preset_rows()
        self.status(f"配置已删除：{path.stem}")

    def on_export_preset(self, row: dict) -> None:
        if row.get("pending"):
            return                      # 占位行没有“导出为 Excel”按钮
        try:
            from essay_tool.config.preset_io import export_to_excel, load_preset
            cfg = load_preset(row["path"])
        except Exception as exc:
            self.status(f"导出失败：读取 {row['path'].stem} 时出错（{exc}）")
            return
        out = filedialog.asksaveasfilename(
            title="导出为 Excel", defaultextension=".xlsx",
            initialfile=f"{row['path'].stem}.xlsx", initialdir=self._dialog_dir())
        if not out:
            return
        try:
            export_to_excel(cfg, out)
        except Exception as exc:
            self.status(f"导出失败：{exc}")
            return
        self.status(f"已导出为 Excel：{pathlib.Path(out).name}")

    def _on_agree(self) -> None:
        self._set_exports_enabled(self.var_agree.get())

    def _set_exports_enabled(self, on: bool) -> None:
        state = "normal" if on else "disabled"
        self.btn_exp_blank.config(state=state)
        self.btn_exp_doc.config(state=state)

    def _current_cover_type(self) -> str:
        """封面页签当前选的形态 → 引擎的封面类型键（thesis/year/course/tenstar）。"""
        from essay_tool.config.schema import COVER_TYPE_BY_FORM
        return COVER_TYPE_BY_FORM[self.cv_cover["form"].get()]

    def _current_cover_fields(self):
        """把「封面」页签"内容"区当前填的值收成一份 CoverFields。"""
        from essay_tool.engine.cover import CoverFields
        values = {k: v.get() for k, v in self.cv_cover_fields.items()}
        return CoverFields(**values)

    def _current_session(self) -> str:
        """封面页签「届数」框当前填的值（预设页眉里的 `{届数}` 用它替换）。"""
        var = self.cv_cover_fields.get("session")
        return var.get().strip() if var is not None else ""

    def on_export_blank(self) -> None:
        # 用户 2026-09-14：不要求先“应用预设”——没选预设时页签里就是默认
        # 配置（启动时已经照内置毕业论文填好），照样允许导出。
        out = filedialog.asksaveasfilename(
            title="导出空白模板到……", defaultextension=".docx",
            initialfile="空白模板.docx", initialdir=self._dialog_dir())
        if not out:
            return
        out_path = pathlib.Path(out)
        with_cover = bool(self.cv_cover["on"].get())
        cover_type = self._current_cover_type()
        self.status(f"正在生成空白模板：{out_path.name}…")
        self.update_idletasks()
        try:
            # 跟 on_export_doc 同一个道理：用页签当场的表单值，不是文件快照。
            preset_config = self._collect_current_preset_config()
            from essay_tool.engine.alerts import AlertCollector
            from essay_tool.engine.template import build_template
            alerts = AlertCollector()
            build_template(None, out_path, work_dir=session_work_dir(),
                           with_cover=with_cover, cover_type=cover_type,
                           cover_fields=self._current_cover_fields() if with_cover else None,
                           preset_config=preset_config, alerts=alerts,
                           session=self._current_session())
        except Exception as exc:
            self.status(f"生成空白模板失败：{exc}")
            return
        # 预设名：2026-09-14 起不强制先“应用预设”，没应用就说“当前设置”
        preset_label = self.applied_preset.stem if self.applied_preset else "当前设置"
        msg = (f"空白模板已生成：{out_path.name}"
               f"（预设：{preset_label}"
               f"{'；含' + self.cv_cover['form'].get() if with_cover else ''}）")
        self.status(msg)
        self._report_alerts(alerts)

    def _report_alerts(self, alerts) -> None:
        """把引擎报的提醒**逐条**写进日志。

        用户 2026-09-14 要求：不能只报"有几条提醒"——具体错误和位置都要写出来。
        位置就是结构页"段号"列里的那个段号（`Alert.line` 拼好的，没有对应
        段落的那些不写位置）。

        **不弹窗**（用户 2026-09-16："导出时不要再次提醒"）：告警的弹窗只在
        切到「结构」页签那一刻弹一次（`_on_tab_changed`），导出时再弹一遍是
        重复打断；这里只把同样的内容逐条写进日志。

        **没有告警时什么都不说**（审计 2026-09-15 问过，这里写明是故意的）：
        日志里上方已经有"论文已生成：xxx.docx"，再补一句"未发现提醒"会把日志
        变噪，与用户 2026-09-14 "只在有东西可报时才出现"的口径一致。
        """
        if not alerts.items:
            return
        for line in [a.line for a in alerts.items]:
            self.status(line)

    def _on_tab_changed(self, event=None) -> None:
        """切到「结构」页签时识别一遍，有告警就**弹窗**（用户 2026-09-15 定）。

        方案 3 + 弹窗：导入文件时**不**识别；打开结构页才算"开始核对结构"，
        这一刻把工具从原文里读到的问题列出来。

        **为什么弹窗而不是只写日志**（用户 2026-09-15）：只写日志的话，来回
        切几次页签下方日志就迅速被挤满，而且那些行跟其他提示混在一起，本来
        也不显著、很难被注意到。所以：
        * 弹窗每次都弹（一次性、看完就关，不脏任何地方）；
        * 日志**同一份结论只记一次**（`_log_review_once`）——日志是留给
          "回头再看"的，同一批内容反复追加没有意义；
        * 弹窗里带一句"后续如需再次查看，可以前往屏幕下方状态窗口"。

        识别走 `pipeline.review()`，跟导出用同一批探测器与规则，所以结构页
        说的话和导出时说的话不会对不上（实测 36～307 段的文档只要 7～120 ms）。

        ⚠️ **一次切页签会弹两次窗**（用户 2026-09-15 实拍，排查结果）：ttk 的
        `<<NotebookTabChanged>>` 一次切换可能投递**两个**事件（切换本身一个、
        状态落定又一个），而弹窗是**模态**的、它自己会跑一个嵌套事件循环——
        嵌套循环里把第二个事件也消化了，于是关掉第一个弹窗立刻又来一个
        （实测：一次"切进来"能弹两次）。两道闸：
          1. **不当场干活**：记一个 `after()` 任务，250ms 窗口内的多个事件
             只兑现一次（弹窗跑在事件处理之外，不会再被嵌套事件撞上）；
          2. **同一次驻留只报一次**（`_review_popped`）——万一第二个事件
             晚到（过了防抖窗口），账还在，也不会再弹；离开这一页就清账，
             下次进来照旧报告。
        """
        if self.nb.select() != str(self.struct_tab):
            # 一离开这一页就清账（不留给防抖任务），下次进来要能重新报告
            self._review_popped = None
            if self._review_job is not None:
                try:
                    self.after_cancel(self._review_job)
                except Exception:                      # noqa: BLE001
                    pass
                self._review_job = None
            return
        if self._review_job is not None:
            try:
                self.after_cancel(self._review_job)
            except Exception:                          # noqa: BLE001
                pass
        self._review_job = self.after(REVIEW_DEBOUNCE_MS, self._run_review)

    def _doc_fingerprint(self) -> str:
        """当前文档的"身份"：路径 + 修改时间 + 大小（取不到就退化成路径）。

        `_review_popped` 用它判"这一次进结构页是不是已经报过"。**不能只比
        路径**：用户可能把同一个文件在外部改过再重新导入——只比路径就会当成
        "刚报过"、不再报告（2026-09-15 核实审计报告时发现，是我自己上一轮
        引入的）。取不到 mtime（文件被删/权限）就退化成路径：宁可多报一次，
        不能漏报。标题编号方案也算进去：换了方案，识别结果和告警都跟着变。
        """
        scheme = self.var_heading_scheme.get()
        try:
            st = self.doc_path.stat()
            return f"{self.doc_path}|{st.st_mtime_ns}|{st.st_size}|{scheme}"
        except OSError:
            return f"{self.doc_path}|{scheme}"

    def _run_review(self) -> None:
        """真正去识别并（需要时）弹窗——由 `_on_tab_changed` 延后叫起来。"""
        self._review_job = None
        if self.doc_path is None:
            return
        if self.nb.select() != str(self.struct_tab):
            # 防抖窗口里用户又切走了：这一刻他不在看结构页，不用报告
            return
        if self._review_popped == self._doc_fingerprint():
            return                       # 这一次进这一页已经报告过了（重复事件）
        fingerprint = self._doc_fingerprint()
        from essay_tool.engine.pipeline import review
        try:
            # 用户手敲的封面范围要带下去（用户 2026-09-15）：封面区间不参与
            # 识别，不带的话弹窗会报出**导出里根本不会有**的告警（实测：封面里
            # 写了句“【摘要】”，弹窗报“重复的部件：摘要”）。
            # （“题注重编”那个参数 2026-09-16 随题注改造一起删掉了。）
            inst = self._adjust_instance
            same_doc = inst is not None and inst.doc_path == self.doc_path
            alerts = review(self.doc_path, self._collect_current_preset_config(),
                            cover_end=(inst.cover_end if same_doc else None),
                            heading_scheme=self.var_heading_scheme.get())
        except Exception as exc:                      # noqa: BLE001
            self.status(f"结构识别失败：{exc}")
            # **不记账**（审计 2026-09-15 核实）：原来在调 review 之前就把指纹
            # 记下了，识别一旦抛异常，账已经记过 ⇒ 用户再切回这一页也不会重试，
            # 屏幕上只剩一行日志，等于警告静默失效。
            return
        self._review_popped = fingerprint
        if not alerts.items:
            self._log_review_once(("结构识别完成：未发现需要确认的问题。",))
            return
        lines = tuple(alert.line for alert in alerts.items)
        self._log_review_once(
            (f"结构识别完成：有 {len(lines)} 处需要确认。", *lines))
        messagebox.showwarning(
            "结构识别结果",
            "\n".join(lines) + "\n\n" + REVIEW_ALERTS_HINT, parent=self)

    def _log_review_once(self, lines: tuple[str, ...]) -> None:
        """结构页的识别结论同一份只写一次日志（弹窗照旧每次都弹）。

        用户 2026-09-15：来回切页签会把下方日志迅速挤满——同一批内容反复
        追加没有意义，日志是留给"回头再看"的。文档换了（`doc_path` 变了）
        或者结论本身变了，算新的一份，照旧写。
        """
        key = (str(self.doc_path), lines)
        if key == self._review_logged:
            return
        self._review_logged = key
        for line in lines:
            self.status(line)

    def _structure_insert_text(self) -> dict[str, str]:
        """结构页"显示成导出时的写法"要用的「插入文字」表（用户 2026-09-15）。

        跟导出同一个来源：当前表单/预设收出来的那份 `PresetConfig`——用户在页签
        里当场改了写法，结构页显示的就是改完的写法。收不出来（还没建好表单、
        或者配置本身有问题）就退回 `STRUCTURE_PARTS` 的第三列，那份就是"不套用
        任何预设"时的默认写法；实在拿不到也没关系，`shown_text()` 拿不到表就
        显示原文。
        """
        try:
            return dict(self._collect_current_preset_config().structure_insert_text)
        except Exception:
            return {name: text for name, _on, text, _note in STRUCTURE_PARTS if text}

    def _structure_reference_style(self) -> str | None:
        """结构页显示参考文献条目序号用的写法（预设「引注」里那个下拉）。

        跟导出同源：导出时 `pipeline.run` 也是从这份配置里取
        `citation.reference_numbering_style`（用户 2026-09-16：用户直接用自动
        编号排参考文献时，条目得显示成"会写成的那一种序号"）。取不到就返回
        None，让 `reference_display_text` 用默认写法（方括号）。
        """
        try:
            return self._collect_current_preset_config().citation.reference_numbering_style
        except Exception:
            return None

    def _build_ops_for_export(self):
        """有对着当前这份文档做过的调整（哪怕调整窗口已经关了）就把结构页
        **上一次点"确认"时**的行序（`inst.confirmed_rows`）换算成
        `(index, ops)`；否则返回 `None`，`pipeline.run()` 走它自己的自动识别
        （`index=None, ops=None`）。

        ⚠️ 这里读的是 `confirmed_rows`，**不是**实时变动的 `inst.rows`
        （用户 2026-09-19："GUI 不点确认，即使点导出也不认"）：拖动/换列/
        调换这些调整只有点过"确认"才算数，导出永远只按最近一次确认的样子
        走。有没有"确认过的调整还没导出"这件事（`inst.dirty`）由
        `on_export_doc` 在调这里之前先问用户，不在这个函数里处理。

        `self.doc_path is None` 时也直接返回 None：这种情况下 `_adjust_instance`
        是空白模板骨架（`blank_mode`），没有真正的文档可换算——`on_export_doc`
        本身也会在 `doc_path is None` 时提前拦下、根本不会走到这里，这条判断
        只是防止两边状态没对齐时（`inst.doc_path` 和 `self.doc_path` 恰好都是
        None）误把空白骨架当成"这份文档调整过"。

        `doc_index.build_ops()` 会**原地**改写 `inst.confirmed_index` 里每条 `Entry` 的
        `.part`/`.meta`——这一步必须在 `pipeline.run()` 之前做完：`run()`
        内部的 `normalize_document`/`apply_from_index`/告警检查都读这份
        `index`，晚了就看不到结构页里的调整（见 `pipeline.run()` 里 `index`
        形参的说明）。`insert_toc`/`list_kinds` 这两个参数要跟 `run()` 自己
        算的完全一致（要不要插目录域、清单该收哪些题注种类），所以这里现开
        一份文档、调同一套 `collect_source_alerts()` 现算一遍，不能凭空猜。
        """
        inst = self._adjust_instance
        if (inst is None or self.doc_path is None or inst.doc_path != self.doc_path
                or inst.confirmed_index is None):
            return None
        try:
            from docx import Document as _Document
            from essay_tool.engine.cleanup import unwrap_section_break_sdt
            from essay_tool.engine.doc_index import (
                build_ops, body_list_field_kinds, body_toc_field_parts)
            from essay_tool.engine.pipeline import collect_source_alerts

            regenerate_fields = frozenset(
                row.block for row in inst.confirmed_rows
                if row.kind == "toc_field" and row.regenerate and row.block)
            doc = _Document(str(inst.doc_path))
            unwrap_section_break_sdt(
                doc, keep_content=bool(body_toc_field_parts(doc) - regenerate_fields))
            preset_config = self._collect_current_preset_config()
            # 行序和识别索引、封面范围都取**上次确认时**的那一份（成对的，见
            # `AdjustWindow.__init__` 里 `confirmed_index` 的说明）。
            _collector, list_plan, _captions = collect_source_alerts(
                document=doc, index=inst.confirmed_index,
                structure_required=preset_config.structure_required,
                existing_list_kinds=body_list_field_kinds(doc),
                cover_end=inst.confirmed_settings["cover_end"],
                regenerate_fields=regenerate_fields,
                include_source_checks=False,
            )
            existing_toc = "目录" in (body_toc_field_parts(doc) - regenerate_fields)
            # `ENABLED` 是结构页跨文档共享的一个实时字典，不能直接读它——不然
            # "有调整还没确认"那句警告（凭 `inst.dirty` 判断）跟实际导出结果
            # 就对不上号：警告说"这次改动不会应用"，实际上未确认的启用/停用
            # 勾选照样会用上（2026-09-22 审计揪出的真 bug）。跟 `confirmed_rows`
            # 一样，只认"上次点确认那一刻"的快照。
            ops = build_ops(inst.confirmed_rows, inst.confirmed_index,
                            enabled=dict(inst.confirmed_enabled),
                            insert_toc=not existing_toc, list_kinds=list_plan)
            return inst.confirmed_index, ops
        except Exception as exc:
            # 用户 2026-09-14：这条不单独报，**合并进“导出失败”**——换算出
            # 你在结构页做的调整就不该静默按自动识别的结果出成品。
            raise RuntimeError(f"调整结果换算失败（{exc}）") from exc

    def on_export_doc(self) -> None:
        if self.doc_path is None:
            self.status("请先导入论文再导出！")
            return
        # 结构页有调整、但还没点"确认"：导出只认最近一次确认的结果，这次
        # 没确认的改动不会生效（用户 2026-09-19 要求，见
        # `adjust_preview.AdjustWindow.on_confirm`）。先问一句，别让调整
        # 悄悄不生效。
        inst = self._adjust_instance
        if (inst is not None and inst.doc_path == self.doc_path and inst.dirty
                and not messagebox.askyesno(
                    "有调整还没确认",
                    "论文结构页里有尚未确认的更改，现在导出将不会应用，"
                    "仍然导出吗？",
                    parent=self)):
            self.status("已取消导出——先在「论文结构」页点「确认」再导出。")
            return
        # 部件顺序跟规范顺序不一致（写死的本科规范顺序，不来自预设）：只提醒，不阻断（2026-09-19
        # 用户要求——"如果导出时的部件顺序和预设要求不一致的，弹窗告警"，
        # GUI 现在允许"无条件调换"，用户完全可能故意把顺序拖成跟规范不
        # 一样，所以只问一句，不是"确认"那种"不点就不让导出"的硬卡点）。
        # **只有已经读入预设时才告警**（用户 2026-09-19：没读入预设就没有
        # "预设要求的顺序"可比；识别之后也不再弹"是否自动排序"）。
        if (self.applied_preset is not None
                and inst is not None and inst.doc_path == self.doc_path):
            from essay_tool.gui.adjust_preview import (
                BACK_ROLES,
                FRONT_ROLES,
                find_order_violations,
            )
            violations = (find_order_violations(inst.confirmed_rows, FRONT_ROLES)
                         + find_order_violations(inst.confirmed_rows, BACK_ROLES))
            if violations:
                lines = "\n".join(f"“{a}”应在“{b}”之前" for a, b in violations)
                if not messagebox.askyesno(
                        "部件顺序跟规范不一致",
                        f"当前部件顺序和规范顺序不一致：\n{lines}\n\n"
                        "可以在论文结构页一键重排，或者手动拖动。\n\n"
                        "仍然导出吗？",
                        parent=self):
                    self.status("已取消导出——顺序跟规范不一致。")
                    return
        out = filedialog.asksaveasfilename(
            title="导出论文到……", defaultextension=".docx",
            initialfile=self.doc_path.stem + "-自动排版.docx",
            initialdir=self._dialog_dir())
        if not out:
            return
        out_path = pathlib.Path(out)
        # 选到"原文件自己"时要问一句（第一轮审计就点了这条，一直没做）：
        # 管线是"读源文件 → 在 work 里生成 → copyfile 到目标"，目标＝源文件时
        # 那一步就是**覆盖掉原稿**，而用户点的是"导出"。备份目录里那份是导入
        # 时的快照，但没人该在导出这一步丢掉原稿。
        try:
            same_file = out_path.resolve() == self.doc_path.resolve()
        except OSError:
            same_file = False
        if same_file and not messagebox.askyesno(
                "要覆盖原文件吗？",
                "你选的导出位置就是原文件本身：\n"
                f"{out_path.name}\n\n"
                "将覆盖原文件，请另存！\n\n"
                "仍然覆盖吗？",
                parent=self):
            self.status("已取消导出（没有覆盖原文件）。")
            return
        cover_type = self._current_cover_type()
        cover = self._current_cover_fields() if self.cv_cover["on"].get() else None
        self.status(f"正在生成论文：{out_path.name}…")
        self.update_idletasks()
        try:
            # 换算结构页的调整结果放在 try 里：失败就走下面的“导出失败”
            # （用户 2026-09-14 要求合并成一条，不再单独报）。
            export_index_ops = self._build_ops_for_export()
            # 导出必须用「页面/格式/结构/封面」几个页签**当场**的表单值，不能
            # 只按 self.applied_preset 这个文件路径重新读一遍——那份文件是
            # 上一次点"应用"时候的快照，用户之后在表单里改的任何东西（哪怕
            # 只是顺手调了下页边距）都会被无声吃掉，导出的和屏幕上看到的对
            # 不上。这是 2026-09-14 定性的"恶性 bug"，`preset_config` 就是
            # 为了堵上它加的——`pipeline.run()` 拿到它就不会再去读文件。
            preset_config = self._collect_current_preset_config()
            from essay_tool.engine.alerts import AlertCollector
            from essay_tool.engine.pipeline import run
            alerts = AlertCollector()
            # 封面范围 / 导出时删不删封面：**这里也要显式传**（审计 2026-09-15
            # 核实）。有结构页调整结果时它自带这两个值（下面用 `inst.` 取）；
            # 没有（没对着这份文档打开过结构页／实例归属不匹配）时，`run()`
            # 只能看这两个参数——不传的话，用户设过的封面范围会静默失效，
            # 封面区重新参与识别与加工。取不到就是 None/False，跟以前完全一致。
            inst = self._adjust_instance
            same_doc = inst is not None and inst.doc_path == self.doc_path
            index, ops = export_index_ops if export_index_ops is not None else (None, None)
            # 长引文段号 / "重新为文档生成"的目录清单：都是结构页手动勾选的
            # 结果，只有对着这份文档做过调整时才有意义——跟 `_build_ops_for_export`
            # 一样读 `confirmed_rows`，不读实时变动的 `inst.rows`。
            quote_paragraphs = ([row.doc_index for row in inst.confirmed_rows
                                if row.special == "引用" and row.doc_index is not None]
                                if same_doc else [])
            regenerate_fields = (frozenset(
                row.block for row in inst.confirmed_rows
                if row.kind == "toc_field" and row.regenerate and row.block)
                if same_doc else frozenset())
            result = run(self.doc_path, None, session_work_dir(),
                        preset_config=preset_config,
                        index=index, ops=ops,
                        alerts=alerts, cover=cover, cover_type=cover_type,
                        cover_end=(inst.confirmed_settings["cover_end"] if same_doc else None),
                        delete_cover=(inst.confirmed_settings["delete_cover"]
                                      if same_doc else False),
                        quote_paragraphs=quote_paragraphs,
                        regenerate_fields=regenerate_fields,
                        renumber_references=(inst.renumber_refs if same_doc else True),
                        heading_scheme=self.var_heading_scheme.get(),
                        session=self._current_session())
            shutil.copyfile(result, out_path)
            # 中间产物用完即删（用户 2026-09-14）；原文件与导出结果都不在临时目录里
            if pathlib.Path(result).resolve() != pathlib.Path(self.doc_path).resolve():
                pathlib.Path(result).unlink(missing_ok=True)
            self._last_dir = str(out_path.parent)
        except Exception as exc:
            self.status(f"生成论文失败：{exc}")
            return
        msg = f"论文已生成：{out_path.name}"
        self.status(msg)
        self._report_alerts(alerts)

    # ------------------------------------------------------------ 流程图联动

    def _blink(self, widget: tk.Misc, times: int = 3) -> None:
        """闪烁提示：按钮闪底色，框子闪边线（同时只保留一个闪烁任务）。

        原色记在控件上（_blink_normal）：重复触发时如果前一次闪到一半被取消，
        直接 cget 会把闪色当成原色，最后就"卡在黄色"了（用户 2026-09-13 抓到）。
        """
        if getattr(widget, "_blink_fill", False):
            self._blink_fill(widget, times)
            return
        job = getattr(widget, "_blink_job", None)
        if job is not None:
            try:
                self.after_cancel(job)
            except Exception:
                pass
        is_button = isinstance(widget, tk.Button)
        key = "bg" if is_button else "highlightbackground"
        normal = getattr(widget, "_blink_normal", None)
        if normal is None:
            normal = widget.cget(key)
            widget._blink_normal = normal
        flash = "#ffe9a3" if is_button else "#e0a800"
        widget.config(**{key: normal})   # 清掉上一次可能残留的闪色

        def step(n: int) -> None:
            # 控件可能已经被销毁（窗口关了、页签重建了）——那时 `config()` 会抛
            # TclError（审计 2026-09-15 核实：这是真的 `after` 回调，排到队上之后
            # 控件已经不存在了）。先查一下，别把异常扔进 Tk 的事件循环。
            if not widget.winfo_exists():
                widget._blink_job = None
                return
            if n <= 0:
                widget.config(**{key: normal})
                widget._blink_job = None
                return
            widget.config(**{key: flash if n % 2 else normal})
            widget._blink_job = self.after(140, step, n - 1)

        step(times * 2)

    _FLASH_FILL = "#ffe08a"

    def _blink_fill(self, widget: tk.Misc, times: int = 3) -> None:
        """闪**整个方块**：`widget` 里所有 Frame/Label 的底色、自造复选框的底色一起
        变成黄色再变回来（用户 2026-09-20："闪烁框不太明显"）。原值第一次闪之前
        记下，重复触发/被打断时不会把闪色当成原色。"""
        job = getattr(widget, "_blink_job", None)
        if job is not None:
            try:
                self.after_cancel(job)
            except Exception:
                pass

        def walk(w):
            yield w
            for c in w.winfo_children():
                yield from walk(c)

        saved = getattr(widget, "_fill_saved", None)
        if saved is None:
            saved = []
            for w in walk(widget):
                if isinstance(w, ttk.Checkbutton):
                    saved.append((w, "style", str(w.cget("style"))))
                elif isinstance(w, (tk.Frame, tk.Label)):
                    saved.append((w, "bg", w.cget("bg")))
                    if isinstance(w, tk.Frame):
                        saved.append((w, "highlightbackground", w.cget("highlightbackground")))
            widget._fill_saved = saved

        def apply(on: bool) -> None:
            for w, key, value in saved:
                if not w.winfo_exists():
                    continue
                if key == "style":
                    if on:
                        m = re.search(r"Size(\d+)", value)
                        size = int(m.group(1)) if m else 12
                        w.configure(style=checkbutton_style_on(self, size, self._FLASH_FILL))
                    else:
                        w.configure(style=value)
                else:
                    w.configure(**{key: self._FLASH_FILL if on else value})

        def step(n: int) -> None:
            if not widget.winfo_exists():
                widget._blink_job = None
                return
            if n <= 0:
                apply(False)
                widget._blink_job = None
                return
            apply(bool(n % 2))
            widget._blink_job = self.after(140, step, n - 1)

        apply(False)
        step(times * 2)

    def _goto_hint(self, tab: int | None, widget) -> None:
        """流程图节点＝导航：跳页签到对应位置，让该点的按钮闪烁。

        **不写状态栏**（用户 2026-09-14：这类导航说明一律不要）——点图上的框
        就是"跳过去＋闪一下"，不再配一句话。

        widget 可以给一个控件，也可以给一串控件（比如「导出」要同时闪两个导
        出按钮）。
        """
        if tab is not None:
            self.nb.select(tab)
        if widget is not None:
            targets = widget if isinstance(widget, (list, tuple)) else [widget]
            for wgt in targets:
                self._blink(wgt)

    def _on_flow(self, nid: str, label: str) -> None:
        """流程图（`flow_nav`）节点 → 跳页签＋闪烁。页签序号（2026-09-20 重排后）：
        0 文件 / 1 结构 / 2 封面 / 3 格式 / 4 页面 / 5 关于。节点 id 见 `flow_nav`
        文件头的接线对照。"""
        if nid == "import":
            self._goto_hint(0, self.btn_import)
        elif nid == "scheme":
            self._goto_hint(0, self.row_scheme)
        elif nid == "excel":
            # 导入后还要点某一行的「应用」，所以导入按钮和各行「应用」一起闪
            btns = getattr(self, "_apply_btns", [])
            self._goto_hint(0, [self.btn_import_preset, *btns])
        elif nid in ("preset", "tpl_preset"):
            # 闪**每一行**的「应　用」，不只第一个（用户 2026-09-14 要求）——
            # 想用哪一份预设就点哪一行的按钮。"（待创建）"那行没有按钮，
            # 不在 _apply_btns 里。
            btns = getattr(self, "_apply_btns", [])
            self._goto_hint(0, btns if btns else self.box_presets)
        elif nid in ("recog", "tpl_struct"):
            self._goto_hint(1, None)
        elif nid == "p_page":
            self._goto_hint(4, None)
        elif nid == "p_fmt":
            self._goto_hint(3, None)
        elif nid == "p_cover":
            self._goto_hint(2, getattr(self, "cover_sw_box", None))
        elif nid == "export":
            self._goto_hint(0, [self.btn_exp_blank, self.btn_exp_doc])
        elif nid == "save":
            self._goto_hint(0, self.btn_save_preset)
        # 其余节点（如果有）不接线：静默返回——用户 2026-09-14 要求这类
        # "未接线节点"诊断也别写状态栏。


def _show_startup_crash(exc: BaseException) -> None:
    """连 `MainWindow()` 都建不起来时的兜底：`report_callback_exception` 挂在
    实例上，这时候还没有实例可挂，只能单独起一个干净的 `tk.Tk()` 来弹提示。
    """
    tb_text = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    code = _make_crash_code(type(exc), tb_text)
    _log_crash(code, type(exc), exc, exc.__traceback__)
    try:
        from essay_tool.config.paths import import_backup_dir
        root = tk.Tk()
        apply_app_icon(root)
        root.withdraw()

        def close() -> None:
            root.destroy()

        _render_crash_popup(
            root, code, import_backup_dir(),
            on_contact=open_developer_contact,
            on_close=close)
        root.mainloop()
    except Exception:
        print(f"程序启动失败，错误代码：{code}", flush=True)


def main() -> None:
    purge_stale_work_dirs()             # 清掉上次崩溃/被强杀遗留的临时目录
    try:
        app = MainWindow()
    except Exception as exc:  # noqa: BLE001 —— 启动阶段的最后一道兜底，必须接住一切
        _show_startup_crash(exc)
        return
    print("主窗口已打开。", flush=True)
    try:
        app.mainloop()
    finally:
        cleanup_session_work_dir()      # 本次的临时工作目录用完即删


if __name__ == "__main__":
    main()
