"""告警：照常出成品，但把"看着可疑"的地方标出来让用户自己看一眼。

跟 `ConfigError` 那类"干不下去了"的错误不是一回事——**告警不阻断流程**。

现在有七条规则（写新规则前先看看有没有能复用的）：
1. `check_duplicate_parts`——同一个部件被识别出多次。规范 [226] 其实明确允许
   有多个附录（"有多个附录的，序列靠后的附录均应当另起一页编排"），所以这
   不算错；但题目/摘要/关键词/引言/结论/参考文献/致谢这些本来就应该只有
   一份，多了通常意味着识别认错了、或者源文件里有重复，值得提醒一句。
2. `check_demoted_titles`——"看着像标题但被判回正文"的段落（宽进识别的
   保险丝，防止标题漏判后无声无息）。
3. `check_sequence_issues`——标题编号顺序异常（1、3、5 跳号之类）。
4. `check_unaccepted_revisions`——文档里还留着未接受的修订标记。
5. `check_list_parts`——三张清单部件（插图清单等）的插入计划：拿到了标签/真域
   就不管，没有就插；启用但缺对应题注时报一条"不存在的清单对象"（原来的
   "清单已存在""只生成一半""形态冲突"三条已按用户 2026-09-15 的决定删掉）。
6. `check_front_matter_body`——目录/三张清单的标签下方写了正文（2026-09-15
   用户要求：这一页只该装目录域生成的条目，手打正文要报出来）。
7. `check_unnumbered_objects`——有图片/表格却没识别到题注（2026-09-15）。
   "有题注"**只算挂图题/表题样式的那几类**：用户当日裁决"只有续表？报未编号"
   ——续行与英文对照行不算。
8. `check_duplicate_front_labels`——摘要/关键词/Abstract/Keywords 这四个
   前置字面标签里，某一个在前置区出现了不止一次（2026-09-19）。识别层
   （`recognize._scan_forward`）只认第一次出现的位置划内容范围，第二次
   开始"认得出来但不采信"，不会被标 `.part`，所以 `check_duplicate_parts`
   看不到它——这条规则专门读 `Entry.meta["duplicate_label_positions"]`
   把这批"识别到但没采信"的位置报出来。

（2026-09-16 删掉的第 6 条：`check_caption_sequence`“题注编号不连续”。用户
当日定“我们只管样式、编号与我们无关”——不写域、不改号、不重排、不告警，
所以这条连同分图字母连续性一起下线。）

**接口是"先收集、后统一输出"，不是边走边 print**：将来接 GUI 时，告警要显示在
一个面板里，所以收集阶段必须跟"怎么显示"解耦。`Alert` 只描述"哪里、什么事"，
呈现交给调用方（命令行调 `print_report()`，GUI 自己遍历 `items` 去上色/分组）。

加新规则的姿势：写一个 `check_xxx(self, ...)` 方法，里面调 `self.add(...)`，
然后在 pipeline 里调一次。**别在 check_xxx 里 print**——那就把两件事又绑回去了。
"""

from __future__ import annotations

from dataclasses import dataclass

from docx.oxml.ns import qn

from essay_tool.engine.doc_index import (
    LIST_PART_KINDS,
    has_toc_field,
    is_toc_paragraph,
    paragraph_ordinals,
    toc_field_span_indices,
)

# "这一页只该装目录/清单条目"的四个部件——`recognize._GAP_LABEL_WORDS` 也是
# 同一份名单，两处各自维护是因为 alerts 不该 import recognize（层次相反：
# alerts 服务于 pipeline，pipeline 才 import recognize）。
_TOC_PAGE_PARTS = frozenset({"目录", "插图清单", "附表清单", "插图和附表清单"})

# 需要检查"是否重复"的部件——前置标签 + 锚点标签，都是"理论上只该出现
# 一次"的角色（附录允许在成品里正常出现多次，但重复了也提醒一句，跟
# 旧 detect.py 的口径一致）。
_DUPLICATE_CHECK_PARTS = frozenset({
    "摘要", "关键词", "英文摘要", "英文关键词",
    "目录", "插图清单", "附表清单", "插图和附表清单",
    "引言", "结论", "参考文献", "附录", "致谢",
})


def _last_component(number) -> int:
    """点分编号（`(1, 2, 1)`）只比最后一段——层级已经由 `level` 表达，同一
    层级内"是不是接着上一个数"就是最后一段的事；中文体系的整数原样返回。"""
    return number[-1] if isinstance(number, tuple) else number


def check_sequence(index) -> list[tuple[int, str, str]]:
    """检查标题编号的三种异常：未归零、跳号、重号、孤立。跳级本身不算问题。

    2026-09-18 从 `detect.py` 搬到这里、改成读识别索引（`Entry.meta`）——
    判定逻辑本身一个字没变，只是数据来源从 `list[HeadingMatch|None]` 换成
    `list[doc_index.Entry]`。返回 `(段落下标, 错误名称, 错误概要)` 列表，
    名称/概要口径见 `docs/告警文案编号表.md`。
    """
    ordinals = paragraph_ordinals(index)
    # 引言/结论只有带编号（点分体系的 `1 引言`）时才有 `meta["number"]`，那种
    # 算一级标题那串编号里的一员；不带编号的锚点没有 number，自然被滤掉。
    numbered_parts = ("一级标题", "二级标题", "三级标题", "四级标题", "引言", "结论")
    flat = [
        (i, entry.meta["level"], _last_component(entry.meta["number"]))
        for i, entry in sorted(ordinals.items())
        if entry.part in numbered_parts and entry.meta.get("number") is not None
    ]
    warnings: list[tuple[int, str, str]] = []
    last_number: dict[int, int | None] = {}
    expect_reset: dict[int, bool] = {}

    for idx, (paragraph_index, level, number) in enumerate(flat):
        for deeper in [lv for lv in last_number if lv > level]:
            del last_number[deeper]
        for deeper in [lv for lv in expect_reset if lv > level]:
            del expect_reset[deeper]

        if expect_reset.get(level) and number != 1:
            warnings.append((paragraph_index, "未归零", "未从头开始的标题序号"))
        expect_reset[level] = False
        expect_reset[level + 1] = True

        prev = last_number.get(level)
        if prev is not None and number != prev + 1:
            name = "重号" if number == prev else "跳号"
            warnings.append((paragraph_index, name, "未连续使用的标题序号"))
        last_number[level] = number

        if (not _has_sibling(flat, idx, level, -1)
                and not _has_sibling(flat, idx, level, 1)):
            warnings.append((paragraph_index, "孤立", "单项使用的标题序号"))

    return warnings


def _has_sibling(flat, idx: int, level: int, direction: int) -> bool:
    i = idx + direction
    while 0 <= i < len(flat):
        other_level = flat[i][1]
        if other_level < level:
            return False
        if other_level == level:
            return True
        i += direction
    return False


def _kind_desc(kinds) -> str:
    """题注种类集合 -> 给人看的说法（"图题" / "表题" / "图题或表题"）。"""
    names = ["图题" if kind == "图" else "表题" for kind in ("图", "表") if kind in kinds]
    return "或".join(names)


def _snippet(text: str, limit: int = 12) -> str:
    """把段落开头截一小段出来当定位线索——光报段落下标，用户对不上号。"""
    text = " ".join(text.split())
    return text[:limit] + ("…" if len(text) > limit else "")


@dataclass(frozen=True)
class Alert:
    """一条告警。

    `kind` 是机器可读的分类（如 `"duplicate_part"`），方便调用方按类型过滤或
    配色；`name` 是给人看的**错误名称**；`part` 是涉及的部件名（可能已经写进
    `message` 里，别的地方要按部件筛选时用它）；`message` 是**错误概要**——
    **不带位置、不带句末句号**，位置统一由 `paragraph_indices` 拼（见 `line`）。

    一整行的格式是用户 2026-09-15 定的：

        `错误名称：错误概要，在段落（段号）。`

    没有对应段落的（比如"没识别到题注，所以没生成清单"）末尾就不写
    "在段落（…）"那一截。

    段落下标就是**结构页"段号"列里显示的那个数字**（`adjust_preview` 直接拿
    `row.doc_index` 当段号画，没有加一）——用户 2026-09-14 要求告警必须能
    对上屏幕上看到的段号，所以两边用的是同一套数字。
    """

    kind: str
    name: str
    part: str
    message: str
    paragraph_indices: tuple[int, ...] = ()

    @property
    def position(self) -> str:
        """段号部分（只有数字）：`34、35、36`；没有对应段落时返回空串。

        **全部列出、不截断**（2026-09-15 审计核实后改）：原来超过 4 个就用
        `…` 收尾，而 `consolidate()` 会把多条同类告警合成一条——合成之后
        段号本来就可能很多，截掉了用户就没地方看后面那几个段号（弹窗里那
        一份是完整列表，可一旦关掉就无从查起）。宁可写长，也不丢位置。
        """
        if not self.paragraph_indices:
            return ""
        return "、".join(str(i) for i in self.paragraph_indices)

    @property
    def line(self) -> str:
        """一整行：`错误名称：错误概要，在段落（段号）。`

        （GUI 日志、弹窗、命令行报告共用；没有段号就写 `错误名称：错误概要。`）
        """
        tail = f"，在段落（{self.position}）。" if self.paragraph_indices else "。"
        return f"{self.name}：{self.message}{tail}"

    def __str__(self) -> str:
        return self.line


class AlertCollector:
    """收集告警。用法见模块开头的说明。"""

    def __init__(self) -> None:
        self._items: list[Alert] = []

    def add(
        self,
        kind: str,
        name: str,
        part: str,
        message: str,
        paragraph_indices: tuple[int, ...] = (),
    ) -> None:
        self._items.append(Alert(kind, name, part, message, tuple(paragraph_indices)))

    # -------------------------------------------------- 规则：部件重复

    def check_duplicate_parts(self, index) -> None:
        """同一个部件被识别出多次就告警。

        2026-09-18 改成读识别索引：统计的是每个"标签位"（`Entry.meta["slot"]
        =="label"`，或者压根没有 slot 概念的目录/三张清单）出现了几次——
        只数标签本身，不把同一个部件的多段内容也算进去。

        附录也算重复（用户 2026-09-15："附录标签不允许多次（如果有多个附录，
        应该只有一个附录，然后是附录A、附录B，而不是两个附录标签）"）。
        """
        ordinals = paragraph_ordinals(index)
        seen: dict[str, list[int]] = {}
        for i, entry in ordinals.items():
            if entry.part in _DUPLICATE_CHECK_PARTS and entry.meta.get("slot", "label") == "label":
                seen.setdefault(entry.part, []).append(i)

        for part, indices in seen.items():
            if len(indices) < 2:
                continue
            self.add("duplicate_part", "重复的部件", part, f"“{part}”", tuple(indices))

    # -------------------------------------------------- 规则：前置标签重复（识别到但没采信的第二次出现）

    def check_duplicate_front_labels(self, index) -> None:
        """摘要/关键词/Abstract/Keywords 这四个"纯字面前置标签"，某一个在
        前置区出现了不止一次（2026-09-19 用户拍板新加）。

        `recognize._scan_forward()` 里这四个标签各自独立扫描，只认**第一次**
        出现的位置划"标签+内容"范围；同一个标签又出现一次，识别到但不采信
        第二次的位置（不然像"附录允许重复"那样把内容归属算乱）——这批"没
        采信"的位置**不会**被写进 `Entry.part`，所以 `check_duplicate_parts`
        （只看 `.part`）扫不到它们，得单独读 `recognize()` 留在 canonical
        那条 Entry 上的 `meta["duplicate_label_positions"]`。

        跟 `check_duplicate_parts` 用同一套 `kind`/`name`/`message` 措辞
        （"重复的部件"），呈现上没有区别，用户不需要知道背后是两条不同的
        规则在管。
        """
        ordinals = paragraph_ordinals(index)
        for i, entry in ordinals.items():
            extra = entry.meta.get("duplicate_label_positions")
            if not extra:
                continue
            self.add("duplicate_part", "重复的部件", entry.part,
                      f"“{entry.part}”", (i, *extra))

    # -------------------------------------------------- 规则：标题被降级

    def check_demoted_titles(self, index) -> None:
        """"疑似标题"：**只因为超过长度上限（60 字）**才没被认成标题的段落。

        2026-09-20 用户裁决：旧版"带圈数字/第某章/罗马数字……"那一大类疑似标题
        告警不要；只留"太长"这一种——编号写法合规、不以句末标点结尾，却因为
        去掉空格后超过 60 字而归了正文，同时提醒一句（识别层给这类段落打
        `meta["overlong_title"]`，见 `recognize._scan_backward`）。

        这些段落按**正文**处理，拿不到标题格式；静默丢掉才是最坏的结果。
        用户在结构页里把它拖成标题后（身份已是一至四级标题）不再报。
        """
        ordinals = paragraph_ordinals(index)
        heading_parts = ("一级标题", "二级标题", "三级标题", "四级标题")
        for i, entry in sorted(ordinals.items()):
            if not entry.meta.get("overlong_title") or entry.part in heading_parts:
                continue
            self.add(
                "demoted_title",
                "疑似标题",
                "标题",
                f"请检查“{_snippet(entry.text.strip())}”是否为标题",
                (i,),
            )

    # -------------------------------------------------- 规则：序号异常

    def check_sequence_issues(self, index) -> None:
        """标题编号掉号 / 跨级没归零 / 孤立列举（判定逻辑见本模块的 `check_sequence()`）。"""
        for paragraph_index, name, message in check_sequence(index):
            self.add("sequence", name, "标题编号", message, (paragraph_index,))

    # -------------------------------------------------- 规则：未接受的修订

    def check_unaccepted_revisions(self, document, cover_end: int | None = None) -> None:
        """文档里还有没接受/拒绝的修订（Word 的"修订"功能）。

        `w:ins`（插入）/`w:del`（删除）里的文字**不在 `paragraph.runs` 里**——
        标题、标签、图题识别的每一步都看不见它们，所以带修订的稿子会大量漏判，
        还可能"以为部件缺失、又补一套"，弄出两套摘要来。
        **内容本身不会丢**（实测 `w:ins`/`w:del` 原样保留），只是不参与识别。
        """
        body = document._element.body
        inserted = len(body.findall(f".//{qn('w:ins')}"))
        deleted = len(body.findall(f".//{qn('w:del')}"))
        total = inserted + deleted
        if not total:
            return
        # 首个带修订的段落下标——告警要能对上结构页的段号列。
        # 封面区间（`cover_end`）里的修订不管：那是用户自己的地盘，工具既不
        # 识别也不加工（用户 2026-09-15："一旦已经被识别为封面，那就不应该管，
        # 也不算进清单里，原则上也不应该告警"）。整篇只有封面里带修订时，
        # 这条告警就不报。
        def _hits(paragraph) -> bool:
            return bool(paragraph._p.findall(f".//{qn('w:ins')}")
                        or paragraph._p.findall(f".//{qn('w:del')}"))

        first = next(
            (i for i, p in enumerate(document.paragraphs)
             if (cover_end is None or i > cover_end) and _hits(p)),
            None,
        )
        if first is None:
            # 顶层段落里一处都没有：要么全在封面区间里（按上面的裁决不报），
            # 要么全在表格单元格里（`document.paragraphs` 不含单元格段落）。
            # 后者以前是静默漏报（审计 2026-09-16）：计数扫的是全篇 XML，定位
            # 只扫顶层段落，两者一对不上就 return。改后照报，只是给不出段号——
            # 段号列本来也没有对应的行。
            in_cover = 0
            if cover_end is not None:
                in_cover = sum(
                    len(p._p.findall(f".//{qn('w:ins')}"))
                    + len(p._p.findall(f".//{qn('w:del')}"))
                    for i, p in enumerate(document.paragraphs) if i <= cover_end)
            if total == in_cover:
                return
            self.add("revision_marks", "未接受的修订", "修订", "请先接受修订")
            return
        self.add(
            "revision_marks",
            "未接受的修订",
            "修订",
            "请先接受修订",
            (first,),
        )

    # -------------------------------------------------- 规则：目录/清单页正文

    def check_front_matter_body(self, document, index) -> None:
        """目录／三张清单的标签下面出现了正文。

        用户 2026-09-15："只能认真正的 toc 域，不能认用户手动输入的标签和
        正文。目录列本来就不应该有正文。如果用户只写了标签，把标签拿过来
        用是可以的；如果目录有用户正文，报警。"

        所以判据是：标签下面（直到下一个前置标签或下一个标题）出现的、
        既不是目录/清单域、也不是域生成的条目（`toc 1` 这类样式）的文字，
        就算误放在这一页的正文。压力！.docx 实拍："目录"两个字下面写了
        "不会真把这个当成目录了吧？"，成品里它就粘在目录域后面。

        **不做任何移动或删除**：用户 2026-09-15 定的是"跟在上一个标签的
        正文之后，让用户自己看着办"——只报，不动文档。

        **一路看到下一个部件为止，段号全部列出**（用户 2026-09-16：原来那两个
        "往下只看 60 段、每个标签最多报 6 个段号"的封顶删掉，按表来）。

        ⚠️ **域渲染出来的条目还算了"域区间"这一路**（2026-09-16）：Word 的
        "自动目录"内容控件里，条目段可能一个 `toc N` 样式都不挂（`is_toc_paragraph`
        认不出来）——那样整页旧目录都会被当成"多余的正文"报出来（民总.docx
        实测报了 18 条）。判据跟 `cleanup.drop_regenerated_fields` 用同一个
        `doc_index.toc_field_span_indices()`。
        """
        paragraphs = document.paragraphs
        ordinals = paragraph_ordinals(index)
        label_at = {i: e.part for i, e in ordinals.items() if e.part in _TOC_PAGE_PARTS}
        claimed = {i for i, e in ordinals.items() if e.part}
        in_field = toc_field_span_indices(paragraphs)
        for label_index in sorted(label_at):
            label = label_at[label_index]
            stray: list[int] = []
            for j in range(label_index + 1, len(paragraphs)):
                if j in claimed:
                    break      # 到了下一个部件/正文标题，这一页到此为止
                paragraph = paragraphs[j]
                if not paragraph.text.strip():
                    continue
                if j in in_field or has_toc_field(paragraph) \
                        or is_toc_paragraph(paragraph):
                    continue   # 目录/清单域本体，以及域渲染出来的条目：本来就是它的内容
                stray.append(j)
            if stray:
                self.add(
                    "toc_body",
                    "多余的正文",
                    label,
                    f"请检查{label}下疑似多余的正文",
                    tuple(stray),
                )

    # -------------------------------------------------- 规则：清单部件

    def check_list_parts(
        self,
        structure_required: dict[str, bool],
        caption_kinds: set[str],
        existing_field_kinds: set[str] = frozenset(),
        regenerate_fields: set[str] = frozenset(),
    ) -> dict[str, set[str]]:
        r"""插图清单/附表清单/插图和附表清单（规范 [91]-[93]）的决策 + 告警。

        返回值有点特别：{部件名: 要插的题注种类}，只有"启用 + 源文件没有 +
        有题注可收"的部件才会在字典里。**这个返回值就是插入清单的依据**
        ——本方法里每个判断都配着一句话要说，把"判断"和"告警"放在同一个
        地方，两边就不会对不上（apply_missing_parts 拿到什么就插什么，
        不再自己做判断）。

        `caption_kinds`：**会被清单域收进条目**的题注种类（中文"图"/"表"；
        英文对照行不挂样式、不进清单，调用方已过滤）。

        `existing_field_kinds`：源文件里**真正存在**的清单域收了哪几种题注
        ——**两种写法都认**（`detect.list_kinds_in_instr()` 是唯一真源）：
        ① `TOC \c "图"` 式，取引号里的字；② 我们自己写的
        `TOC \h \z \t "图题,1"` 式，按样式名映射成“图”/“表”。
        用户 2026-09-15：已有就"不动"——四个位置的判据统一，不再看用户手打的字。

        四种情况：
        * 启用 + 没有真域 + 有题注：插入（进字典），不告警（标签在不在都一样要插
          ——标签只当插入位置用）；
        * 启用 + 没有真域 + 没题注：不插——空域刷新后会显示"没有指定样式的
          文字"报错字样，很难看；告警说明可能的原因；
        * 启用 + 已有真域：**静默不动**——既不定插也不报警（原第 13 条"清单已
          存在"2026-09-15 已取消：用户——"按我们的逻辑修完就没有适用
          情况了"）；
        * 没启用：什么都不做。源文件里原有的清单也不碰。

        ⚠ “有没有真域”只看**域**，不看标签（用户 2026-09-16：“按文档来”）：
        源文件里只写了【插图清单】四个字、没有域时，旧写法把它当成"已有"直接
        静默 continue，成品里那张清单**标签在、下面空着，也不告警**（实测：
        一个清单域都不插）。现在标签只当插入位置（跟目录一个口径：用户只写了
        标签，把标签拿过来用）。
        """
        plan: dict[str, set[str]] = {}
        for part, needed in LIST_PART_KINDS.items():
            if not structure_required.get(part, False):
                continue
            # 用户勾了"重新为文档生成"的部件（`regenerate_fields`）：标签还在
            # 是**正常情形**（用户要的就是"标签留着、域重做"），不能因此跳过
            # 不插——以前就是卡在这儿：勾了之后旧域被删、新域没插，成品变成
            # "有标题、没条目"（2026-09-15 实测）。
            regenerating = part in regenerate_fields
            if not regenerating and set(needed) <= existing_field_kinds:
                # 源文件里已经有这张清单的真域了 → 什么都不做，也**不再
                # 报警**：用户 2026-09-15——"按照我们的逻辑修完就没有适用
                # 情况了"（已有就不动、没有就生成，本来没有需要用户决定的事）。
                continue
            available = set(needed) & caption_kinds
            missing = set(needed) - caption_kinds
            if missing:
                # 14（一个题注都没识别到）与 15（只识别到一半）合并成这一条，
                # 概要里点名缺的是哪种（用户 2026-09-15 定的文案）。
                self.add(
                    "list_missing",
                    "不存在的清单对象",
                    part,
                    f"启用了{part}，但不存在{_kind_desc(missing)}",
                )
            if available:
                plan[part] = available

        # 16（插图和附表清单与独立清单同时启用）**不再报警**（用户 2026-09-15）：
        # 改成在结构界面直接不让选中冲突的组合，见 adjust_preview._do_toggle。
        # 预设文件里真写成了冲突组合的话，这里按"合并形态优先"处理。
        if "插图和附表清单" in plan:
            plan.pop("插图清单", None)
            plan.pop("附表清单", None)
        return plan

    # -------------------------------------------------- 规则：未编号的图表

    def check_unnumbered_objects(self, document, captions,
                                 cover_end: int | None = None) -> None:
        """论文里有图片/表格，却没识别到题注编号 → 报出来（用户 2026-09-15）。

        判据很保守，只看"有没有题注"这一件事：
        * **图片**：段落里有嵌入图（`w:drawing` / `w:pict`），而这一段、它
          的上一段、下一段都不是识别到的题注；
        * **表格**：表格是块级元素、不在 `document.paragraphs` 里，所以按 body
          子元素顺序走一遍，看表格**紧邻的上一段/下一段**是不是题注（规范把
          表题放在表上方、图题放在图下方，两头都算）。

        `cover_end`：用户手敲的封面范围（第 0 段到第 N 段含）。**封面区间里
        的图和表不参与这条规则**——用户 2026-09-15 的裁决："一旦已经被识别为
        封面，那就不应该管，也不算进清单里，原则上也不应该告警。"（这条规则
        是直接扫全篇的，拿不到识别阶段已经滤过封面那一份结果，所以要单独传
        进来。）

        位置：图片行/表格邻近段落的段号（就是结构页里那个数字），能给就给，
        给不了就不写（`Alert.line` 会自动省掉那一截）。

        ⚠️ **"有题注"只算挂了图题/表题样式的那一类**（2026-09-16 起"分图/
        分表"这个概念已删，凡挂上样式的都算正题注）。
        续行（`表 1（续）`）与英文对照行不挂样式 ⇒ **不算这张表有题注**——
        用户 2026-09-15 拍板："只有续表？报未编号。"（只有续表就说明主表
        那张表的题注没在这里，用户确实该去看一眼）。
        """
        caption_at = {c.paragraph_index for c in captions
                      if c.wears_caption_style}
        image_stray: list[int] = []
        for i, paragraph in enumerate(document.paragraphs):
            if cover_end is not None and i <= cover_end:
                continue      # 封面区间：不参与（用户 2026-09-15）
            element = paragraph._p
            if not (element.findall(f".//{qn('w:drawing')}")
                    or element.findall(f".//{qn('w:pict')}")):
                continue
            if {i - 1, i, i + 1} & caption_at:
                continue
            image_stray.append(i)
        if image_stray:
            self.add(
                "unnumbered_image",
                "未编号的图片/表格",
                "图片",
                "论文中存在图片，但未识别到题注，请检查",
                tuple(image_stray),
            )

        table_stray: list[int] = []
        index = -1
        previous: int | None = None
        pending: list[int | None] = []
        for child in document._element.body.iterchildren():
            if child.tag == qn("w:p"):
                index += 1
                previous = index
            elif child.tag == qn("w:tbl"):
                pending.append(previous)
        # 表格后面那一段的段号：拿表格之后第一个段落的序号（没有就算了）
        index = -1
        following: list[int | None] = []
        pending_after: list[bool] = []
        for child in document._element.body.iterchildren():
            if child.tag == qn("w:tbl"):
                pending_after.append(True)
            elif child.tag == qn("w:p"):
                index += 1
                if pending_after:
                    following.extend([index] * len(pending_after))
                    pending_after.clear()
        following.extend([None] * len(pending_after))
        for before, after in zip(pending, following):
            if cover_end is not None and (before is None or before <= cover_end):
                continue      # 表格落在封面区间里（它前面那段也是封面的）：不管
            touched = {before, after}
            if touched & caption_at:
                continue
            table_stray.append(before if before is not None else -1)
        if table_stray:
            self.add(
                "unnumbered_table",
                "未编号的图片/表格",
                "表格",
                "论文中存在表格，但未识别到题注，请检查",
                tuple(i for i in table_stray if i >= 0),
            )

    # -------------------------------------------------- 取值 / 输出

    def consolidate(self) -> None:
        """把**同一个问题**的多条告警并成一条，段号并列。

        用户 2026-09-15 的裁决："反正所有问题合着只报一次，段号并列，就不用
        管了。"合并的键是 **(机器分类 + 错误名称 + 错误概要)**——三者完全一样
        才算同一个问题：`孤立` 报 6 次就并成 1 次（段号 38、40、43…）；而
        `疑似标题` 报 3 次**不并**（三段文字不同，概要里写的是各自的片段，并了
        就看不出来是哪一段了）。段号去重、升序；`Alert.position` 本来只列前 4
        个、多的用 `…` 收尾。

        调用方是 `pipeline.collect_source_alerts()`——那是 `run()`（导出）与
        `review()`（打开结构页）唯一的汇总口子，所以两边的条数天生一致。
        """
        if len(self._items) < 2:
            return
        merged: dict[tuple[str, str, str], Alert] = {}
        order: list[tuple[str, str, str]] = []
        for item in self._items:
            key = (item.kind, item.name, item.message)
            if key in merged:
                previous = merged[key]
                merged[key] = Alert(
                    previous.kind, previous.name, previous.part, previous.message,
                    previous.paragraph_indices + item.paragraph_indices,
                )
            else:
                merged[key] = item
                order.append(key)
        self._items = [
            Alert(a.kind, a.name, a.part, a.message,
                  tuple(sorted(set(a.paragraph_indices))))
            for a in (merged[key] for key in order)
        ]

    @property
    def items(self) -> list[Alert]:
        """按添加顺序返回全部告警——GUI 用这个自己渲染。"""
        return list(self._items)

    def __len__(self) -> int:
        return len(self._items)

    def __bool__(self) -> bool:
        return bool(self._items)

    def report(self) -> str:
        """拼成给人看的多行文本（不带结尾换行）。没有告警时返回空串。

        每行都带位置（结构页"段号"列里的数字）＋部件名，见 `Alert.line`。
        """
        if not self._items:
            return ""
        lines = [f"！有 {len(self._items)} 处需要确认（不影响出成品）："]
        lines.extend(f"  · {alert.line}" for alert in self._items)
        return "\n".join(lines)

    def print_report(self) -> None:
        text = self.report()
        if text:
            print(f"\n{text}")
