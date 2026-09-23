r"""AAT 认的"识别层数据类"——**这类引用全部收在这一个文件里**。

背景：识别层正在整块重写（`essay_tool/engine/detect.py` 要下线，换成
`essay_tool/engine/recognize.py`）。AAT 的判据端（`aat_expect.py`）算的是
**期望**、不依赖生产代码；可它得把期望翻译成 `pipeline.run()` 认的那几种
数据类，原先直接 `from essay_tool.engine.detect import ...`——老模块一删，
这些 import 一起断。

做法：**跟着生产代码的接线走**。

1. 先问 `pipeline`：它认的 `ManualStructure` 住在哪个模块（`__module__`）。
   数据类搬了家，`pipeline` 必须跟着 import 新的；我们跟它走，就不会出现
   "AAT 用新的、`pipeline` 用旧的"这种两套类混用。
2. 再从这个模块里把**整组**符号一次取出来（同源，不跨模块拼）。
3. 问不到时（以后 `run()` 不再收 `ManualStructure` 了），按
   `_FALLBACK_MODULES` 逐个试。

模块级就解析（import 本文件即报错），失败时抛 `AatEngineApiError`，把"缺哪个
名字、找过哪些模块、AAT 哪几处受影响"一次说完——不留给别人一条 ImportError
堆栈。生产代码再搬家时，**只改本文件**。
"""

from __future__ import annotations

import importlib
import pathlib
import sys

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent
for entry in (str(ROOT), str(HERE)):
    if entry not in sys.path:
        sys.path.insert(0, entry)

# 整组符号：少一个就换下一个模块，不跨模块拼
GROUP = ("ManualStructure", "HeadingMatch", "TitleMatch",
         "FrontMatterLabel", "CaptionMatch")

# 兜底名单：pipeline 问不到出处时按这个顺序试
_FALLBACK_MODULES = (
    "essay_tool.engine.detect",     # 2026-09-18 之前住这儿
    "essay_tool.engine.recognize",  # 新识别层
    "essay_tool.engine.pipeline",   # 中转站（收 manual_structure 参数就得导出 ManualStructure）
)

_AFFECTED = (
    "aat_expect.to_manual_structure()（旧路径；现役走 expect_index()，函数留着备查）",
    "_probe_regen_section.run_case()（探针）",
)


class AatEngineApiError(RuntimeError):
    """识别层的数据类一个都找不着——AAT 得跟着生产代码改。"""


def _home_module_name() -> str | None:
    """问 pipeline：它认的那批数据类住在哪个模块。问不到返回 None。"""
    try:
        from essay_tool.engine import pipeline
    except Exception:  # noqa: BLE001 —— pipeline 都进不去的话只能靠兜底名单
        return None
    cls = getattr(pipeline, "ManualStructure", None)
    name = getattr(cls, "__module__", None)
    return name or None


def _load_group(module_name: str):
    """从 module_name 里取整组符号。返回 (符号字典, 缺的名字)。"""
    module = importlib.import_module(module_name)
    missing = [name for name in GROUP if getattr(module, name, None) is None]
    if missing:
        return None, missing
    return {name: getattr(module, name) for name in GROUP}, []


def _resolve() -> tuple[str, dict]:
    tried: list[str] = []
    home = _home_module_name()
    order = ([home] if home else []) + [m for m in _FALLBACK_MODULES if m != home]
    for module_name in order:
        try:
            group, missing = _load_group(module_name)
        except Exception as exc:  # noqa: BLE001 —— 逐个尝试，失败的记下来继续
            tried.append(f"{module_name}（进不去：{exc}）")
            continue
        if group is not None:
            return module_name, group
        tried.append(f"{module_name}（缺 {'、'.join(missing)}）")
    raise AatEngineApiError(
        "AAT 找不到识别层的数据类。找过：\n  " + "\n  ".join(tried)
        + "\n\n生产代码把它们搬走/改名时，改本文件的 GROUP 或 _FALLBACK_MODULES。"
        + "\n受影响：\n  " + "\n  ".join(_AFFECTED)
    )


HOME, _GROUP = _resolve()

ManualStructure = _GROUP["ManualStructure"]
HeadingMatch = _GROUP["HeadingMatch"]
TitleMatch = _GROUP["TitleMatch"]
FrontMatterLabel = _GROUP["FrontMatterLabel"]
CaptionMatch = _GROUP["CaptionMatch"]

def recognize(document, scheme: str | None = None):
    """跑一遍生产侧的自动识别，返回“带身份的条目清单”。

    收在这里的理由：返回值的**字段顺序**与元素的**属性名**
    （`match.role`／`match.level`／`caption.kind`／`label.label`／
    `title.title_paragraph_index` …）是 AAT 与生产代码之间最实的一条契约。

    `scheme`：原文标题序号方案（"chinese"/"dotted"）；None 走生产默认。
    """
    from essay_tool.engine.recognize import grammar_for
    from essay_tool.engine.recognize import recognize as _engine_recognize

    grammar = grammar_for(scheme) if scheme else None
    return _engine_recognize(document, grammar=grammar)


def engine_roles(source, scheme: str | None = None) -> dict[int, str]:
    """生产侧认出来的身份：段号 → 身份（措辞与 AAT 判据端一致）。

    这一段拼装原先写在 `gen_matrix.engine_roles()` 里，探针也要用，就收到这里。
    函数内 import `aat_expect`：它是判据端，不欠这个模块什么，但避免将
    来变成循环 import。

    `scheme`：原文标题序号方案，与期望侧用同一个值（写法层两套方案各自对拍）。
    """
    from essay_tool.engine.doc_index import paragraph_ordinals

    document = _as_document(source)
    ordinals = paragraph_ordinals(recognize(document, scheme=scheme))

    roles: dict[int, str] = {}
    for ordinal, entry in ordinals.items():
        role = _role_of(entry)
        if role is not None:
            roles[ordinal] = role

    # 图题/表题：识别层不管（它不是"部件身份"），生产侧仍走旧接口——对面
    # 2026-09-19 明确保留，非部件身份用途只剩它和"附录序号题名"两处。
    headings = {i for i, e in ordinals.items() if e.part in _LEVEL_PARTS}
    for caption in _captions(document, headings):
        role = f"{caption.kind}题"
        if caption.is_english:
            role += "（英文对照）"
        elif getattr(caption, "is_continuation", False):
            role += "（续行）"
        roles[caption.paragraph_index] = role
    return roles


# 部件身份 → AAT 判据端的措辞。两边就这一张对照表，改词只改这里。
_LABEL_PARTS = {
    "摘要", "关键词", "英文摘要", "英文关键词",
    "目录", "插图清单", "附表清单", "插图和附表清单",
}
_BODY_SLOT_PARTS = {"摘要", "关键词", "英文摘要", "英文关键词"}
_ANCHOR_PARTS = {"引言", "结论", "参考文献", "附录", "致谢"}
_LEVEL_PARTS = ("一级标题", "二级标题", "三级标题", "四级标题")


def _role_of(entry) -> str | None:
    """一条索引记录的身份 → AAT 的措辞；没有身份返回 None。"""
    part = entry.part
    if not part:
        return None
    if part == "副标题":
        # 识别层写"副标题"，判据端写"中文副标题"（2026-09-19 对齐时发现）
        return "中文副标题"
    if part in _BODY_SLOT_PARTS:
        # 标签段与它下面紧贴的内容段同名，靠 slot 区分
        slot = entry.meta.get("slot", "label")
        return f"{part}正文" if slot == "content" else f"{part}标签"
    if part in _LABEL_PARTS:
        # 目录与三张清单只有标签，不设 slot
        return f"{part}标签"
    if part in _ANCHOR_PARTS:
        if entry.meta.get("slot") == "content":
            # 表里没有"引言正文"这类角色，判据端按 R18 算正文
            return "正文"
        return part
    return part     # 题目 / 一级~四级标题 / 将来新增的


def _as_document(source):
    """路径就自己读盘；已经是文档就原样用（调用方可能已做过前置处理）。"""
    from docx import Document

    if isinstance(source, (str, pathlib.Path)):
        return Document(str(source))
    return source


def _captions(document, headings):
    """图题/表题——生产侧仍在用旧接口（对面 2026-09-19 明确保留）。

    哪天它也搬走，只改这一个函数。
    """
    from essay_tool.engine.detect import detect_captions

    return detect_captions(document, headings)


# ---------------------------------------------------------------- 期望索引
# 期望对拍那条路的入口：造一份"按期望改过身份"的索引，交给导出。
# 2026-09-19 之前这里是"造一份人工识别结果"（ManualStructure），导出换成
# "清单 + 指令表"之后那条路没了。


def expect_index(prepared, expectation, scheme: str | None = None) -> list:
    """把期望身份写进一份索引。

    `prepared`：跟期望同源的那份"管线看到的文档"（`aat_expect.prepare()`
    的产物）——段号必须与 `expectation` 一致。
    `scheme`：与期望、自动两条路用同一个方案（写回时虽会逐段覆盖身份，
    底稿的编号层级/规范化文字仍按它算）。
    """
    from docx import Document
    from essay_tool.engine.doc_index import paragraph_ordinals
    from essay_tool.engine.recognize import grammar_for, recognize

    grammar = grammar_for(scheme) if scheme else None
    index = recognize(Document(str(prepared)), grammar=grammar)
    for ordinal, entry in paragraph_ordinals(index).items():
        _apply_role(entry, expectation.roles.get(ordinal))
    return index


def _apply_role(entry, role: str | None) -> None:
    """AAT 的期望身份 → 识别层的词表（写回记录的身份与附加信息）。

    图题/表题这类识别层不管的角色**跳过**：它们走的是另一套识别，
    这里动不了，也不该动。

    ⚠️ 两条护栏（都是 2026-09-19 扩夹具时实拍出来的）：

    1. **`role is None`（表没说话）→ 什么都不动**。空段落不归 R18 管
       （"非空"才认正文），判据端对它们不给角色；而识别层会把"摘要/关键词…
       标签下面那个空段"标成 `slot=content`。照清的话 want 侧这些段全掉
       格式（PaperTest/图表脚注测试/写作模板 一大批差异都出自这一条）。
    2. **「正文」只清误认的身份（标题/标签/题目…），不清「锚点部件的正文
       内容」**（`part=引言/结论/参考文献/附录/致谢` 且 `slot=content`）。
       理由：表里根本没有"参考文献正文"这类角色（R18：未识别为其他部件的
       段落都是正文）——判据端对这些段落写的正是「正文」。要是照清，want 侧
       整个参考文献区会掉身份：不套条目样式、不加自动编号、还被当成"这部件
       没内容"补一条占位语（2026-09-19 实测：民总 8 段、扩夹具后 18 份里
       大部分差异都出自这一条）。
    """
    if role is None:
        return
    if role.startswith(("图题", "表题", "多余的正文")):
        return
    if role == "中文副标题":
        entry.part = "副标题"
    elif role == "正文":
        if entry.part in _ANCHOR_PARTS and entry.meta.get("slot") == "content":
            return
        entry.part = None
        entry.meta.pop("slot", None)
    elif role.endswith("标签"):
        entry.part = role[:-2]
        entry.meta["slot"] = "label"
    elif role.endswith("正文") and role[:-2] in _BODY_SLOT_PARTS:
        entry.part = role[:-2]
        entry.meta["slot"] = "content"
    else:
        entry.part = role
        for level, name in enumerate(_LEVEL_PARTS, start=1):
            if role == name:
                entry.meta["level"] = level
        if role in _ANCHOR_PARTS:
            # 覆盖成锚点角色时，槽位要规范成「标题那一行」：引擎原先把这一段
            # 标成**别的部件的内容**时（slot=content），残留的 slot 会让下游
            # 把它当内容处理（2026-09-19 图表脚注测试实拍：期望把裸「附录」
            # 认成锚点，但 slot 残留 content，附录区探测跟着受影响）。
            entry.meta["slot"] = "label"


__all__ = ["HOME", "AatEngineApiError", "recognize", "engine_roles",
           "ManualStructure", "HeadingMatch", "TitleMatch",
           "FrontMatterLabel", "CaptionMatch"]
