"""文档结构清单：汇总"该有的部件，哪些找到了、哪些没找到"。

只负责"知道哪里是哪里"——纯描述性的清点结果，不决定缺了要不要补、
补什么内容。这是有意的分工：识别和处理分两步做，识别先做完、做扎实，
处理（自动补充缺失部件、套格式等）由 apply_structure.py 负责——那个
"下一阶段"已经做完了。**主管线目前不经过本模块**：pipeline 用的是
found_index + apply_missing_parts 这条路，本模块只服务于诊断工具
tools/check_structure.py。

部件的必需/可选性现在是预设（"结构部件"小节）里的用户设置，
不再由本模块的规范解读决定；本模块只负责按名字去查这份设置。
"""

from __future__ import annotations

from dataclasses import dataclass, field

from docx.document import Document

from essay_tool.engine.detect import (
    detect_document_title,
    detect_english_subtitle,
    detect_english_title,
    detect_front_matter_labels,
    detect_headings,
)

_ANCHOR_PARTS = ["引言", "结论", "参考文献", "附录", "致谢"]
# 前置部件。以前只有摘要/关键词/英文摘要/目录四项，清单里少的人家那几个
# （英文题目、英文副标题、英文关键词、三张清单）就永远“找不到”——诊断工具
# 报出来的“缺件”是假的（审计 2026-09-15 核实）。现在跟 `gui.adjust_preview.
# BLOCK_NAME` 那份部件名单对齐，一个不漏。
_FRONT_MATTER_PARTS = [
    "摘要", "关键词", "英文题目", "英文副标题", "英文摘要", "英文关键词",
    "目录", "插图清单", "附表清单", "插图和附表清单",
]
# `detect_front_matter_labels` 报出来的标签 → 部件名。英文那两个标签在探测器
# 里叫 Abstract/Keywords（那是文档里的字面），部件名叫“英文摘要/英文关键词”。
_LABEL_TO_PART = {
    "Abstract": "英文摘要",
    "Keywords": "英文关键词",
}


@dataclass
class StructurePart:
    name: str
    required: bool
    found: bool
    paragraph_index: int | None = None


@dataclass
class StructureInventory:
    title: StructurePart
    subtitle: StructurePart  # 可选，found 且 paragraph_index 有值才算存在
    parts: list[StructurePart] = field(default_factory=list)

    @property
    def missing_required(self) -> list[StructurePart]:
        return [p for p in [self.title, *self.parts] if p.required and not p.found]

    @property
    def missing_optional(self) -> list[StructurePart]:
        return [p for p in [self.subtitle, *self.parts] if not p.required and not p.found]


def build_structure_inventory(
    document: Document, structure_required: dict[str, bool]
) -> StructureInventory:
    matches, _demoted = detect_headings(document)
    front_matter = detect_front_matter_labels(document)
    title = detect_document_title(document, matches, front_matter)

    front_matter_by_label = {f.label: f.paragraph_index for f in front_matter}
    anchor_by_role = {m.role: m.paragraph_index for m in matches if m and m.role}
    front_matter_by_part: dict[str, int] = {}
    for label, index in front_matter_by_label.items():
        front_matter_by_part.setdefault(_LABEL_TO_PART.get(label, label), index)
    # 英文题目/英文副标题不在 `detect_front_matter_labels` 里（它们没标签，得
    # 看位置和“像不像英文题名”，见 detect_english_title/detect_english_subtitle），
    # 单独查一道补进来。
    english_title = detect_english_title(
        document, front_matter,
        known={title.title_paragraph_index, title.subtitle_paragraph_index} if title else set())
    if english_title is not None:
        front_matter_by_part["英文题目"] = english_title
        # `recognized` 跟主管线（`pipeline.recognize_source`）传的一样：已认出的
        # 标题/前置标签不能再被当成英文副标题，否则两边会差一项（探针
        # `tools/_check_inventory_vs_pipeline.py` 就是拿这个对账的）。
        recognized = {i for i, m in enumerate(matches) if m is not None}
        recognized |= {f.paragraph_index for f in front_matter}
        english_subtitle = detect_english_subtitle(document, english_title, recognized)
        if english_subtitle is not None:
            front_matter_by_part["英文副标题"] = english_subtitle

    title_part = StructurePart(
        "题目", structure_required.get("题目", False), found=title is not None,
        paragraph_index=title.title_paragraph_index if title else None,
    )
    subtitle_part = StructurePart(
        "副标题", structure_required.get("副标题", False),
        found=bool(title and title.subtitle_paragraph_index is not None),
        paragraph_index=title.subtitle_paragraph_index if title else None,
    )

    parts: list[StructurePart] = []
    for name in _FRONT_MATTER_PARTS:
        idx = front_matter_by_part.get(name)
        parts.append(
            StructurePart(name, structure_required.get(name, False), found=idx is not None, paragraph_index=idx)
        )
    for role in _ANCHOR_PARTS:
        idx = anchor_by_role.get(role)
        parts.append(
            StructurePart(role, structure_required.get(role, False), found=idx is not None, paragraph_index=idx)
        )

    return StructureInventory(title=title_part, subtitle=subtitle_part, parts=parts)
