"""套用页面设置（纸张、方向、页边距、装订线、页眉页脚距边）。

以前还有一项"打印方式"（单面/双面），2026-09-14 已从预设里整体删除：
它是打印机/打印任务的设置，不是文档属性，docx 里没有能可靠生效的对应
属性，留在预设里只会让人以为它能用。
"""

from __future__ import annotations

from docx.document import Document as DocumentObject
from docx.enum.section import WD_ORIENT

from essay_tool.config.schema import PageSetup

ORIENTATIONS = {
    "纵向": WD_ORIENT.PORTRAIT,
    "横向": WD_ORIENT.LANDSCAPE,
}


def apply_page_setup(document: DocumentObject, page: PageSetup) -> None:
    if page.orientation not in ORIENTATIONS:
        raise ValueError(f'不认识的纸张方向："{page.orientation}"（应为"纵向"或"横向"）')

    orientation = ORIENTATIONS[page.orientation]
    width, height = page.page_width, page.page_height
    if orientation == WD_ORIENT.LANDSCAPE:
        width, height = height, width

    for section in document.sections:
        section.orientation = orientation
        section.page_width = width
        section.page_height = height
        section.top_margin = page.margin_top
        section.bottom_margin = page.margin_bottom
        section.left_margin = page.margin_left
        section.right_margin = page.margin_right
        section.gutter = page.gutter
        section.header_distance = page.header_distance
        section.footer_distance = page.footer_distance
