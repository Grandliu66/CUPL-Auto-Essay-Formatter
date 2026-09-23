"""把三张真封面烘成 PNG —— GUI「封面」页签下面的样张（论文 / 课程 / 学年论文）。

链路：engine/cover.py 生成样张 docx（留空、不放占位）→ 本机 Word 导出 PDF
→ PyMuPDF 把第 1 页转 PNG → essay_tool/gui/assets/。

为什么不直接在 GUI 里"画一张示意图"：封面排版有 30 多处像素级调整（逐段
钉死行距、悬挂缩进、下划线宽度、标题换行补偿……），画出来的图迟早和真货
不一致；烘真图一劳永逸，而且 GUI 运行时只读 PNG、不依赖 Word。

用法：.venv\\Scripts\\python.exe tools\\make_cover_preview.py
产出：essay_tool\\gui\\assets\\cover_thesis.png / cover_course.png / cover_year.png
（中间产物 docx/pdf 留在 work\\ 里，方便出问题时对照）
"""

from __future__ import annotations

import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from essay_tool.engine.cover import (  # noqa: E402
    CoverFields,
    build_course_cover_document,
    build_thesis_cover_document,
    build_year_cover_document,
)

OUT_DIR = ROOT / "essay_tool" / "gui" / "assets"
WORK_DIR = ROOT / "work"

# 样张内容（用户 2026-09-19 拍板）：横线全部留空、不放占位；**只有题目位置**
# 放“[论文题目]”当示意（带方括号，表明是占位），其余一律空。
TITLE_PLACEHOLDER = "[论文题目]"
THESIS_PREVIEW = CoverFields(title=TITLE_PLACEHOLDER)
COURSE_PREVIEW = CoverFields()


def build_year_preview_document():
    """学年论文封面样张：直接用引擎里的学年论文封面（＝论文封面把标题行的
    “毕 业”换成“学 年”）。"""
    return build_year_cover_document(THESIS_PREVIEW)

# 渲染倍率：A4 纸 2 倍 ≈ 1190×1684 像素——缩略图清晰，放大窗口也够用
ZOOM = 2.0


def export_pdf(docx_path: pathlib.Path, pdf_path: pathlib.Path) -> None:
    """用本机 Word 把 docx 导成 PDF（ExportAsFixedFormat 17 = PDF）。"""
    import win32com.client

    word = win32com.client.DispatchEx("Word.Application")   # 独立进程，不碰用户开着的 Word
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(str(docx_path), ReadOnly=True,
                                  AddToRecentFiles=False)
        doc.ExportAsFixedFormat(str(pdf_path), 17)
        doc.Close(False)
    finally:
        word.Quit()


def render_png(pdf_path: pathlib.Path, png_path: pathlib.Path) -> tuple[int, int]:
    """PDF 第 1 页 → PNG。"""
    import fitz

    with fitz.open(pdf_path) as doc:
        pix = doc[0].get_pixmap(matrix=fitz.Matrix(ZOOM, ZOOM), alpha=False)
        pix.save(str(png_path))
        return pix.width, pix.height


def bake(name: str, build, out_name: str) -> None:
    docx_path = WORK_DIR / f"封面预览_{name}.docx"
    pdf_path = WORK_DIR / f"封面预览_{name}.pdf"
    png_path = OUT_DIR / out_name
    build().save(str(docx_path))
    export_pdf(docx_path, pdf_path)
    w, h = render_png(pdf_path, png_path)
    print(f"[ok] {out_name}  {w}x{h}px")


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    WORK_DIR.mkdir(exist_ok=True)
    bake("论文", lambda: build_thesis_cover_document(THESIS_PREVIEW),
         "cover_thesis.png")
    bake("课程", lambda: build_course_cover_document(COURSE_PREVIEW),
         "cover_course.png")
    bake("学年", build_year_preview_document, "cover_year.png")
    print("完成 →", OUT_DIR.relative_to(ROOT))


if __name__ == "__main__":
    main()
