"""烘学术十星封面样张（硕博封面 2026-09-19 挂起、素材已归档，相关条目移除）。

链路：reference 里的官方封面 docx（十星先摘掉页眉里的文字水印）→ 本机
Word 导出 PDF → PyMuPDF 第 1 页转 PNG → essay_tool/gui/assets/。

按用户 2026-09-19 的交代：
- 横线全部留空，不填演示数据、不放占位；只有题目位置写“[论文题目]”；
- 十星的“第二十四届”是文件自带文字，暂时照原样保留，将来届数由用户填；
- 十星页眉里的“学术十星”文字水印要摘掉（它是 Word 水印，挂在页眉上，
  正文与版面其余部分不受影响）。

用法：.venv\\Scripts\\python.exe tools\\make_cover_preview_extra.py
产出：essay_tool\\gui\\assets\\cover_tenstar.png
中间产物（无水印 docx、各 PDF）留在 work\\ 里，方便后面对照。
"""

from __future__ import annotations

import pathlib
import re
import zipfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
REF = ROOT / "reference"
WORK = ROOT / "work"
OUT = ROOT / "essay_tool" / "gui" / "assets"

ZOOM = 2.0   # 与 make_cover_preview.py 一致：A4 两倍 ≈ 1190×1684 像素


def strip_watermark(src: pathlib.Path, dst: pathlib.Path) -> None:
    """把 docx 里各页眉上的 w:pict（文字水印图形块）摘掉，其余部件原样复制。"""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("word/header") and item.filename.endswith(".xml"):
                text = data.decode("utf-8")
                text, n = re.subn(r"<w:pict\b[^>]*>.*?</w:pict>", "", text, flags=re.S)
                left = text.count("textpath")
                print(f"    {item.filename}: 摘掉 {n} 个图形块，剩余水印痕迹 {left}")
                data = text.encode("utf-8")
            zout.writestr(item, data)


def mark_title_placeholder(docx_path: pathlib.Path) -> None:
    """题目位置的示意文字：官方模板自带“论文题目”，样张统一写成“[论文题目]”
    （与论文/学年论文样张一致，带方括号表明是占位）。其余位置不放占位。"""
    from docx import Document

    doc = Document(str(docx_path))
    for p in doc.paragraphs:
        if p.text.strip() == "论文题目":
            for t in p._p.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"):
                if t.text:
                    t.text = t.text.replace("论文题目", "[论文题目]")
            break
    doc.save(str(docx_path))


def export_pdf(docx_path: pathlib.Path, pdf_path: pathlib.Path) -> None:
    """用本机 Word 把 docx 导成 PDF（独立进程，不碰用户开着的 Word）。"""
    import win32com.client

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(str(docx_path), ReadOnly=True, AddToRecentFiles=False)
        try:
            doc.ExportAsFixedFormat(str(pdf_path), 17)
        finally:
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


def bake_one(name: str, source: pathlib.Path, out_name: str) -> None:
    pdf_path = WORK / f"封面预览_{name}.pdf"
    png_path = OUT / out_name
    export_pdf(source, pdf_path)
    wpx, hpx = render_png(pdf_path, png_path)
    print(f"[ok] {out_name}  {wpx}x{hpx}px  <- {source.name}")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)

    # 跳过 Word 打开原文件时留下的 `~$` 锁文件（不是合法 docx）
    star_src = next((p for p in REF.glob("*学术十星*封面*.docx")
                     if not p.name.startswith("~$")), None)
    if star_src is None:
        print("[跳过] 找不到学术十星封面文件")
    else:
        star_clean = WORK / "封面预览_十星_无水印.docx"
        print("[处理] 学术十星：摘页眉水印")
        strip_watermark(star_src, star_clean)
        mark_title_placeholder(star_clean)
        bake_one("十星", star_clean, "cover_tenstar.png")

    print("完成 →", OUT.relative_to(ROOT))


if __name__ == "__main__":
    main()
