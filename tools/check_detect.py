"""用法：.venv\\Scripts\\python.exe tools\\check_detect.py <docx路径>

把标题识别结果和掉号检查结果打出来，人工核对。
"""

import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))

from docx import Document

from essay_tool.engine.detect import (
    check_sequence,
    detect_document_title,
    detect_front_matter_labels,
    detect_headings,
    find_unresolved_auto_numbered_paragraphs,
)


def main():
    path = sys.argv[1]
    doc = Document(path)
    paragraphs = doc.paragraphs
    matches, demoted = detect_headings(doc)
    front_matter = detect_front_matter_labels(doc)
    title = detect_document_title(doc, matches, front_matter)

    front_matter_by_index = {f.paragraph_index: f for f in front_matter}

    for i, (p, m) in enumerate(zip(paragraphs, matches)):
        text = p.text
        if m:
            src = "" if m.source == "regex" else f"  <{m.source}>"
            role = f"  role={m.role!r}" if m.role else ""
            print(f"[{i}] L{m.level} 编号={m.number!s:<4} 标记={m.marker!r:6} 标题={m.title!r}{src}{role}")
        elif i in front_matter_by_index:
            print(f"[{i}]         (前置标签) {front_matter_by_index[i].label!r}  原文={text!r}")
        elif title and i == title.title_paragraph_index:
            print(f"[{i}]         (大标题) {text!r}")
        elif title and i == title.subtitle_paragraph_index:
            print(f"[{i}]         (副标题) {text!r}")
        elif text.strip():
            print(f"[{i}]         (正文) {text!r}")

    print()
    print("--- 不规范标题（已自动识别为正文，弹窗提醒用户） ---")
    if demoted:
        for d in demoted:
            print(f"第{d.paragraph_index}段：L{d.level} {d.marker!r} {d.title!r}  原因={d.reason}")
    else:
        print("无")

    print()
    print("--- 序号检查（未归零/跳号/重号/孤立） ---")
    warnings = check_sequence(matches)
    if warnings:
        for index, name, message in warnings:
            print(f"第{index}段：{name}：{message}")
    else:
        print("无")

    print()
    print("--- Word 自动编号列表，反算不出编号 ---")
    auto = find_unresolved_auto_numbered_paragraphs(doc)
    if auto:
        for i in auto:
            print(f"第{i}段：{paragraphs[i].text!r}")
    else:
        print("无")


if __name__ == "__main__":
    main()
