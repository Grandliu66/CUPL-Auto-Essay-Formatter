"""用法：.venv\\Scripts\\python.exe tools\\dump_docx.py <docx路径> [输出txt路径] [关键字] [--codes]

把 docx 的正文（含表格文字）导成纯文本，每段带 [下标] 前缀——这个下标跟
detect/structure 报出来的段落下标是同一套，方便跟识别结果对照着看。给了
关键字就只**打印**命中的段落，但写进输出文件的仍然是全量。

`--codes`：在每段下面附一行非汉字字符的 Unicode 码点。**这个项目里全角/半角
反复出问题**（规范要求半角的地方、标题编号用哪种句点、空格是全角还是半角），
肉眼根本分不出来，一定要靠码点确认，别靠看。

重新导出规范原文：

    .venv\\Scripts\\python.exe tools\\dump_docx.py ^
        "reference\\中国政法大学本科毕业论文撰写格式规范.docx" --out work\\_spec.txt

核对某几段的码点（比如规范原文里"1．"到底用的哪种句点）：

    .venv\\Scripts\\python.exe tools\\dump_docx.py ^
        "reference\\中国政法大学本科毕业论文撰写格式规范.docx" ^
        --grep "第一级标题应当是" --codes
"""

import argparse
import pathlib
import zipfile

from docx import Document


def codepoint_note(text: str) -> str:
    """把非汉字字符列成 "字符=U+XXXX" 的形式，用来分辨全角/半角。"""
    marks = [f"{c}=U+{ord(c):04X}" for c in text if not ("\u4e00" <= c <= "\u9fff")]
    return "     " + " ".join(marks)


def collect(path: pathlib.Path) -> list[str]:
    doc = Document(path)
    lines = []
    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()
        if text:
            lines.append(f"[{i}] {text}")
    # 表格里的文字也要——规范里的数值表往往在表格里，不导就漏了
    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            cells = [c.text.strip() for c in row.cells]
            if any(cells):
                lines.append(f"[表{ti}-{ri}] " + " | ".join(cells))
    return lines


def main() -> None:
    parser = argparse.ArgumentParser(description="把 docx 正文导成纯文本")
    parser.add_argument("path", type=pathlib.Path, help="要导出的 docx")
    # out / keyword 用选项而不是位置参数：位置参数让调用方没法跳过前一个，
    # 而 PowerShell 会把空字符串参数直接吃掉，一错位就写出一堆乱名字的文件。
    parser.add_argument("--out", type=pathlib.Path, help="输出 txt 路径，省略则只打印")
    parser.add_argument("--grep", dest="keyword", help="只打印含这个关键字的段落")
    parser.add_argument("--codes", action="store_true", help="附上非汉字字符的 Unicode 码点")
    parser.add_argument(
        "--part",
        help="直接打印包里某个部件的原始 XML（如 word/footnotes.xml）——脚注之类的内容不在正文段落里，只能这样看",
    )
    args = parser.parse_args()

    if args.part:
        with zipfile.ZipFile(args.path) as package:
            print(package.read(args.part).decode("utf-8"))
        return

    lines = collect(args.path)
    if args.out:
        args.out.write_text("\n".join(lines), encoding="utf-8")
        print(f"已导出：{args.out}（共 {len(lines)} 行）")

    hits = [line for line in lines if args.keyword in line] if args.keyword else lines
    if args.out is None or args.keyword:
        if args.keyword:
            print(f"\n===== 含「{args.keyword}」的段落（{len(hits)} 条） =====")
        for line in hits:
            print(line)
            if args.codes:
                # 行首的 "[下标] " 前缀不参与码点
                body = line.split("] ", 1)[1] if "] " in line else line
                print(codepoint_note(body))


if __name__ == "__main__":
    main()
