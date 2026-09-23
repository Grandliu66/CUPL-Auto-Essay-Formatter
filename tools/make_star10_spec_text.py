# -*- coding: utf-8 -*-
"""把学术十星格式要求 PDF 抽成内置文本模块（essay_tool/gui/star10_spec_text.py）。

为什么"生成"而不是"运行期读 PDF"：发布包不带 PyMuPDF（见 main_window 里
第三方声明那段注释），所以在开发期把原文抽好、生成一个 Python 常量模块，
运行时零新依赖（同《工具操作逻辑》正文的做法，但那个是手写维护的）。

抽法（PyMuPDF 按"块"取，块≈自然段）：
- 块内换行拼成一行（英文单词跨行断开时补一个空格）；
- 跨块断开的段（上一段末尾没有句末标点、且两段都不是标题）拼回去
  —— 实测这份 PDF 有几个段正好切在页边界（"…Times New Roman 打" / "印，…"），
  第 1 页导语也被按行拆成了五块；
- 大字号块（≥13 磅：题名 15.9、各级标题 14.1）不参与拼接（题名两行本来
  就不是一段）；各级标题另有编号模式兜底（1.提交要求 这类是正文号字的）；
- 页眉的 "- N -" 页码块丢掉。
正文一个字不改。

用法：
    .venv\\Scripts\\python.exe tools\\make_star10_spec_text.py           # 生成模块
    .venv\\Scripts\\python.exe tools\\make_star10_spec_text.py --show    # 只打印统计与样本，不写文件
"""
from __future__ import annotations

import json
import pathlib
import re
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from essay_tool.engine.detect import _try_numbered  # noqa: E402

REF = ROOT / "reference"
OUT = ROOT / "essay_tool" / "gui" / "star10_spec_text.py"

_PAGE_NUM = re.compile(r"^[-—–]\s*\d+\s*[-—–]$")
_APPENDIX = re.compile(r"^附录[一二三四五六七八九十\d]+[：:、]")
# 子项行（"1）刻本"这类）：短行，不该被拼到下一段（实测）
_SUB_ITEM = re.compile(r"^\d{1,2}\s*）")
# 句末标点（这些收尾的段不会被拼到下一段；"）》《"也算——实测有段正好以
# "（详见《…封面范例》）"收尾，下一段是"注意：…"，不该拼。
# ⚠️ 右引号 "”" 不算——实测"…统一使用'西文字体' / 中的'Times New Roman'。"
# 就是被它误断成两段的。
_TERMINAL = "。！？；：…—）》』】"
# 该字号以上的块视作题名/标题（实测：题名 15.9、标题 14.1、正文 12.0）
_FLOAT_SIZE = 13.0


def _looks_heading(text: str) -> bool:
    first = text.splitlines()[0].strip() if text else ""
    return (bool(_APPENDIX.match(first)) or bool(_SUB_ITEM.match(first))
            or _try_numbered(first) is not None)


# 短于此长度的"标题行"既不吸收下文也不被吸收；超过的说明是标题自身换行
# （实测："（4）引文注释格式要求…关于引文注释的规" / "定》" 被断成两段）
_TITLE_BLOCK_MAX = 24


def _blocking_title(text: str) -> bool:
    return _looks_heading(text) and len(text) <= _TITLE_BLOCK_MAX


def _join_lines(lines: list[str]) -> str:
    out = ""
    for ln in lines:
        if not out:
            out = ln
            continue
        if (out[-1].isascii() and out[-1].isalnum()
                and ln[0].isascii() and ln[0].isalnum()):
            out += " " + ln
        else:
            out += ln
    return out


def extract_paragraphs(pdf_path: pathlib.Path,
                       debug_merges: bool = False) -> list[str]:
    import fitz                                     # PyMuPDF：只在开发期用

    blocks: list[tuple[str, float]] = []            # (文本, 该块最大字号)
    with fitz.open(str(pdf_path)) as doc:
        for page in doc:
            for block in page.get_text("dict", sort=True)["blocks"]:
                if block["type"] != 0:              # 图片块
                    continue
                lines: list[str] = []
                size = 0.0
                for line in block["lines"]:
                    text = "".join(s["text"] for s in line["spans"]).strip()
                    for span in line["spans"]:
                        size = max(size, span["size"])
                    if text:
                        lines.append(text)
                if not lines:
                    continue
                text = _join_lines(lines)
                if _PAGE_NUM.match(text):
                    continue
                blocks.append((text, size))

    merged: list[tuple[str, float]] = []
    merges: list[str] = []
    for text, size in blocks:
        big = size >= _FLOAT_SIZE                   # 题名/标题块：不参与拼接
        if merged:
            prev, prev_size = merged[-1]
            if (not big and prev_size < _FLOAT_SIZE
                    and prev and prev[-1] not in _TERMINAL
                    and not _blocking_title(prev) and not _looks_heading(text)):
                sep = (" " if (prev[-1].isascii() and prev[-1].isalnum()
                              and text[0].isascii() and text[0].isalnum()) else "")
                merges.append(f"{prev[-20:]}  ⨝  {text[:20]}")
                merged[-1] = (prev + sep + text, prev_size)
                continue
        merged.append((text, size))

    if debug_merges:
        print(f"— 拼接了 {len(merges)} 处 —")
        for m in merges:
            print("   " + m)
    return [text for text, _ in merged]


def write_module(paras: list[str], source_name: str) -> None:
    head = (
        "# -*- coding: utf-8 -*-\n"
        '"""学术十星格式要求原文（内置文本，供查阅窗口 `spec_viewer.open_star10_viewer` 用）。\n'
        "\n"
        "由 `tools/make_star10_spec_text.py` 从官方 PDF 自动提取生成，**勿手改**；\n"
        "换届数或规范更新时重跑生成器。原文一字未改（只做块内换行拼接、跨页断段\n"
        f"拼接、去页码）。共 {len(paras)} 段。\n"
        f"源文件：{source_name}\n"
        '"""\n'
        "\n"
        "STAR10_SPEC_TEXTS: tuple[str, ...] = (\n"
    )
    body = "".join("    " + json.dumps(p, ensure_ascii=False) + ",\n" for p in paras)
    OUT.write_text(head + body + ")\n", encoding="utf-8")


def main() -> int:
    pdf = next((p for p in sorted(REF.glob("*学术十星*.pdf"))
                if not p.name.startswith("~$")), None)
    if pdf is None:
        print("找不到学术十星 PDF（reference/ 里文件名带'学术十星'的 .pdf）。")
        return 1
    paras = extract_paragraphs(pdf, debug_merges=("--merges" in sys.argv))
    print(f"源文件：{pdf.name}")
    print(f"抽出 {len(paras)} 段")
    if "--show" in sys.argv:
        print("\n— 样本（首 25 段）—")
        for i, p in enumerate(paras[:25]):
            print(f"[{i:3d}] {p[:96]}")
        print("\n— 样本（末 12 段）—")
        for i, p in enumerate(paras[-12:], start=len(paras) - 12):
            print(f"[{i:3d}] {p[:96]}")
        return 0
    write_module(paras, pdf.name)
    print(f"已写入 {OUT.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
