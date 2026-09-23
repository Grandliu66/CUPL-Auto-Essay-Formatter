"""把 Excel 里人能看懂的写法（"小四"“2字符”“固定值 20磅”）换算成 python-docx 认识的值。

解析失败时抛 ValueError，附中文说明；excel_io.py 负责把它包装成带单元格坐标的报错。
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from enum import Enum

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Cm, Emu, Pt

# 中文字号 -> 磅值（国标印刷字号表）
CHINESE_FONT_SIZES: dict[str, float] = {
    "初号": 42, "小初": 36,
    "一号": 26, "小一": 24,
    "二号": 22, "小二": 18,
    "三号": 16, "小三": 15,
    "四号": 14, "小四": 12,
    "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5,
    "七号": 5.5,
    "八号": 5,
}

ALIGNMENTS: dict[str, WD_ALIGN_PARAGRAPH] = {
    "两端对齐": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "居中": WD_ALIGN_PARAGRAPH.CENTER,
    "左对齐": WD_ALIGN_PARAGRAPH.LEFT,
    "右对齐": WD_ALIGN_PARAGRAPH.RIGHT,
}

LINE_RULES: dict[str, WD_LINE_SPACING] = {
    "多倍行距（倍）": WD_LINE_SPACING.MULTIPLE,
    "固定值（磅）": WD_LINE_SPACING.EXACTLY,
    # 下面四个是格式页「行距」下拉框里原来就有、但这张表一直没收的写法
    # （2026-09-22 审计揪出的真 bug：格式页能选、一保存/一导出就报"无法识别的
    # 行距类型"，流程直接卡死）。"最小值（磅）"跟"固定值（磅）"一样要磅数，
    # 见 `parse_line_value`；"单倍/双倍/1.5倍"是固定比例，不需要数值。
    "最小值（磅）": WD_LINE_SPACING.AT_LEAST,
    "单倍": WD_LINE_SPACING.SINGLE,
    "双倍": WD_LINE_SPACING.DOUBLE,
    "1.5倍": WD_LINE_SPACING.ONE_POINT_FIVE,
}

# 这三种行距不需要"行距值"（比例是固定的，python-docx 赋值 line_spacing_rule
# 时会自己把 w:line 也一并写好，见 `engine.styles`）——`parse_line_value` 对
# 它们一律返回 None，不管"行距值"那一格里写的是什么。
_NO_VALUE_LINE_RULES = frozenset({
    WD_LINE_SPACING.SINGLE, WD_LINE_SPACING.DOUBLE, WD_LINE_SPACING.ONE_POINT_FIVE,
})

PAGE_NUMBER_FORMATS: dict[str, str] = {
    "1,2,3": "decimal",
    "I,II,III": "upperRoman",
    "i,ii,iii": "lowerRoman",
    "A,B,C": "upperLetter",
    "a,b,c": "lowerLetter",
}

_LENGTH_UNIT_RE = re.compile(r"^(-?\d+(?:\.\d+)?)\s*(厘米|cm|毫米|mm|磅|pt)$", re.IGNORECASE)
_MULTIPLE_RE = re.compile(r"^(\d+(?:\.\d+)?)\s*倍?$")

# 纸张规格名字 -> (宽, 高)，按"纵向"下的标准尺寸给
NAMED_PAPER_SIZES: dict[str, tuple[Emu, Emu]] = {
    "A4": (Cm(21), Cm(29.7)),
    "A3": (Cm(29.7), Cm(42)),
    "B5": (Cm(17.6), Cm(25)),
}


def parse_cm(text) -> Emu:
    """单位已经写进表头/项目名了，这里的数据格只放纯数字，单位固定厘米。"""
    text = "" if text is None else str(text).strip()
    try:
        return Cm(float(text))
    except ValueError:
        raise ValueError(f'无法识别的长度数值："{text}"（应为纯数字，单位厘米）') from None


def parse_font_size(text: str) -> Pt:
    """接受"小四"这类中文字号，"12磅"/"12pt"这类带单位数值，或"12"这类纯阿拉伯数字（按磅算）。"""
    text = str(text).strip()
    if text in CHINESE_FONT_SIZES:
        return Pt(CHINESE_FONT_SIZES[text])
    m = _LENGTH_UNIT_RE.match(text)
    if m and m.group(2).lower() in ("磅", "pt"):
        return Pt(float(m.group(1)))
    try:
        return Pt(float(text))
    except ValueError:
        raise ValueError(f'无法识别的字号写法："{text}"（应为"小四"这类中文字号，或"12"这类数字）') from None


class IndentKind(Enum):
    NONE = "none"
    CHARS = "chars"
    LENGTH = "length"


@dataclass(frozen=True)
class Indent:
    """缩进既可能是"2字符"（按当前字号换算，写入 w:firstLineChars 等属性），
    也可能是"0.85厘米"这种固定长度。两种写法在 Word 里含义不同，必须分开保留，
    不能都强行换算成厘米。
    """

    kind: IndentKind
    chars: float = 0.0
    length: Emu | None = None


def _parse_plain_number(text) -> float | None:
    """单位和数值现在是分开两列；"", "-", "0" 都当作"没有"。"""
    text = "" if text is None else str(text).strip()
    if text in ("", "-", "0"):
        return None
    try:
        return float(text)
    except ValueError:
        raise ValueError(f'无法识别的数值："{text}"（应为纯数字）') from None


def parse_indent_pair(unit, value) -> Indent:
    """"首行缩进/左缩进"现在拆成"单位"+"值"两列。"""
    num = _parse_plain_number(value)
    if num is None:
        return Indent(IndentKind.NONE)
    unit_text = "" if unit is None else str(unit).strip()
    if unit_text == "字符":
        return Indent(IndentKind.CHARS, chars=num)
    if unit_text == "厘米":
        return Indent(IndentKind.LENGTH, length=Cm(num))
    raise ValueError(f'无法识别的缩进单位："{unit_text}"（应为"字符"或"厘米"）')


def parse_optional_length_pair(unit, value) -> Emu | None:
    """"段前/段后"现在拆成"单位"+"值"两列，目前单位只支持"磅"。"""
    num = _parse_plain_number(value)
    if num is None:
        return None
    unit_text = "" if unit is None else str(unit).strip()
    if unit_text == "磅":
        return Pt(num)
    raise ValueError(f'无法识别的单位："{unit_text}"（目前只支持"磅"）')


def parse_yn(text) -> bool:
    text = str(text).strip()
    if text == "是":
        return True
    if text == "否":
        return False
    raise ValueError(f'无法识别的是否写法："{text}"（应为"是"或"否"）')


def parse_alignment(text) -> WD_ALIGN_PARAGRAPH:
    text = str(text).strip()
    if text not in ALIGNMENTS:
        raise ValueError(f'无法识别的对齐方式："{text}"（应为{"/".join(ALIGNMENTS)}）')
    return ALIGNMENTS[text]


def parse_line_rule(text) -> WD_LINE_SPACING:
    text = str(text).strip()
    if text not in LINE_RULES:
        raise ValueError(f'无法识别的行距类型："{text}"（应为{"/".join(LINE_RULES)}）')
    return LINE_RULES[text]


def parse_bold(text) -> bool:
    return parse_yn(text)


def parse_line_value(rule: WD_LINE_SPACING, text) -> Emu | float | None:
    """"行距值"这一列的含义取决于"行距类型"：固定值/最小值是磅数，多倍行距是倍数；
    单倍/双倍/1.5倍这三种比例固定，不需要行距值，不管这一格写了什么都返回 None。
    单位已经写进"行距类型"里了（"多倍行距（倍）"/"固定值（磅）"/"最小值（磅）"），
    要数值的那两种这里只接受纯数字。
    """
    if rule in _NO_VALUE_LINE_RULES:
        return None
    text = "" if text is None else str(text).strip()
    if rule == WD_LINE_SPACING.MULTIPLE:
        m = _MULTIPLE_RE.match(text)
        if not m:
            raise ValueError(f'无法识别的倍数写法："{text}"（应为纯数字，如"1"或"1.5"）')
        return float(m.group(1))
    try:
        return Pt(float(text))
    except ValueError:
        raise ValueError(f'无法识别的行距值："{text}"（应为纯数字，如"20"）') from None


def parse_page_number_format(text) -> str:
    text = "" if text is None else str(text).strip()
    if text not in PAGE_NUMBER_FORMATS:
        raise ValueError(f'无法识别的页码编码："{text}"（应为{"/".join(PAGE_NUMBER_FORMATS)}）')
    return PAGE_NUMBER_FORMATS[text]
