"""反算 Word 自动编号列表实际会渲染出的编号文字。

Word 的"多级列表"编号不写在段落文字里，是 Word 根据 numbering.xml 里的规则、
按段落在文档里出现的顺序现算现画的。要认出这类段落是不是标题，得先自己把
这套算法实现一遍：解析 abstractNum/num 定义，按文档顺序推进每个 numId 各层级
的计数器，再按 numFmt（十进制/中文数字/罗马数字/字母/带圈数字……）和 lvlText
模板（比如"%1."、"(%1)"）拼出最终文字。

范围说明：只处理"遇到更浅层级就把更深层级计数器清零重来"这一种（也是最常见
的）重置规则，不处理 <w:lvlRestart> 自定义重置层级；<w:num> 上的
<w:lvlOverride><w:startOverride>（用户在 Word 里手动"重新开始编号"时会生成）
支持。bullet/none 等不是"数字编号"的格式，判定为渲染不出编号，返回 None。

反方向（写入）也在这个文件里：`ensure_reference_numbering()` 往 numbering.xml
**新建**（或复用）一个按预设写法的编号定义（`[%1]` / `%1.` / `(%1)`，见
`REFERENCE_LEVEL_TEXTS`），给参考文献条目用。序号必须是真
的自动编号——参考文献是全文最常增删的一节，手打的话每加一条都得把后面全部
重编一遍。
"""

from __future__ import annotations

import copy

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.exceptions import InvalidXmlError

try:
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
except ImportError:  # pragma: no cover
    RT = None


# ---------------------------------------------------------------- 写入方向

# 参考文献条目的编号写法 → Word 的 lvlText 模板（`%1` 是第一个计数器的占位符）。
# 键就是预设「引注」里 `reference_numbering_style` 的取值（GUI 那个
# 「参考文献编号样式」下拉里也是这四个字），值是要写进 numbering.xml 的模板。
# 用户 2026-09-16：“当用户直接用自动编号排参考文献时，重新编号失效，没有调整为
# 指定编号格式。”——以前写死成 `[%1]`，预设选别的写法也不生效。
REFERENCE_LEVEL_TEXTS: dict[str, str] = {
    "[1][2][3]": "[%1]",
    "1.2.3.": "%1.",
    "(1)(2)(3)": "(%1)",
}
DEFAULT_REFERENCE_STYLE = "[1][2][3]"


def reference_level_text(style: str | None) -> str:
    """预设写法 → lvlText 模板；空值或认不出的写法退回默认（方括号）。"""
    key = (style or "").strip()
    return REFERENCE_LEVEL_TEXTS.get(key, REFERENCE_LEVEL_TEXTS[DEFAULT_REFERENCE_STYLE])


# 旧名字（默认写法就是方括号那支），保留别名免得别处引用不到。
_BRACKET_LEVEL_TEXT = REFERENCE_LEVEL_TEXTS[DEFAULT_REFERENCE_STYLE]


def _level_text(abstract, level: str = "0") -> str | None:
    for element in abstract.findall(qn("w:lvl")):
        if element.get(qn("w:ilvl")) != level:
            continue
        lvl_text = element.find(qn("w:lvlText"))
        if lvl_text is not None:
            return lvl_text.get(qn("w:val"))
    return None


def _level_attr(abstract, tag: str, level: str = "0") -> str | None:
    """第 `level` 层某个子元素（`w:numFmt`/`w:start`……）的 `w:val`。"""
    for element in abstract.findall(qn("w:lvl")):
        if element.get(qn("w:ilvl")) != level:
            continue
        child = element.find(qn(tag))
        return child.get(qn("w:val")) if child is not None else None
    return None


def _next_free_id(numbering) -> int:
    """挑一个没被占用的 id。`w:numId="0"` 是"这一段没有编号"的保留值，避开。"""
    used: set[int] = set()
    for element in numbering:
        for attribute in ("w:numId", "w:abstractNumId"):
            value = element.get(qn(attribute))
            if value is None:
                continue
            try:
                used.add(int(value))
            except ValueError:
                pass
    return max(used, default=0) + 1


def _build_reference_abstract(abstract_id: int, lvl_text_value: str):
    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))

    # 子元素顺序按 schema 来（跟别处手工建元素一样，位置错了 Word 可能报
    # "内容有问题"）：abstractNum 里 multiLevelType 排在 lvl 前；lvl 里是
    # start, numFmt, lvlText, lvlJc——中间那几项不用就不写，相对顺序不能乱。
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "singleLevel")
    abstract.append(multi)

    level = OxmlElement("w:lvl")
    level.set(qn("w:ilvl"), "0")
    for tag, value in (("w:start", "1"), ("w:numFmt", "decimal")):
        element = OxmlElement(tag)
        element.set(qn("w:val"), value)
        level.append(element)
    lvl_text = OxmlElement("w:lvlText")
    lvl_text.set(qn("w:val"), lvl_text_value)
    level.append(lvl_text)
    lvl_jc = OxmlElement("w:lvlJc")
    lvl_jc.set(qn("w:val"), "left")
    level.append(lvl_jc)

    abstract.append(level)
    return abstract


def _numbering_element(document):
    """取 numbering 根元素；文档没有 numbering.xml 就**手动建一个空部件**。

    为什么必须手动建：python-docx 的 `NumberingPart.new()` 是个未实现的空壳
    （1.2.0 仍然如此），靠“没有就自动补”的写法在缺 numbering.xml 的文档上直接
    NotImplementedError（2026-09-13 `test2.docx` 实拍——很多导出器就不带这个
    部件）。手动建的要点：部件用 `NumberingPart`（而不是裸 Part），这样后续
    `document.part.numbering_part` 读回来的对象照样有 `.element`；部件挂到
    document.part 的关系上，Content_Types 保存时由 python-docx 按 content_type
    自动登记。
    """
    try:
        return document.part.numbering_part.element
    except (KeyError, NotImplementedError):
        pass

    from docx.opc.constants import CONTENT_TYPE as CT
    from docx.opc.packuri import PackURI
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    from docx.parts.numbering import NumberingPart

    element = parse_xml(f"<w:numbering {nsdecls('w')}/>")
    part = NumberingPart(
        PackURI("/word/numbering.xml"),
        CT.WML_NUMBERING,
        element,
        document.part.package,
    )
    document.part.relate_to(part, RT.NUMBERING)
    return element


def _abstract_of(document, num_id: int):
    """某个 numId 用的 abstractNum 元素（找不到返回 None）。"""
    numbering = _numbering_element(document)
    abstracts = {
        abstract.get(qn("w:abstractNumId")): abstract
        for abstract in numbering.findall(qn("w:abstractNum"))
    }
    for num in numbering.findall(qn("w:num")):
        if num.get(qn("w:numId")) != str(num_id):
            continue
        reference = num.find(qn("w:abstractNumId"))
        if reference is None:
            return None
        return abstracts.get(reference.get(qn("w:val")))
    return None


def paragraph_num_id(paragraph) -> int | None:
    """这一段现在挂着哪个自动编号（没有就 None）。"""
    pPr = paragraph._p.find(qn("w:pPr"))
    numPr = pPr.find(qn("w:numPr")) if pPr is not None else None
    if numPr is None:
        return None
    num_id = numPr.find(qn("w:numId"))
    if num_id is None:
        return None
    try:
        return int(num_id.get(qn("w:val")))
    except (TypeError, ValueError):
        return None


def copy_level_indent(document, from_num_id: int, to_num_id: int) -> None:
    """把"用户原来那条编号"的段落属性（悬挂缩进之类）搬到我们的定义上。

    用户用 Word 自动编号排参考文献时，缩进常常写**在编号定义里**（实测
    `民总-排好1版.docx`：`w:ind left=425 hanging=425`）。我们把条目改指到
    自己的定义之后那一块就没了，条目会顶着左边距、转行也不对齐。所以换定义
    之前把 `w:lvl/w:pPr` 搬一次。目标定义已经有了就不动（不覆盖我们自己的）。
    """
    if from_num_id == to_num_id:
        return
    source = _abstract_of(document, from_num_id)
    target = _abstract_of(document, to_num_id)
    if source is None or target is None:
        return
    pPr = _level_pr(source)
    if pPr is None:
        return
    for level in target.findall(qn("w:lvl")):
        if level.get(qn("w:ilvl")) != "0":
            continue
        if level.find(qn("w:pPr")) is not None:
            return                      # 已经有了，别覆盖
        # lvl 的子元素顺序：start, numFmt, lvlText, lvlJc, pPr, rPr——pPr 必须
        # 排在 rPr 前面。以前 `lvlJc` 不存在时直接 `level.append(...)`，如果
        # 这一层恰好已经有 `rPr`（没有 `lvlJc`、只有 `rPr`），pPr 会被摆到
        # rPr 后面，顺序就反了（2026-09-22 审计揪出的）。改成找不到 `lvlJc`
        # 就退一步找 `rPr`，插到它前面；两个都没有才真的直接 append。
        lvl_jc = level.find(qn("w:lvlJc"))
        if lvl_jc is not None:
            lvl_jc.addnext(copy.deepcopy(pPr))
        else:
            rPr = level.find(qn("w:rPr"))
            if rPr is not None:
                rPr.addprevious(copy.deepcopy(pPr))
            else:
                level.append(copy.deepcopy(pPr))
        return


def _level_pr(abstract):
    """编号定义里第 0 层的 `w:pPr`（悬挂缩进等）。"""
    for element in abstract.findall(qn("w:lvl")):
        if element.get(qn("w:ilvl")) != "0":
            continue
        return element.find(qn("w:pPr"))
    return None


def ensure_reference_numbering(document, style: str | None = None,
                               existing_ids=()) -> int:
    """确保文档里有一个"按预设写法"的参考文献编号定义，返回它的 numId。

    `style` 是预设 `citation.reference_numbering_style`（`[1][2][3]` /
    `1.2.3.` / `(1)(2)(3)`），对应 lvlText `[%1]` / `%1.` / `(%1)`。

    `existing_ids`：**参考文献条目自己当前实际在用**的 numId 集合（调用方
    在改动任何条目之前先扫一遍传进来，见 `apply_from_index`）。只在这个
    集合里找"lvlText 已经一致、可以直接复用"的定义——**不能**像以前那样
    扫全篇 numbering.xml 随便撞上一个格式凑巧一样的 `w:num`就复用：那很可能
    是正文里跟参考文献毫不相干的另一份编号列表（用户自己的多级列表、目录
    条目……），복用它俩就共用同一个计数器了，参考文献会接着正文列表的号
    往下数，而且下次回炉再跑这个函数，"先找到哪个"还可能变，编号跟着换
    （2026-09-22 审计揪出的真 bug）。`existing_ids` 为空（没有条目、或条目都
    还没有任何编号）时直接新建，不去猜。
    """
    wanted = reference_level_text(style)
    numbering = _numbering_element(document)

    abstracts = {
        abstract.get(qn("w:abstractNumId")): abstract
        for abstract in numbering.findall(qn("w:abstractNum"))
    }
    num_to_abstract_id = {
        num.get(qn("w:numId")): num.find(qn("w:abstractNumId"))
        for num in numbering.findall(qn("w:num"))
    }
    for existing_id in existing_ids:
        reference = num_to_abstract_id.get(str(existing_id))
        if reference is None:
            continue
        abstract = abstracts.get(reference.get(qn("w:val")))
        if abstract is None or _level_text(abstract) != wanted:
            continue
        # 光比 lvlText 模板不够——`_build_reference_abstract` 造出来的定义固定是
        # `numFmt=decimal`、`start=1`；外部已有的定义恰好用了同一套写法（比如
        # `%1.`）但编号方式或起始值不同（罗马数字、字母编号、从别的数起），
        # lvlText 单看字符串会一样，误判成"能复用"就会套错编号
        # （2026-09-22 审计揪出的）。三项都对得上才真的能直接复用。
        if (_level_attr(abstract, "w:numFmt") == "decimal"
                and _level_attr(abstract, "w:start") == "1"):
            return int(existing_id)

    new_id = _next_free_id(numbering)
    abstract = _build_reference_abstract(new_id, wanted)
    # CT_Numbering 的顺序是：numPicBullet*、abstractNum*、num*。所以 abstractNum
    # 必须插到**第一个 w:num 之前**，直接 append 会排在 num 后面，Word 会当成
    # 无效定义。
    first_num = numbering.find(qn("w:num"))
    if first_num is None:
        numbering.append(abstract)
    else:
        first_num.addprevious(abstract)

    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(new_id))
    reference = OxmlElement("w:abstractNumId")
    reference.set(qn("w:val"), str(new_id))
    num.append(reference)
    numbering.append(num)
    return new_id


# 旧名字：默认写法（方括号）就是它，保留别名免得别处（含历史脚本）引用不到。
ensure_bracket_numbering = ensure_reference_numbering


def _int_to_chinese(n: int) -> str:
    if n <= 0:
        return str(n)
    digits = "〇一二三四五六七八九"
    units = ["", "十", "百", "千"]
    s = str(n)
    length = len(s)
    parts: list[str] = []
    for i, ch in enumerate(s):
        d = int(ch)
        power = length - i - 1
        if d == 0:
            if parts and parts[-1] != digits[0]:
                parts.append(digits[0])
            continue
        if d == 1 and power == 1 and i == 0:
            parts.append(units[power])  # "十一" 不写成"一十一"
        else:
            parts.append(digits[d] + units[power])
    text = "".join(parts).rstrip(digits[0])
    return text or digits[0]


def _int_to_roman(n: int) -> str:
    table = [
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
        (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
        (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
    ]
    result = []
    for value, sym in table:
        count, n = divmod(n, value)
        result.append(sym * count)
    return "".join(result)


def _int_to_letter(n: int) -> str:
    result = []
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result.append(chr(ord("a") + rem))
    return "".join(reversed(result))


def _int_to_circled(n: int) -> str | None:
    if 1 <= n <= 20:
        return chr(0x2460 + n - 1)
    return None  # 带圈数字超过20没有对应的通用字符，不硬凑


_FORMATTERS = {
    "decimal": lambda n: str(n),
    "decimalZero": lambda n: f"{n:02d}",
    "upperRoman": _int_to_roman,
    "lowerRoman": lambda n: _int_to_roman(n).lower(),
    "upperLetter": lambda n: _int_to_letter(n).upper(),
    "lowerLetter": _int_to_letter,
    "chineseCounting": _int_to_chinese,
    "chineseCountingThousand": _int_to_chinese,
    "ideographDigital": _int_to_chinese,
    "decimalEnclosedCircle": _int_to_circled,
}


def _format_number(num_fmt: str, n: int) -> str | None:
    fn = _FORMATTERS.get(num_fmt)
    if fn is None:
        return None  # bullet/none 等不是编号，或者是我们没支持的格式
    return fn(n)


def _parse_abstract_nums(root) -> dict[str, dict[int, dict]]:
    result: dict[str, dict[int, dict]] = {}
    for abstract in root.findall(qn("w:abstractNum")):
        abstract_id = abstract.get(qn("w:abstractNumId"))
        levels: dict[int, dict] = {}
        for lvl in abstract.findall(qn("w:lvl")):
            ilvl_raw = lvl.get(qn("w:ilvl"))
            if ilvl_raw is None:
                continue
            ilvl = int(ilvl_raw)
            start_el = lvl.find(qn("w:start"))
            start = int(start_el.get(qn("w:val"))) if start_el is not None else 1
            num_fmt_el = lvl.find(qn("w:numFmt"))
            num_fmt = num_fmt_el.get(qn("w:val")) if num_fmt_el is not None else "decimal"
            lvl_text_el = lvl.find(qn("w:lvlText"))
            lvl_text = lvl_text_el.get(qn("w:val")) if lvl_text_el is not None else "%1."
            levels[ilvl] = {"start": start, "numFmt": num_fmt, "lvlText": lvl_text or ""}
        result[abstract_id] = levels
    return result


def _parse_nums(root) -> dict[str, dict]:
    result: dict[str, dict] = {}
    for num in root.findall(qn("w:num")):
        num_id = num.get(qn("w:numId"))
        abstract_id_el = num.find(qn("w:abstractNumId"))
        abstract_id = abstract_id_el.get(qn("w:val")) if abstract_id_el is not None else None
        overrides: dict[int, int] = {}
        for lvl_override in num.findall(qn("w:lvlOverride")):
            ilvl_raw = lvl_override.get(qn("w:ilvl"))
            if ilvl_raw is None:
                continue
            start_override_el = lvl_override.find(qn("w:startOverride"))
            if start_override_el is not None:
                overrides[int(ilvl_raw)] = int(start_override_el.get(qn("w:val")))
        result[num_id] = {"abstractNumId": abstract_id, "overrides": overrides}
    return result


class NumberingScheme:
    """有状态：按段落在文档里出现的顺序调用 render()，内部计数器会跟着推进。"""

    def __init__(self, root):
        if root is None:
            self._abstract: dict[str, dict[int, dict]] = {}
            self._nums: dict[str, dict] = {}
        else:
            self._abstract = _parse_abstract_nums(root)
            self._nums = _parse_nums(root)
        self._counters: dict[str, dict[int, int]] = {}
        # 已经**真正渲染过**的级别（`num_id -> {ilvl}`）。用来区分"这个计数器
        # 是自己递增出来的"还是"为了让 %1 这类占位符有值、顺手垫上的"——
        # 垫上的那种不能算数（见 `render()`）。
        self._rendered: dict[str, set[int]] = {}

    def _level_info(self, num_id: str, ilvl: int) -> tuple[int, str, str] | None:
        num = self._nums.get(num_id)
        if num is None:
            return None
        levels = self._abstract.get(num["abstractNumId"])
        if levels is None:
            return None
        info = levels.get(ilvl)
        if info is None:
            return None
        start = num["overrides"].get(ilvl, info["start"])
        return start, info["numFmt"], info["lvlText"]

    def render(self, num_id: str, ilvl: int) -> str | None:
        """推进 (num_id, ilvl) 的计数器一次，返回这次应该渲染出的编号文字。"""
        info = self._level_info(num_id, ilvl)
        if info is None:
            return None
        start, num_fmt, lvl_text = info

        counters = self._counters.setdefault(num_id, {})
        rendered = self._rendered.setdefault(num_id, set())
        for shallower in range(ilvl):
            if shallower not in counters:
                shallower_info = self._level_info(num_id, shallower)
                counters[shallower] = shallower_info[0] if shallower_info else 1
        for deeper in [lv for lv in counters if lv > ilvl]:
            del counters[deeper]
            # 深层级被重置掉：下次它再出现时要从 start 重新数（跟着 `counters`
            # 一起清，否则会被当成"已经渲染过"）。
            rendered.discard(deeper)

        if ilvl in counters and ilvl not in rendered:
            # 这个级别只是**被垫过**（为了让 `%1` 这种占位符有值），一次都没真
            # 渲染过：第一次就应该是 start，不能 +1。
            # 实测过的毛病：先出现"（一）"这种二级、再出现第一条一级标题时，
            # 一级会被渲染成 2（`counters[shallower] = start` 那一步把它垫成 1，
            # 后面 `+1` 就成了 2）。
            counters[ilvl] = start
        else:
            counters[ilvl] = counters.get(ilvl, start - 1) + 1
        rendered.add(ilvl)

        rendered = _format_number(num_fmt, counters[ilvl])
        if rendered is None:
            return None

        text = lvl_text
        for lv in range(ilvl + 1):
            placeholder = f"%{lv + 1}"
            if placeholder not in text:
                continue
            value = counters.get(lv, start)
            lv_info = self._level_info(num_id, lv)
            lv_fmt = lv_info[1] if lv_info else num_fmt
            formatted = _format_number(lv_fmt, value)
            text = text.replace(placeholder, formatted if formatted is not None else str(value))
        return text


def load_numbering_scheme(document) -> NumberingScheme:
    root = None
    if RT is not None:
        try:
            part = document.part.part_related_by(RT.NUMBERING)
            root = part.element
        except (KeyError, InvalidXmlError):
            root = None
    return NumberingScheme(root)
