# -*- coding: utf-8 -*-
r"""例外条款机械切分——**2026-09-20 按重写后的识别表重建**。

新表没有单独的「允许的例外」列：允许事项就写在「识别规则」格子里
（“允许半角括号，且不必成对使用”“允许‘摘要’两字间有半角或全角空格”…）。
所以本层的输入改成：**识别规则格里含“允许/不必/不要求/可以”的条款**，
逐条机械生成样本，跑“表 vs 引擎”对拍。

条款 → 样本的分类（关键词驱动，机械）：
  * 【/】/[/]        → 括号样本（不要求成对，单边不超过一个）
  * 括号             → 半角/全角括号、不成对写法
  * 空格             → 字间空格（半角/全角）与词尾空格
  * 大小写           → 大小写样本
  * 顿号+点号        → 分隔符替换样本（. ／ ． ／ 、）
  * 阿拉伯数字写法   → “1 标题”“1. 标题”“1.标题”三种接法（方案=点分）
  * 标点（且“不…”）  → 尾部标点样本
  * 长度/60          → 超长样本
  * 其余             → 登记“没分类”，不偷偷猜

方案（chinese/dotted）按条款文字取：含“阿拉伯数字序号层级/阿拉伯”走点分，
含“中文序号层级”走中文，其余默认中文。

用法（运行方式不变）：
    .venv\Scripts\python.exe "Adversarial Automatic Test\gen_exceptions.py"
    .venv\Scripts\python.exe "Adversarial Automatic Test\gen_exceptions.py" --only H2
"""
from __future__ import annotations

import concurrent.futures
import itertools
import pathlib
import re
import shutil
import sys
import time

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent
for entry in (str(ROOT), str(HERE)):
    if entry not in sys.path:
        sys.path.insert(0, entry)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import gen_matrix as gm                                          # noqa: E402
import rules_parser as rp                                        # noqa: E402

REPORT = HERE / "_out" / "gen-exceptions.txt"
DEFAULT_JOBS = 4

_QUOTED = re.compile(r"[“\"]([^”\"]+)[”\"]")
_BRACKETS = set("【】[]")
_CLAUSE_TRIGGERS = ("允许", "不必", "不要求", "可以")
# 尾部标点样本的标点表（条款只写“不许以标点结尾”，没列标点——用一张固定表扫）。
TAIL_PUNCT = ["。", "；", "，", ".", "！", "？", "：", "…", "”", "》", "）"]

CN_SKELETON = {"一级": "一、标题文字", "二级": "（一）标题文字",
               "三级": "1. 标题文字", "四级": "（1）标题文字"}
AR_NUMBER = {"一级": "1", "二级": "1.1", "三级": "1.1.1", "四级": "1.1.1.1"}


def clauses_of(cell: str) -> list[str]:
    """把一个格子的文字按「；」「。」换行切成条款，并剥掉“对某某层级：”前缀。"""
    parts = re.split(r"[；。\n]+", cell)
    out: list[str] = []
    for part in parts:
        part = re.sub(r"^对?[^：]{0,12}层级[:：]", "", part.strip()).strip()
        if part:
            out.append(part)
    return out


def is_allowance(clause: str) -> bool:
    return any(trigger in clause for trigger in _CLAUSE_TRIGGERS)


def quoted_tokens(text: str) -> list[str]:
    return [token for token in _QUOTED.findall(text) if token.strip()]


def level_of(name: str) -> str | None:
    for level in CN_SKELETON:
        if level in name:
            return level
    return None


def base_token(clause: str, name: str) -> str:
    """被“允许”的那截文字：条款引号 → 识别对象名。"""
    word_like = [t for t in quoted_tokens(clause) if not (set(t) <= _BRACKETS)]
    if word_like:
        return word_like[0]
    return name


def skeleton_of(name: str) -> str | None:
    """标题骨架（一、标题文字／（一）标题文字…）或题注骨架（图 1 标题文字）。"""
    level = level_of(name)
    if level:
        return CN_SKELETON[level]
    if name == "图题":
        return "图 1 标题文字"
    if name == "表题":
        return "表 1 标题文字"
    return None


def scheme_of(clause: str, name: str) -> str:
    if "阿拉伯" in clause or "阿拉伯" in name:
        return "dotted"
    return "chinese"


def samples_for(rule: str, name: str, clause: str) -> tuple[list[str], str, str]:
    """`(样本, 类名, 方案)`；生成不了给 `([], "没分类", 方案)`。"""
    scheme = scheme_of(clause, name)
    level = level_of(name)
    skeleton = skeleton_of(name)
    base = base_token(clause, name)
    text = skeleton or base

    # 阿拉伯数字写法：1 标题／1. 标题／1.标题（按级别取编号段）
    if level and scheme == "dotted" and ("阿拉伯数字" in clause or "点号" in clause
                                         or "空格" in clause or "汉字" in clause):
        number = AR_NUMBER[level]
        return [f"{number} 标题文字", f"{number}. 标题文字",
                f"{number}.标题文字"], "阿拉伯数字写法", scheme
    if any(ch in clause for ch in _BRACKETS):
        result = [front + text + back
                  for front, back in itertools.product(("", "【", "["), ("", "】", "]"))
                  if front or back]
        return sorted(set(result)), "括号", scheme
    if "括号" in clause:
        # “半角括号、不必成对”：全角/半角 × 成对/缺一边，都试
        pairs = [("（", "）"), ("(", ")"), ("（", ")"), ("(", "）")]
        stripped = text.strip("（）()")
        forms = [f"（{stripped}）", f"({stripped})", f"（{stripped})", f"({stripped}）",
                 f"（{stripped}", f"({stripped}", f"{stripped}）", f"{stripped})"]
        result = {form for form in forms if form != text}
        return sorted(result), "括号（半角/全角/不成对）", scheme
    if "空格" in clause:
        if "字间" in clause:
            return [" ".join(text), "　".join(text)], "字间空格", scheme
        return [text + " ", text + "　"], "词尾空格", scheme
    if "大小写" in clause or "不敏感" in clause:
        if text.isascii():
            return sorted({text.upper(), text.lower()}), "大小写", scheme
        # 英文标签的基词另有写法（ABSTRACT 等），条款没说清 → 登记
        return [], "没分类", scheme
    if "点号" in clause and ("顿号" in clause or "半角" in clause or "全角" in clause):
        marks = ["."]
        if "全角" in clause:
            marks.append("．")
        if "顿号" in clause:
            marks.append("、")
        pivot = next((mark for mark in ".．、" if mark in text), None)
        if pivot is None or len(marks) < 2:
            return [], "没分类", scheme
        return sorted({text.replace(pivot, mark, 1) for mark in marks} - {text}), \
            "点号/顿号", scheme
    if "标点" in clause and ("不" in clause):
        if skeleton is None:
            return [], "没分类", scheme
        return [skeleton + punct for punct in TAIL_PUNCT], "尾部标点", scheme
    if ("长度" in clause or "60" in clause) and skeleton is not None:
        prefix = skeleton[: len(skeleton) - len("标题文字")]
        return sorted({prefix + "字" * 60, prefix + "字" * 61}), "超长", scheme
    return [], "没分类", scheme


def main(argv: list[str]) -> int:
    only = argv[argv.index("--only") + 1] if "--only" in argv else None
    jobs = int(argv[argv.index("--jobs") + 1]) if "--jobs" in argv else DEFAULT_JOBS
    limit = int(argv[argv.index("--limit") + 1]) if "--limit" in argv else None

    # 现算，不读 _out/rule_tables.json（那是产物、会被清理；2026-09-20 实拍 FileNotFoundError）
    tables = rp.parse_tables()
    plan: list[tuple[str, str, str, str]] = []          # (说明, 文本, 规则, 方案)
    unclassified: list[str] = []
    seen: set[tuple[str, str]] = set()
    for table in tables:
        for row in table["rows"]:
            cells = row["cells"]
            if "识别规则" not in cells:
                continue
            rule = row["id"]
            if only and not rule.startswith(only):
                continue
            name = (cells.get("识别对象") or "").strip()
            for clause in clauses_of(cells["识别规则"]):
                if not is_allowance(clause):
                    continue
                samples, kind, scheme = samples_for(rule, name, clause)
                if not samples:
                    unclassified.append(f"  {rule}  [{kind}] {clause}")
                    continue
                for text in samples:
                    key = (rule, text)
                    if key in seen:
                        continue
                    seen.add(key)
                    plan.append((f"{rule}｜{kind}｜{clause[:20]}", text, rule, scheme))
    if limit is not None:
        plan = plan[:limit]

    payloads = [(gm.CONTEXT + [text], 2, scheme)
                for _name, text, _rule, scheme in plan]
    total = {"✓一致": 0, "✘分歧": 0, "？未判定": 0}
    diffs: list[str] = []
    undecided: list[str] = []
    started = time.perf_counter()
    if jobs > 1:
        with concurrent.futures.ProcessPoolExecutor(max_workers=jobs) as pool:
            outcomes = list(pool.map(gm._run_payload, payloads))
    else:
        outcomes = [gm.run_case(lines, probe, scheme)
                    for lines, probe, scheme in payloads]

    for (name, text, rule, _scheme), (conclusion, mine, theirs, got_rule) in zip(plan, outcomes):
        total[conclusion] += 1
        line = (f"  {conclusion}  {name}\n"
                f"        样本={text!r}\n"
                f"        表={mine}  引擎={theirs}  规则={got_rule or rule}")
        if conclusion == "✘分歧":
            diffs.append(line)
        elif conclusion == "？未判定":
            undecided.append(line)

    ran = sum(total.values())
    elapsed = time.perf_counter() - started
    print(f"例外条款切出 {ran} 条样本（{jobs} 进程，{elapsed:.1f} 秒）："
          + "、".join(f"{key} {value}" for key, value in total.items()))
    if unclassified:
        print(f"没分类的条款 {len(unclassified)} 条（留给人工挑）：")
        for item in unclassified[:20]:
            print(item)
        if len(unclassified) > 20:
            print(f"  …（还有 {len(unclassified) - 20} 条）")
    for title, items in (("分歧（表与引擎不一致）", diffs),
                         ("表里没写（登记，不当 bug）", undecided)):
        if not items:
            continue
        print()
        print("=" * 74)
        print(f"{title}：{len(items)} 条")
        print("=" * 74)
        for item in items[:30]:
            print(item)
        if len(items) > 30:
            print(f"  …（还有 {len(items) - 30} 条，见 _out/gen-exceptions.txt）")

    report = "\n".join(unclassified + diffs + undecided)
    REPORT.write_text(report, encoding="utf-8")
    print()
    print(f"明细已写出：{REPORT.relative_to(ROOT)}")
    shutil.rmtree(gm.WORK, ignore_errors=True)
    return 1 if diffs else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
