"""只读审计（第二遍）：专盯"可疑写法"这一类，**不修改任何文件**。

跟 `audit_readonly.py` 分工：那个扫"有没有人用、有没有重复、有没有 BOM"，
这个扫那些"语法上没错、但十有八九会咬人"的写法：

* 可变默认参数（`def f(x=[])`）——同一个对象被所有调用共享；
* 光秃秃的 `except:`（连 KeyboardInterrupt 都吞）；
* 打开文本文件不写 encoding —— 在别的机器上会按 GBK/ANSI 读，中文直接乱；
* `return`/`raise`/`continue`/`break` 之后还写了语句（永远跑不到）；
* `is` 拿去比字符串/数字字面量（CPython 里有时碰巧成立，有时不成立）；
* `eval` / `exec` / `pickle.loads` 之类的口子；
* `assert` 出现在生产代码里（`-O` 之下会被整条删掉）；
* TODO / FIXME / XXX / HACK / 待办 之类的占位；
* `type: ignore` / `noqa` 的分布（宽 except 有多少处）。

结果写进 `work\\_audit_readonly2.md`。
"""

from __future__ import annotations

import ast
import pathlib
import re

ROOT = pathlib.Path(__file__).resolve().parents[1]
PROD_DIR = ROOT / "essay_tool"
OUT = ROOT / "work" / "_audit_readonly2.md"

STMT_WORDS = ("return", "raise", "continue", "break")


def read(path: pathlib.Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def rel(path: pathlib.Path) -> str:
    return path.relative_to(ROOT).as_posix()


def py_files(base: pathlib.Path) -> list[pathlib.Path]:
    return sorted(p for p in base.rglob("*.py") if "__pycache__" not in p.parts)


def const_source(node: ast.AST) -> str:
    try:
        return ast.unparse(node)
    except Exception:                     # noqa: BLE001 —— 反解不出来就不写源码
        return "?"


def walk(node: ast.AST):
    """按源码顺序遍历所有节点（ast.walk 的顺序不保证，这里自己走）。"""
    for child in ast.iter_child_nodes(node):
        yield child
        yield from walk(child)


def scan(path: pathlib.Path, base: pathlib.Path) -> dict[str, list[str]]:
    text = read(path)
    found: dict[str, list[str]] = {}

    def hit(kind: str, lineno: int, msg: str) -> None:
        found.setdefault(kind, []).append(f"- `{rel(path)}:{lineno}`  {msg}")

    try:
        tree = ast.parse(text)
    except SyntaxError as exc:
        hit("解析失败", exc.lineno or 0, str(exc))
        return found

    for node in walk(tree):
        # ---- 可变默认参数 ----
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            for default in list(node.args.defaults) + [
                d for d in node.args.kw_defaults if d is not None
            ]:
                if isinstance(default, (ast.List, ast.Dict, ast.Set)):
                    hit("可变默认参数", node.lineno,
                        f"`{node.name}` 的默认值是 {const_source(default)}")
                elif isinstance(default, ast.Call):
                    name = const_source(default.func)
                    if name.split(".")[-1] in ("list", "dict", "set", "defaultdict", "OrderedDict"):
                        hit("可变默认参数", node.lineno,
                            f"`{node.name}` 的默认值是 {const_source(default)}")

        # ---- 光秃秃的 except ----
        if isinstance(node, ast.ExceptHandler) and node.type is None:
            hit("裸 except", node.lineno, "`except:`（连 Ctrl-C 都吞）")

        # ---- is 比字面量 ----
        if isinstance(node, ast.Compare):
            for op, right in zip(node.ops, node.comparators):
                if isinstance(op, (ast.Is, ast.IsNot)) and isinstance(
                        right, (ast.Str, ast.Constant)):
                    value = getattr(right, "value", None)
                    if isinstance(value, (str, int, float, bytes)):
                        hit("`is` 比字面量", node.lineno,
                            f"`{const_source(node)}`")

        # ---- 死语句：return/raise/continue/break 之后还有语句 ----
        for field in ("body", "orelse", "finalbody"):
            block = getattr(node, field, None)
            if not isinstance(block, list):
                continue
            for index, stmt in enumerate(block):
                if index + 1 >= len(block) or not isinstance(stmt, ast.stmt):
                    continue
                if isinstance(stmt, (ast.Return, ast.Raise, ast.Continue, ast.Break)):
                    nxt = block[index + 1]
                    hit("死语句", getattr(nxt, "lineno", 0),
                        f"`{STMT_WORDS[0]}` 之类之后还有 `{type(nxt).__name__}`"
                        f"（第 {stmt.lineno} 行已 {type(stmt).__name__.lower()}）")

        # ---- 危险调用 ----
        if isinstance(node, ast.Call):
            name = const_source(node.func)
            tail = name.split(".")[-1]
            if tail in ("eval", "exec") and not name.startswith("ast."):
                hit("危险调用", node.lineno, f"`{name}(…)`")
            if name in ("pickle.loads", "pickle.load", "marshal.loads"):
                hit("危险调用", node.lineno, f"`{name}(…)`")

            # ---- 文本 I/O 没写 encoding ----
            if isinstance(node.func, ast.Name) and node.func.id == "open":
                keywords = {kw.arg for kw in node.keywords}
                mode = ""
                if len(node.args) >= 2 and isinstance(node.args[1], ast.Constant):
                    mode = node.args[1].value or ""
                for kw in node.keywords:
                    if kw.arg == "mode" and isinstance(kw.value, ast.Constant):
                        mode = kw.value.value or ""
                if "b" not in mode and "encoding" not in keywords:
                    hit("缺 encoding", node.lineno, f"`open(…)` 没写 encoding（mode={mode!r}）")
            if tail in ("read_text", "write_text"):
                keywords = {kw.arg for kw in node.keywords}
                if "encoding" not in keywords:
                    hit("缺 encoding", node.lineno, f"`{name}(…)` 没写 encoding")

        # ---- assert 出现在生产代码里 ----
        if isinstance(node, ast.Assert):
            hit("assert", node.lineno, "`-O` 跑起来这条会被整条删掉")

    # ---- 文本层面的标记 ----
    for lineno, line in enumerate(text.splitlines(), start=1):
        if re.search(r"\b(TODO|FIXME|XXX|HACK)\b", line):
            hit("遗留标记", lineno, line.strip()[:90])
        if "type: ignore" in line:
            hit("type: ignore", lineno, line.strip()[:90])
        if "noqa" in line:
            hit("noqa", lineno, line.strip()[:90])
    return found


def main() -> None:
    titles = {
        "解析失败": "一、语法解析失败",
        "可变默认参数": "二、可变默认参数",
        "裸 except": "三、裸 except",
        "`is` 比字面量": "四、`is` 比字面量",
        "死语句": "五、永远跑不到的语句",
        "危险调用": "六、eval/exec/pickle",
        "缺 encoding": "七、文本 I/O 没写 encoding",
        "assert": "八、生产代码里的 assert",
        "遗留标记": "九、TODO/FIXME 之类",
        "type: ignore": "十、type: ignore",
        "noqa": "十一、noqa（宽 except 的标记）",
    }
    all_found: dict[str, list[str]] = {}
    for path in py_files(PROD_DIR) + py_files(ROOT / "tools"):
        for kind, items in scan(path, PROD_DIR).items():
            all_found.setdefault(kind, []).extend(items)

    lines = ["# 只读代码审计报告（第二遍：可疑写法）", ""]
    for kind, title in titles.items():
        items = all_found.get(kind, [])
        lines.append(f"## {title}")
        lines.append("")
        if not items:
            lines.append("（无）")
        else:
            # 解析失败最要紧，排前面；其余按文件+行号
            def sort_key(item: str):
                m = re.match(r"- `([^`]+):(\d+)`", item)
                return (m.group(1), int(m.group(2))) if m else (item, 0)

            lines.extend(sorted(items, key=sort_key))
        lines.append("")
    OUT.write_text("\n".join(lines), encoding="utf-8")
    print(f"报告：{rel(OUT)}（{len(lines)} 行）")


if __name__ == "__main__":
    main()
