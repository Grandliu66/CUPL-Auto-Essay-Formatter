# -*- coding: utf-8 -*-
"""只读代码审计：静态扫描 + 交叉引用统计，不导入、不修改任何被测文件。

扫描内容：
  * 每个模块的语法是否能解析（用 ast.parse，不写 __pycache__）；
  * 顶层的类/函数、类里的方法，以及它们被谁引用
    （生产代码 essay_tool / 只有 tools 用 / 谁都没用）；
  * 从未被生产代码 import 的模块；
  * import 了却没用到的名字；
  * 空的异常处理（`except ...: pass` / 只 continue）、裸 except；
  * 过时标记（TODO/FIXME/过时/暂不/待定/未实现…）与可疑注释；
  * 生产代码里的 print（应走日志/状态栏）；
  * 可变默认参数、超长行、缺模块 docstring；
  * 归一化后完全相同的函数体（重复实现）。

产物：work\\_audit_readonly.md（Markdown 报告）
用法：.venv\\Scripts\\python.exe tools\\audit_readonly.py
"""
from __future__ import annotations

import ast
import hashlib
import pathlib
import re
import sys
from collections import defaultdict

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = pathlib.Path(__file__).resolve().parents[1]
OUT = ROOT / "work" / "_audit_readonly.md"
PROD_DIRS = [ROOT / "essay_tool"]
DEV_DIR = ROOT / "tools"
SKIP_PARTS = {"__pycache__", "work", "backup", "_obsolete_tools"}

MARKERS = ("TODO", "FIXME", "XXX", "过时", "暂不", "待定", "待补", "未实现",
           "写死", "临时", "权宜", "凑", "hack", "HACK")
SUSPECT_EXCEPT = re.compile(r"except\s+(Exception|BaseException)?\s*:")

report: list[str] = []


def say(text: str = "") -> None:
    report.append(text)


def py_files(*dirs: pathlib.Path) -> list[pathlib.Path]:
    out: list[pathlib.Path] = []
    for base in dirs:
        for path in sorted(base.rglob("*.py")):
            if SKIP_PARTS & set(path.parts):
                continue
            out.append(path)
    return out


def rel(path: pathlib.Path) -> str:
    return path.relative_to(ROOT).as_posix()


class Func:
    def __init__(self, path: pathlib.Path, node: ast.AST, qual: str, lines: str):
        self.path = path
        self.node = node
        self.qual = qual                       # 形如 "Cleanup.strip_run_direct_formatting"
        self.name = qual.split(".")[-1]
        self.line = node.lineno
        self.body_src = ast.get_source_segment(lines, node) or ""
        self.body_hash = hashlib.sha1(
            re.sub(r"\s+", "", self.body_src).encode("utf-8")).hexdigest()


def scan() -> None:
    prod_files = py_files(*PROD_DIRS)
    dev_files = py_files(DEV_DIR)
    all_files = prod_files + dev_files

    tree_of: dict[pathlib.Path, ast.Module] = {}
    text_of: dict[pathlib.Path, str] = {}
    parse_errors: list[tuple[pathlib.Path, str]] = []
    for path in all_files:
        raw = path.read_bytes()
        text = raw.decode("utf-8-sig")      # 有些文件带 BOM，用 utf-8 解会多出一个字符
        text_of[path] = text
        try:
            tree_of[path] = ast.parse(text)
        except SyntaxError as exc:
            parse_errors.append((path, f"{exc.lineno}: {exc.msg}"))

    say("# 只读代码审计报告")
    say()
    say(f"- 生产代码：{len(prod_files)} 个文件；开发脚本：{len(dev_files)} 个文件")
    say(f"- 语法解析失败：{len(parse_errors)} 个")
    for path, err in parse_errors:
        say(f"  - ★ {rel(path)}：{err}")
    say()
    if parse_errors:
        print("★ 有文件解析不过，先看这些：")
        for path, err in parse_errors:
            print(f"   {rel(path)}：{err}")

    # ---------------------------------------------------------------- 定义与引用
    funcs: list[Func] = []
    classes: dict[str, tuple[pathlib.Path, int]] = {}
    module_names: dict[str, pathlib.Path] = {}
    for path in all_files:
        tree = tree_of.get(path)
        if tree is None:
            continue
        lines = text_of[path]
        module = rel(path)
        module_names[module] = path
        for node in tree.body:
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                funcs.append(Func(path, node, node.name, lines))
            elif isinstance(node, ast.ClassDef):
                classes[f"{node.name}"] = (path, node.lineno)
                for sub in node.body:
                    if isinstance(sub, (ast.FunctionDef, ast.AsyncFunctionDef)):
                        funcs.append(Func(path, sub, f"{node.name}.{sub.name}", lines))

    prod_text = "\n".join(text_of[p] for p in prod_files)
    dev_text = "\n".join(text_of[p] for p in dev_files)

    def refs(name: str, own: pathlib.Path) -> tuple[int, int]:
        """这个名字在生产代码里出现几次、在开发脚本里出现几次。

        **包含定义它自己的那个文件**（同文件里的私有小函数就是在自己文件里用的），
        再把定义处那一次减掉。
        """
        pat = re.compile(rf"\b{re.escape(name)}\b")
        prod = sum(len(pat.findall(text_of[p])) for p in prod_files)
        dev = sum(len(pat.findall(text_of[p])) for p in dev_files)
        return max(0, prod - 1), dev

    never: list[Func] = []
    only_dev: list[Func] = []
    for fn in funcs:
        if fn.name.startswith("__"):
            continue
        prod, dev = refs(fn.name, fn.path)
        if prod == 0 and dev == 0:
            never.append(fn)
        elif prod == 0:
            only_dev.append(fn)

    say("## 一、可能没人用的函数（只按名字统计，同名会互相掩盖）")
    say()
    say(f"生产代码里一次都没提到、连开发脚本也没提的：{len(never)} 个")
    for fn in sorted(never, key=lambda f: (rel(f.path), f.line)):
        kind = "私有" if fn.name.startswith("_") else "公开"
        say(f"- {kind} `{fn.qual}`  （{rel(fn.path)}:{fn.line}）")
    say()
    say(f"只有开发脚本（tools/）在用的：{len(only_dev)} 个")
    for fn in sorted(only_dev, key=lambda f: (rel(f.path), f.line)):
        say(f"- `{fn.qual}`  （{rel(fn.path)}:{fn.line}）")
    say()

    # ---------------------------------------------------------------- 模块引用
    say("## 二、从未被生产代码 import 的模块")
    say()
    prod_imports = set()
    for path in prod_files:
        tree = tree_of.get(path)
        if tree is None:
            continue
        for node in ast.walk(tree):
            if isinstance(node, ast.ImportFrom) and node.module:
                prod_imports.add(node.module)
            elif isinstance(node, ast.Import):
                for alias in node.names:
                    prod_imports.add(alias.name)
    dev_imports = set()
    for path in dev_files:
        tree = tree_of.get(path)
        if tree is None:
            continue
        for node in ast.walk(tree):
            if isinstance(node, ast.ImportFrom) and node.module:
                dev_imports.add(node.module)
            elif isinstance(node, ast.Import):
                for alias in node.names:
                    dev_imports.add(alias.name)

    for module, path in sorted(module_names.items()):
        if not module.startswith("essay_tool/"):
            continue
        dotted = module[:-3].replace("/", ".")
        if dotted.endswith(".__init__"):
            continue
        in_prod = any(dotted == imp or dotted.startswith(imp + ".") or imp.startswith(dotted)
                      for imp in prod_imports)
        in_dev = any(dotted == imp or dotted.startswith(imp + ".") or imp.startswith(dotted)
                     for imp in dev_imports)
        if not in_prod:
            say(f"- `{module}`：生产代码没用"
                + ("（只有 tools/ 在用）" if in_dev else "（谁都没用）"))
    say()

    # ---------------------------------------------------------------- 未用 import
    say("## 三、import 了却没在文件里用到的名字")
    say()
    for path in all_files:
        tree = tree_of.get(path)
        if tree is None:
            continue
        text = text_of[path]
        body_lines = text.splitlines()
        for node in ast.walk(tree):
            names: list[tuple[str, int]] = []
            if isinstance(node, ast.Import):
                names = [(a.asname or a.name.split(".")[0], node.lineno) for a in node.names]
            elif isinstance(node, ast.ImportFrom):
                names = [(a.asname or a.name, node.lineno) for a in node.names
                         if a.name != "*"]
            for name, line in names:
                pattern = re.compile(rf"\b{re.escape(name)}\b")
                hits = sum(1 for i, ln in enumerate(body_lines, start=1)
                           if i != line and pattern.search(ln))
                if hits == 0:
                    say(f"- {rel(path)}:{line}  `{name}`")
    say()

    # ---------------------------------------------------------------- 异常处理
    say("## 四、异常处理里的可疑写法")
    say()
    for path in all_files:
        tree = tree_of.get(path)
        if tree is None:
            continue
        for node in ast.walk(tree):
            if not isinstance(node, ast.Try):
                continue
            for handler in node.handlers:
                body = handler.body
                only_pass = all(isinstance(s, ast.Pass) for s in body)
                only_continue = all(isinstance(s, ast.Continue) for s in body)
                if handler.type is None:
                    say(f"- ★ 裸 except：{rel(path)}:{handler.lineno}")
                elif only_pass or only_continue:
                    kind = "pass" if only_pass else "continue"
                    say(f"- 静默吞掉（{kind}）：{rel(path)}:{handler.lineno}"
                        f"  `{ast.unparse(handler.type)}`")
    say()

    # ---------------------------------------------------------------- 标记与味道
    say("## 五、过时标记与可疑注释")
    say()
    for path in all_files:
        for i, line in enumerate(text_of[path].splitlines(), start=1):
            stripped = line.strip()
            if not stripped.startswith("#") and '"""' not in stripped:
                continue
            for marker in MARKERS:
                if marker in stripped:
                    say(f"- {rel(path)}:{i}  {stripped[:110]}")
                    break
    say()

    say("## 六、生产代码里的 print")
    say()
    for path in prod_files:
        for i, line in enumerate(text_of[path].splitlines(), start=1):
            if re.match(r"^\s*print\(", line):
                say(f"- {rel(path)}:{i}  {line.strip()[:100]}")
    say()

    say("## 七、可变默认参数 / 超长行 / 缺模块 docstring")
    say()
    for path in all_files:
        tree = tree_of.get(path)
        if tree is None:
            continue
        if ast.get_docstring(tree) is None:
            say(f"- 缺模块 docstring：{rel(path)}")
        for node in ast.walk(tree):
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                for default in list(node.args.defaults) + [
                        d for d in node.args.kw_defaults if d is not None]:
                    if isinstance(default, (ast.List, ast.Dict, ast.Set)):
                        say(f"- ★ 可变默认参数：{rel(path)}:{node.lineno} `{node.name}`")
        for i, line in enumerate(text_of[path].splitlines(), start=1):
            if len(line) > 100:
                say(f"- 超长行（{len(line)} 字）：{rel(path)}:{i}")
    say()

    # ---------------------------------------------------------------- 重复实现
    say("## 八、函数体完全相同的重复实现")
    say()
    buckets: dict[str, list[Func]] = defaultdict(list)
    for fn in funcs:
        if len(fn.body_src) > 60:
            buckets[fn.body_hash].append(fn)
    for group in buckets.values():
        if len(group) > 1:
            names = "、".join(f"`{f.qual}`({rel(f.path)}:{f.line})" for f in group)
            say(f"- {names}")
    say()

    # ---------------------------------------------------------------- 规模
    say("## 九、带 BOM 的文件（UTF-8 BOM，对 .py 无害但风格不一致，对 .json 会直接报错）")
    say()
    bom_hits: list[str] = []
    for path in sorted(ROOT.rglob("*")):
        if not path.is_file() or SKIP_PARTS & set(path.parts):
            continue
        if path.suffix.lower() not in (".py", ".json", ".md", ".txt", ".csv"):
            continue
        try:
            if path.read_bytes()[:3] == b"\xef\xbb\xbf":
                bom_hits.append(rel(path))
        except OSError:
            continue
    for item in bom_hits:
        say(f"- {item}")
    say(f"共 {len(bom_hits)} 个。")
    say()
    say("## 十、规模")
    say()
    say("| 文件 | 行数 |")
    say("|---|---|")
    for path in sorted(prod_files, key=lambda p: -len(text_of[p].splitlines())):
        say(f"| {rel(path)} | {len(text_of[path].splitlines())} |")
    say()
    total = sum(len(text_of[p].splitlines()) for p in prod_files)
    say(f"生产代码共 {total} 行。")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text("\n".join(report) + "\n", encoding="utf-8")
    print("\n".join(report[:0]))          # 报告太长，不往终端刷
    print(f"报告：{OUT.relative_to(ROOT)}（{len(report)} 行）")


if __name__ == "__main__":
    scan()
