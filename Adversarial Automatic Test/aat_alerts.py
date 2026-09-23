r"""告警层对拍 v1：**表里写的告警，代码里有没有**（文案层面）。

为什么先做文案：告警是用户唯一直接读到的东西，而"文案漂移"在这个项目里反复发生
（两套编号、文案改来改去、表与代码各写各的）。行为对拍（造文档去触发每一条告警）
更贵，放第二轮。

做法：
  * **表侧** ← 直接从窗口正文里解析两张表（复用 `rules_parser.find_tables`）：
      - 3.7 自动告警（A×10）：拿「告警名称」+「警告内容」（形如 `名称：概要`）；
      - 6.2 其他告警（M×17）：同上（这些不在 `alerts.py` 里，散在 GUI/备份/预设）。
  * **代码侧** ← AST 扫 `essay_tool/engine/alerts.py` 里所有 `self.add(...)` 调用，
    取出第 2、4 个实参（错误名称 / 错误概要）的字面量。
  * 对拍：按名称配对，报三类 —— ✓对上 / ✘表里有、代码里没有 / ✘代码里有、表里没写。
    概要按"去掉 `{占位}` 之后的骨架"比，避免 `{部件名}` 与 `{part}` 这种命名差异误报。
  * 格式体检（`Alert.line` 的约定）：概要里**不该**出现位置（“在段落”）、
    **不该**以句号/逗号结尾 —— 位置与句末标点由 `line` 统一拼。
  * M 表先做粗对账：那 17 条的原文能不能在 `essay_tool/` 里找到（找不到就要人工看）。

用法：
    .venv\Scripts\python.exe "Adversarial Automatic Test\aat_alerts.py"
"""
from __future__ import annotations

import ast
import pathlib
import re
import sys

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import rules_parser as rp                                        # noqa: E402

ALERTS_PY = ROOT / "essay_tool" / "engine" / "alerts.py"
REPORT = HERE / "_out" / "alert_check.txt"
PLACEHOLDER = re.compile(r"\{[^{}]*\}")


def literal(node: ast.AST) -> str | None:
    """把 AST 节点还原成字符串骨架（f-string 里的表达式变成 `{}`）。"""
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        return node.value
    if isinstance(node, ast.JoinedStr):
        pieces = []
        for value in node.values:
            if isinstance(value, ast.Constant) and isinstance(value.value, str):
                pieces.append(value.value)
            else:
                pieces.append("{}")
        return "".join(pieces)
    if isinstance(node, ast.BinOp) and isinstance(node.op, ast.Mod):
        return literal(node.left)                              # "…%s…" % x
    if (isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)
            and node.func.attr == "format"):
        return literal(node.func.value)                        # "…{}…".format(x)
    return None


def code_alerts() -> list[dict]:
    """`alerts.py` 里所有 `self.add(kind, name, part, message, ...)` 的字面量。"""
    tree = ast.parse(ALERTS_PY.read_text(encoding="utf-8"))
    found: list[dict] = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        if not (isinstance(node.func, ast.Attribute) and node.func.attr == "add"):
            continue
        if len(node.args) < 4:
            continue
        name = literal(node.args[1])
        message = literal(node.args[3])
        if name is None:
            continue
        found.append({
            "line": node.lineno,
            "name": name,
            "message": message if message is not None else "（非字面量，读不出来）",
        })
    return found


def skeleton(text: str) -> str:
    """去掉占位符与空白，只留骨架——用于两边比对（`{部件名}` ≡ `{part}`）。"""
    return "".join(PLACEHOLDER.sub("{}", text).split())


def variants(summary: str) -> set[str]:
    """表里的占位符可以带“或”写法：`论文中存在{图片／表格}，但未识别到…`
    意思是有两条文案（一条说图片、一条说表格）。这里把所有取值展开。
    """
    match = re.search(r"\{([^{}]+)\}", summary)
    if match is None:
        return {skeleton(summary)}
    options = re.split(r"[／/]", match.group(1))
    if len(options) == 1:
        return {skeleton(summary)}
    out: set[str] = set()
    for option in options:
        out |= variants(summary[:match.start()] + option + summary[match.end():])
    return out


def rows_of(tables: list[dict], kind: str) -> list[dict]:
    """按表种取行（"A"=3.7 自动告警、"M"=6.2 其他告警）。

    不能按小标题里的“6.2”找：**“3.6.2 文档的识别”里也有“6.2”**，
    会先撞上识别表（2026-09-20 实测的 KeyError）。
    """
    for table in tables:
        if rp.classify(table) == kind:
            headers = table["headers"]
            return [dict(zip(headers, row["cells"])) | {"_id": rp.rule_id(
                rp.classify(table), row, headers)}
                for row in table["rows"]]
    return []


def main() -> int:
    text = rp.load_window_text()
    tables = rp.find_tables(text)
    a_rows = rows_of(tables, "A")
    m_rows = rows_of(tables, "M")
    assert a_rows and m_rows, "两张告警表没解析出来"

    adds = code_alerts()
    by_name: dict[str, list[dict]] = {}
    for item in adds:
        by_name.setdefault(item["name"], []).append(item)

    lines: list[str] = []
    lines.append(f"表侧：3.7 共 {len(a_rows)} 条、6.2 共 {len(m_rows)} 条")
    lines.append(f"代码侧：alerts.py 里 self.add(...) 共 {len(adds)} 处、"
                 f"去重后 {len(by_name)} 个错误名称：{'、'.join(sorted(by_name))}")
    lines.append("")

    # ---------------------------------------------------------- 3.7 ↔ alerts.py
    # 代码全文（动态名称要从这里找：`check_sequence()` 把名字当返回值交给
    # `self.add(...)`，AST 只看得到变量名）
    sources = {path: path.read_text(encoding="utf-8", errors="replace")
               for path in sorted((ROOT / "essay_tool").rglob("*.py"))}
    lines.append("=== 3.7 自动告警 ↔ essay_tool/engine/alerts.py ===")
    matched: set[str] = set()
    bad = 0
    for row in a_rows:
        rid = row["_id"]
        name = row["告警名称"].strip()
        content = rp.pretty(row["警告内容"]).strip()
        head, _, summary = content.partition("：")
        if head.strip() != name:
            lines.append(f"  ⚠️ {rid} 表里“告警名称”是「{name}」，"
                         f"“警告内容”却是「{head}：…」——两列自己不一致")
        hits = by_name.get(name, [])
        if not hits:
            # 名称可能是别的模块里拼好传进来的（不是字面量）⇒ 搜全文
            places = [path.relative_to(ROOT).as_posix()
                      for path, body in sources.items() if name in body]
            if places:
                # 名称在别处拼的，概要也在别处拼（如 `detect.check_sequence()`）——
                # 拿概要里第一段没有占位的文字去同一棵树里找。
                chunk = PLACEHOLDER.split(summary)[0].strip()
                where = [path.relative_to(ROOT).as_posix()
                         for path, body in sources.items() if chunk and chunk in body]
                if where:
                    lines.append(f"  ✓ {rid} {name}：名称与概要都由 "
                                 f"{where[0]} 拼好传给 alerts.py（动态传入，已核对存在）")
                else:
                    bad += 1
                    lines.append(f"  ✘ {rid} {name}：名称在 {'、'.join(places)} 里，"
                                 f"但概要先头「{chunk}」哪都找不到")
            else:
                bad += 1
                lines.append(f"  ✘ {rid} {name}：**整个 essay_tool 里都找不到这个名字**")
            continue
        matched.add(name)
        wanted = variants(summary)
        got = {skeleton(item["message"]) for item in hits}
        missing = wanted - got
        if missing:
            bad += 1
            lines.append(f"  ✘ {rid} {name}：表里有、代码里没有的文案")
            for want in sorted(missing):
                lines.append(f"      表：{want}")
            for item in hits:
                lines.append(f"      码：{item['message']}   （alerts.py:{item['line']}）")
            continue
        lines.append(f"  ✓ {rid} {name}"
                     f"（alerts.py:{hits[0]['line']}，{len(wanted)} 种写法都对上）")

    extra = sorted(set(by_name) - matched)
    if extra:
        bad += len(extra)
        lines.append(f"  ✘ 代码里有、3.7 表里没写：{'、'.join(extra)}")
        for name in extra:
            for item in by_name[name]:
                lines.append(f"      {name}：{item['message']}   （alerts.py:{item['line']}）")

    lines.append("")
    lines.append("=== 概要格式体检（Alert.line 的约定）===")
    for item in adds:
        problems = []
        if "在段落" in item["message"]:
            problems.append("概要里写了位置（“在段落”应由 line 统一拼）")
        if item["message"].rstrip().endswith(("。", "，", "；")):
            problems.append("概要以标点结尾（应由 line 收尾）")
        if problems:
            bad += 1
            lines.append(f"  ✘ {item['name']}：{'；'.join(problems)}"
                         f"   （alerts.py:{item['line']}）")
    lines.append("  （以上没列出来的都合规）")

    # ---------------------------------------------------------- 6.2 粗对账
    lines.append("")
    lines.append("=== 6.2 其他告警 ↔ essay_tool 全文（粗对账：文案找不找得到）===")
    lines.append("（这 17 条不在 alerts.py 里，散在 GUI/备份/预设那边；名称多半是表自己")
    lines.append("  的标题，代码里是拼出来的，所以**只有“文案片段”算判据**。）")
    for row in m_rows:
        rid = row["_id"]
        name = row["告警名称"].strip()
        content = rp.pretty(row["警告内容"]).strip()
        places = [path.relative_to(ROOT).as_posix() for path, body in sources.items()
                  if name in body]
        # 文案：把 {占位} 拆成几段，看有没有哪一段在代码里出现（至少要有骨架）
        pieces = [piece for piece in PLACEHOLDER.split(content) if len(piece.strip()) >= 4]
        hit_pieces = [piece for piece in pieces
                      if any(piece.strip() in body for body in sources.values())]
        miss = [piece.strip() for piece in pieces if piece not in hit_pieces]
        mark = "✓" if not miss else "✘"
        if mark == "✘":
            bad += 1
        lines.append(f"  {mark} {rid} {name}"
                     f"   文案片段 {len(hit_pieces)}/{len(pieces)}"
                     + (f"   未命中：{miss[:3]}" if miss else ""))
        if not places:
            lines.append("      （名称只出现在表里本身，属正常；代码里是拼出来的字符串）")

    lines.append("")
    lines.append("=== 行为对拍：20 份夹具各跑一遍，看实际报了哪些警 ===")
    import gen_pipeline as gp                                    # noqa: PLC0415
    from essay_tool.config.preset_io import load_preset          # noqa: PLC0415
    from essay_tool.engine import pipeline                       # noqa: PLC0415
    from essay_tool.engine.alerts import AlertCollector          # noqa: PLC0415

    seen: dict[str, set[str]] = {}
    first_place: dict[str, str] = {}
    format_bad: list[str] = []
    for fixture in gp.FIXTURES:
        path = gp._fixture_path(fixture)
        if path is None:
            lines.append(f"  ⚠️ 夹具缺失：{fixture}")
            continue
        collector = AlertCollector()
        pipeline.run(path, None, HERE / "_out" / "alert_run" / path.stem,
                     preset_config=load_preset(gp.PRESET), alerts=collector)
        collector.consolidate()
        for alert in collector.items:
            seen.setdefault(alert.name, set()).add(skeleton(alert.message))
            first_place.setdefault(alert.name, fixture)
            line = alert.line
            problems = []
            if not line.endswith("。"):
                problems.append("结尾不是句号")
            if "。。" in line or "，。" in line:
                problems.append("标点重复")
            # ⚠️ 只查**位置那一截**：概要里可能引了文档原文（原文里半角括号很正常）。
            position = line.split("，在段落", 1)[1] if "，在段落" in line else ""
            if position:
                if "(" in position or ")" in position:
                    problems.append("位置用了半角括号")
                if not position.endswith("）。"):
                    problems.append("位置没写成“在段落（段号）。”")
            if problems:
                format_bad.append(f"{fixture}：{line}（{'；'.join(problems)}）")

    table_names = {row["告警名称"].strip() for row in a_rows}
    covered = 0
    for row in a_rows:
        name = row["告警名称"].strip()
        if name in seen:
            covered += 1
            lines.append(f"  ✓ {row['_id']} {name}：触发过"
                         f"（{len(seen[name])} 种概要写法，首见 {first_place[name]}）")
        else:
            lines.append(f"  ○ {row['_id']} {name}：这 20 份夹具里**从没触发过**"
                         f"（要专门造文档才能验）")
    extra_names = sorted(set(seen) - table_names)
    if extra_names:
        bad += len(extra_names)
        lines.append(f"  ✘ 代码报了、3.7 表里没写：{'、'.join(extra_names)}")
        for name in extra_names:
            lines.append(f"      {name}（首见 {first_place[name]}）")
    lines.append(f"  覆盖：{covered}/{len(a_rows)} 条识别告警在夹具上触发过")
    if format_bad:
        bad += len(format_bad)
        lines.append("  告警整行格式不合规：")
        for item in format_bad[:10]:
            lines.append(f"      {item}")
    else:
        lines.append("  告警整行格式（`名称：概要，在段落（段号）。`）全部合规")

    lines.append("")
    lines.append(f"问题条数：{bad}")
    print("\n".join(lines))
    REPORT.write_text("\n".join(lines), encoding="utf-8")
    print()
    print(f"报告已写出：{REPORT.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
