r"""夹具生成器 v1（识别侧穷举）——**2026-09-20 按重写后的 3.6.2 识别表重建**。

分工（用户 2026-09-17 明确，沿用）：
  * **元（维度与取值）由人构思**——工具不猜"可能出什么状况"；
  * 工具只做两件事：把维度**成对展开**成具体段落、逐条跑，然后分类报告。

新表有两套标题编号方案（中文序号层级 / 阿拉伯数字序号层级），写法判据不同，
所以本层的"元"也分两套，各自穷举、各自对拍——**期望侧与引擎侧都按同一个
方案判**（引擎的识别入口现在接得上方案参数）。

三类结论（与重建前同）：
  ✓ 一致   —— 表的判据与引擎的识别结果相同
  ✘ 分歧   —— 两者不同（要么表错、要么代码错，报告里带规则 ID）
  ？未判定 —— 表里没写这种情况（登记，不当 bug）

用法（运行方式不变）：
    .venv\Scripts\python.exe "Adversarial Automatic Test\gen_matrix.py"
    .venv\Scripts\python.exe "Adversarial Automatic Test\gen_matrix.py" --only 尾部
"""
from __future__ import annotations

import concurrent.futures
import itertools
import os
import pathlib
import shutil
import sys
import time

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent
for entry in (str(ROOT), str(HERE)):
    if entry not in sys.path:
        sys.path.insert(0, entry)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import aat_expect as ex                                         # noqa: E402
from aat_docx import read_docx                                  # noqa: E402
from docx import Document                                       # noqa: E402
from aat_engine_api import engine_roles as _engine_roles        # noqa: E402

WORK = HERE / "_out" / "gen"
# 默认只吃 4 个核（用户 2026-09-17：“只要不吃满所有就可以”，本机 16 核 32 线程）。
DEFAULT_JOBS = 4

TEXT = "标题文字"
# 最小上下文：题目 + 空段 + 被测段。题目挡掉 T1（否则被测段会被认成题目），
# 空段挡住 T2（副标题要求紧邻、不隔空段）——被测段因此只会按标题判据被判。
CONTEXT = ["论测试文档里标题的写法", ""]


def case_path() -> pathlib.Path:
    """每个进程一个自己的临时文件——共用一个文件名在多进程下会互相撞。"""
    return WORK / f"p{os.getpid()}" / "case.docx"


# ---------------------------------------------------------------- 元·中文序号层级
CN_DIMENSIONS: dict[str, list[tuple[str, str]]] = {
    # 编号写法：一级（顿号）、二级（括号）、三级（阿拉伯点号/顿号/全角点）、四级（阿拉伯括号）
    "编号写法": [
        ("顿号", "{n}、{t}"), ("括号", "（{n}）{gap}{t}"), ("半角括号", "({n}){gap}{t}"),
        ("左括号无右", "（{n}{gap}{t}"), ("右括号无左", "{n}){gap}{t}"),
        ("阿拉伯点号", "{n}. {t}"), ("阿拉伯顿号", "{n}、{t}"),
        ("阿拉伯全角点", "{n}．{t}"), ("阿拉伯括号", "（{n}）{t}"),
    ],
    # 数字：表只说了“百、千不能作首字”“廿、卅不被承认”——其余常见写法都试
    "数字": [
        ("一", "一"), ("二", "二"), ("十", "十"), ("十一", "十一"), ("二十", "二十"),
        ("廿一", "廿一"), ("卅一", "卅一"), ("百一", "百一"), ("千一", "千一"),
        ("两", "两"), ("零", "零"), ("〇", "〇"), ("万", "万"), ("万一", "万一"),
        ("亿", "亿"), ("阿1", "1"), ("阿11", "11"), ("阿111", "111"), ("阿101", "101"),
    ],
    "间距": [("无", ""), ("半角", " "), ("全角", "　")],
    "尾部": [
        ("无标点", "{t}"), ("句号", "{t}。"), ("分号", "{t}；"), ("逗号", "{t}，"),
        ("顿号", "{t}、"), ("点", "{t}."), ("省略号", "{t}……"),
        ("右括号", "{t}）"), ("冒号", "{t}："),
    ],
    # 长度：表说中文序号层级只对三、四级卡 60 字——一、二级不卡
    "长度": [("短", 5), ("59字", 59), ("60字", 60), ("61字", 61)],
}

# ---------------------------------------------------------------- 元·阿拉伯序号层级
AR_DIMENSIONS: dict[str, list[tuple[str, str]]] = {
    # 编号本体：一~四级（x / x.y / x.y.z / x.y.z.f）+ 越界形态（三位数、五级）
    "编号": [
        ("1", "1"), ("11", "11"), ("111", "111"), ("1.", "1."),
        ("1.1", "1.1"), ("1.11", "1.11"), ("11.11", "11.11"), ("1.111", "1.111"),
        ("111.11", "111.11"), ("1.1.1", "1.1.1"), ("1.1.111", "1.1.111"),
        ("1.1.1.1", "1.1.1.1"), ("1.1.1.1.1", "1.1.1.1.1"),
    ],
    # 编号与题名之间怎么接（表：可附加点号，随后须跟空格，或直接跟汉字、字母）
    "接法": [
        ("空格", "{seg} {t}"), ("点号空格", "{seg}. {t}"), ("直接接", "{seg}{t}"),
        ("顿号", "{seg}、{t}"), ("全角点", "{seg}．{t}"), ("无空格数字", "{seg}{t}"),
    ],
    "尾部": [
        ("无标点", "{t}"), ("句号", "{t}。"), ("分号", "{t}；"), ("逗号", "{t}，"),
        ("点", "{t}."), ("省略号", "{t}……"), ("右括号", "{t}）"), ("冒号", "{t}："),
    ],
    # 长度：表说阿拉伯数字层级一至四级都卡 60 字
    "长度": [("短", 5), ("59字", 59), ("60字", 60), ("61字", 61)],
}

CN_DEFAULTS = {"编号写法": "顿号", "数字": "一", "间距": "无",
               "尾部": "无标点", "长度": "短"}
AR_DEFAULTS = {"编号": "1", "接法": "空格", "尾部": "无标点", "长度": "短"}

CN_LOOKUP = {name: dict(items) for name, items in CN_DIMENSIONS.items()}
AR_LOOKUP = {name: dict(items) for name, items in AR_DIMENSIONS.items()}


def _padded(length: int) -> str:
    text = TEXT
    if length > len(text):
        text = text + "字" * (length - len(text))
    return text


def render_cn(values: dict[str, str]) -> str:
    template = CN_LOOKUP["编号写法"][values["编号写法"]]
    number = CN_LOOKUP["数字"][values["数字"]]
    tail = CN_LOOKUP["尾部"][values["尾部"]].replace(
        "{t}", _padded(int(CN_LOOKUP["长度"][values["长度"]])))
    return (template.replace("{n}", number)
                    .replace("{gap}", CN_LOOKUP["间距"][values["间距"]])
                    .replace("{t}", tail))


def render_ar(values: dict[str, str]) -> str:
    segment = AR_LOOKUP["编号"][values["编号"]]
    tail = AR_LOOKUP["尾部"][values["尾部"]].replace(
        "{t}", _padded(int(AR_LOOKUP["长度"][values["长度"]])))
    return (AR_LOOKUP["接法"][values["接法"]]
            .replace("{seg}", segment).replace("{t}", tail))


# ---------------------------------------------------------------- 两侧的身份
def engine_roles(path: pathlib.Path, scheme: str = "chinese") -> dict[int, str]:
    """引擎认出来的身份：段号 → 身份（与预估同一套措辞；按同一方案识别）。"""
    return _engine_roles(path, scheme=scheme)


def expectation(path: pathlib.Path, scheme: str = "chinese") -> ex.Expectation:
    shape = read_docx(path)
    return ex.expect(shape, generated_entries=ex.toc_field_paragraphs(path),
                     scheme=scheme)


def _run_payload(payload: tuple[list[str], int, str]) -> tuple[str, str, str, str]:
    """进程池的入口：`Executor.map` 只吃单参数的函数。"""
    return run_case(*payload)


def run_case(lines: list[str], probe: int,
             scheme: str = "chinese") -> tuple[str, str, str, str]:
    """返回 `(结论, 表的说法, 引擎的说法, 依据的规则 ID)`。

    ⚠️ 这个函数会被扔进进程池，所以参数与返回值都必须是可序列化的普通对象。
    """
    path = case_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    document = Document()
    for line in lines:
        document.add_paragraph(line)
    document.save(str(path))

    want = expectation(path, scheme)
    got = engine_roles(path, scheme)
    mine = want.roles.get(probe, "正文")
    theirs = got.get(probe, "正文")
    if probe in want.undecided:
        return "？未判定", want.undecided[probe], theirs, want.rules.get(probe, "")
    if mine == theirs:
        return "✓一致", mine, theirs, want.rules.get(probe, "")
    return "✘分歧", mine, theirs, want.rules.get(probe, "")


def cases(only: str | None):
    """产出 `(名称, 组, 段落文字, 方案, 被测段号)`。

    每组：单选（每个取值一次）+ 成对（维度两两之间），其余维度取默认。
    """
    for scheme, group, dimensions, defaults, render in (
            ("chinese", "中文", CN_DIMENSIONS, CN_DEFAULTS, render_cn),
            ("dotted", "阿拉伯", AR_DIMENSIONS, AR_DEFAULTS, render_ar)):
        names = list(dimensions)
        for name in names:
            for label, _value in dimensions[name]:
                values = dict(defaults)
                values[name] = label
                yield (f"{group}·{name}={label}", group, CONTEXT + [render(values)],
                       scheme, 2)
        for a, b in itertools.combinations(names, 2):
            if only and only not in (a, b):
                continue
            for label_a, _ in dimensions[a]:
                for label_b, _ in dimensions[b]:
                    values = dict(defaults)
                    values[a] = label_a
                    values[b] = label_b
                    yield (f"{group}·{a}={label_a} × {b}={label_b}", group,
                           CONTEXT + [render(values)], scheme, 2)


def main(argv: list[str]) -> int:
    only = argv[argv.index("--only") + 1] if "--only" in argv else None
    jobs = int(argv[argv.index("--jobs") + 1]) if "--jobs" in argv else DEFAULT_JOBS
    limit = int(argv[argv.index("--limit") + 1]) if "--limit" in argv else None

    plan = [(name, lines, scheme, probe)
            for name, _group, lines, scheme, probe in cases(only)]
    if limit is not None:
        plan = plan[:limit]
    payloads = [(lines, probe, scheme) for _name, lines, scheme, probe in plan]

    total = {"✓一致": 0, "✘分歧": 0, "？未判定": 0}
    diffs: list[str] = []
    undecided: list[str] = []
    started = time.perf_counter()
    if jobs > 1:
        with concurrent.futures.ProcessPoolExecutor(max_workers=jobs) as pool:
            outcomes = list(pool.map(_run_payload, payloads))
    else:
        outcomes = [run_case(lines, probe, scheme) for lines, probe, scheme in payloads]

    for (name, lines, scheme, _probe), (conclusion, mine, theirs, rule) in zip(plan, outcomes):
        total[conclusion] += 1
        line = (f"  {conclusion}  {name}\n"
                f"        段落={lines[-1][:40]!r}\n"
                f"        表={mine}  引擎={theirs}  规则={rule or '—'}")
        if conclusion == "✘分歧":
            diffs.append(line)
        elif conclusion == "？未判定":
            undecided.append(line)

    ran = sum(total.values())
    elapsed = time.perf_counter() - started
    print(f"跑完 {ran} 条（{jobs} 进程）："
          + "、".join(f"{key} {value}" for key, value in total.items()))
    print(f"耗时 {elapsed:.1f} 秒（{elapsed / max(ran, 1) * 1000:.0f} 毫秒/条）")
    for title, items in (("分歧（表与引擎不一致）", diffs),
                         ("表里没写（登记，不当 bug）", undecided)):
        if not items:
            continue
        print()
        print("=" * 74)
        print(f"{title}：{len(items)} 条")
        print("=" * 74)
        for item in items[:40]:
            print(item)
        if len(items) > 40:
            print(f"  …（还有 {len(items) - 40} 条，见 _out/gen-matrix.txt）")

    report = HERE / "_out" / "gen-matrix.txt"
    report.write_text("\n".join(diffs + undecided), encoding="utf-8")
    print()
    print(f"明细已写出：{report.relative_to(ROOT)}")
    shutil.rmtree(WORK, ignore_errors=True)
    return 1 if diffs else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
