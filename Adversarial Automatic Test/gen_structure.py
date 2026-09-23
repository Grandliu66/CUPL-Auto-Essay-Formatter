r"""夹具生成器 v3（结构层穷举）：部件的有无 / 顺序 / 重复 / 四不像填充。

前两层各绕开了"部件"这一维：
  * `gen_matrix.py`（1229 条）——部件齐全、位置正常，只换**标题的写法**；
  * `gen_pipeline.py`（372 条）——部件不动，只人为弄坏一处 OOXML 构件。
这一层补的就是**部件本身换组合**：缺件、重复、错序、以及把"四不像"的段落
塞在各个缝隙里。

判据从哪来：**不需要新判据**。表 3.6 的识别规则本来就是"逐段的身份断言"，
结构一变，期望跟着变——`aat_expect.expect()` 是那份表的独立手写实现（1229 条
写法层已经用它对齐过），这里直接拿来当标准答案，**逐段**比对，不再只看一段。

为什么这一层不能靠随机抽样：部件是**有限可列**的，列全比抽样强；而且随机会
漏掉"两个部件同时缺"这类组合，而那恰恰是锚链模型最容易翻车的地方（见
`apply_missing_parts()` 的注释：链条断在中途时会一路往上垒到文档最前面）。

用户 2026-09-17 特别交代：**"四不像"的段落要到处塞遍**，而且
「不要局限于'这个和规则不符合，所以不试'，而多试试'虽然这样写是被规则禁止、
不能正确识别的，但还是要看看产物和规则的预测一样不一样'」。
所以填充物里特意混进了**规则明令禁止的写法**（`廿一、` 这种不在允许编号集里的
编号、带句号结尾的"标题"、长到不像标题的标题）——它们的期望同样是确定的
（**不识别为标题／不识别为部件**），工具只要敢认，这里就红。

用法：
    .venv\Scripts\python.exe "Adversarial Automatic Test\gen_structure.py"
    .venv\Scripts\python.exe "Adversarial Automatic Test\gen_structure.py" --only 填充 --jobs 4
"""
from __future__ import annotations

import concurrent.futures
import itertools
import pathlib
import sys
import time
import traceback

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent
for entry in (str(ROOT), str(HERE)):
    if entry not in sys.path:
        sys.path.insert(0, entry)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import aat_expect as ex                                          # noqa: E402
from aat_docx import read_docx                                   # noqa: E402
from docx import Document                                        # noqa: E402
from gen_matrix import engine_roles                              # noqa: E402

WORK = HERE / "_out" / "gen_structure"
DEFAULT_JOBS = 4

# ---------------------------------------------------------------- 部件池
# 每个部件 = 它该有的那几段（标签 + 内容）。**一律不挂样式**：这一层测的是
# "认不认识部件"，样式是写法层与产物层的事，掺进来会让红条归因不清。
PART_POOL: dict[str, list[str]] = {
    "题目": ["论测试文档的结构部件识别"],
    "副标题": ["——以结构层为例"],
    "摘要": ["【摘要】", "本文测的是结构层的部件识别。"],
    "关键词": ["【关键词】", "结构部件  识别  测试"],
    "英文题目": ["A STUDY ON STRUCTURE RECOGNITION"],
    "英文副标题": ["——A CASE STUDY"],
    "英文摘要": ["【ABSTRACT】", "This paper tests structure recognition."],
    "英文关键词": ["【KEYWORDS】", "structure  recognition  test"],
    "目录": ["目录"],
    "正文": ["引言", "一、研究背景", "本文研究的是……。", "（一）问题所在", "细节如上……。"],
    "参考文献": ["参考文献", "[1] 张三. 测试[M]. 北京: 某社, 2026."],
    "致谢": ["致谢", "感谢所有给予帮助的人。"],
    "附录": ["附录A  材料", "附录里的材料内容。"],
}
ORDER = list(PART_POOL)          # 规范顺序（= 表 3.6 的部件次序）

# ---------------------------------------------------------------- 四不像
# 用户 2026-09-17：「将四不像的正文到处塞一塞看看效果，识别的尽可能全一点」。
# 每一句都是**真实稿子里会出现、但归类模糊**的写法；其中三条是规则**明令禁止**
# 的编号写法——期望同样是确定的（不识别为标题），工具敢认就是分歧。
FILLERS: dict[str, str] = {
    "作者行": "张三　指导教师：李四　2026 年 9 月",
    "日期行": "2026 年 9 月 17 日",
    "内部标记": "zmarkA",
    "版本草稿": "V1.2 草稿",
    "括号提示": "（此处填写）",
    "英文散句": "This paragraph should not be a title",
    "下划线行": "____________",
    "单个字": "嗯",
    "带句号短行": "这一句末尾带句号，所以不该被当成题名。",
    "禁止编号廿一": "廿一、这个编号不在允许集里",
    "禁止编号百开头": "百二十一、首字是禁止的",
    "超长标题": "一、这一段标题长得非常离谱，长到根本不像标题而像一整段正文的文字，而且还带着句号。",
}

# 缝隙用**部件名**定位，不用数字下标（结构一变下标就漂）。
SLOTS: list[str] = ["最前", "题目前", "题目后", "副标题后", "摘要前", "摘要后",
                    "正文前", "参考文献前", "末"]

# ---------------------------------------------------------------- 生成
def _lines(tokens: list[tuple[str, str]]) -> list[str]:
    """把"令牌序列"摊平成段落文字。令牌：('part', 部件名) 或 ('text', 一段字)。"""
    out: list[str] = []
    for kind, value in tokens:
        if kind == "part":
            out.extend(PART_POOL[value])
        else:
            out.append(value)
    return out


def _base(skip: tuple[str, ...] = ()) -> list[tuple[str, str]]:
    return [("part", name) for name in ORDER if name not in skip]


def _insert(tokens: list[tuple[str, str]], slot: str,
            item: tuple[str, str]) -> list[tuple[str, str]]:
    """把 item 塞进 slot 指定的缝隙。"""
    if slot == "最前":
        return [item] + tokens
    if slot == "末":
        return tokens + [item]
    # 其余缝隙：找"某个部件的第一段"的位置
    part = {"题目前": "题目", "题目后": "题目", "副标题后": "副标题",
            "摘要前": "摘要", "摘要后": "摘要", "正文前": "正文",
            "参考文献前": "参考文献"}[slot]
    after = slot.endswith("后")
    position = None
    cursor = 0
    for index, (kind, value) in enumerate(tokens):
        if kind == "part" and value == part:
            position = index
            cursor = index
            break
    if position is None:
        return tokens + [item]          # 那个部件本来就不在（缺席轴的情形）⇒ 垫到末尾
    return tokens[:cursor + (1 if after else 0)] + [item] \
        + tokens[cursor + (1 if after else 0):]


def cases(only: str | None):
    """产出 `(名称, 组, 段落文字)`。

    五条轴，都是**有限可列**的：
      ① 缺席：单件 13 条 + 成对 78 条（pairwise 覆盖，不穷举 2^13）
      ② 重复：相邻重复 13 条 + 挪到末尾再写一遍 13 条
      ③ 错序：每件挪到最前 / 挪到最后，26 条
      ④ 填充：12 条四不像 × 9 个缝隙，108 条
    """
    for name in ORDER:
        yield f"缺席={name}", "缺席", _lines(_base((name,)))
    for a, b in itertools.combinations(ORDER, 2):
        yield f"缺席={a}×{b}", "缺席", _lines(_base((a, b)))
    for name in ORDER:
        tokens = [("part", part) for part in ORDER]
        index = ORDER.index(name)
        tokens.insert(index + 1, ("part", name))
        yield f"重复相邻={name}", "重复", _lines(tokens)
    for name in ORDER:
        tokens = [("part", part) for part in ORDER if part != name]
        tokens.append(("part", name))
        yield f"重复在末尾={name}", "重复", _lines(tokens)
    for name in ORDER:
        head = [("part", name)] + [("part", part) for part in ORDER if part != name]
        tail = [("part", part) for part in ORDER if part != name] + [("part", name)]
        yield f"错序={name}挪最前", "错序", _lines(head)
        yield f"错序={name}挪最后", "错序", _lines(tail)
    for filler, text in FILLERS.items():
        for slot in SLOTS:
            tokens = _insert(_base(), slot, ("text", text))
            yield f"填充={filler}@{slot}", "填充", _lines(tokens)


# ---------------------------------------------------------------- 一条
def _roles(both: dict[int, str], count: int) -> dict[int, str]:
    """把缺失的段号补成"正文"，好逐段对齐着比。"""
    return {index: both.get(index, "正文") for index in range(count)}


def _same(want: str, got: str) -> bool:
    """两边的身份算不算同一个。

    表把“标签下面那几段”算作该部件的一部分（R3/R4/R7/R8 的「正文」子项），而引擎
    的识别结果里**没有这一级**（它只报标签、标题、题目、题注）。所以预估说
    “摘要正文”、引擎说“正文”，不算分歧——只要引擎没把那段认成部件或标题就行。

    同理，表里“手打目录条目”分两种下场（R9 例外）：域内的随旧目录一起删、
    **域外的被认成“多余的正文”**。引擎的识别结果里统一叫“正文”，是同一回事，
    不算分歧——它是不是域内，要到产物层（删除）才看得出来。
    """
    if want == got:
        return True
    if got == "正文" and want.endswith("正文"):
        return True
    return got == "正文" and want == "多余的正文（手打目录条目）"


def run_case(payload: tuple[str, str, list[str]]) -> tuple[str, str, str, str]:
    """返回 `(结论, 组, 用例名, 说明)`。

    ⚠️ 会被扔进进程池：参数与返回值必须是可序列化的普通对象。
    """
    name, group, lines = payload
    path = WORK / f"{abs(hash(name)) % 10**8:08d}.docx"
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        document = Document()
        for line in lines:
            document.add_paragraph(line)
        document.save(str(path))

        want = ex.expect(read_docx(path))
        got = engine_roles(path)
    except Exception:
        return ("✘崩溃", group, name,
                traceback.format_exc().strip().splitlines()[-1])

    mine = _roles(want.roles, len(lines))
    theirs = _roles(got, len(lines))
    diffs: list[str] = []
    for index in range(len(lines)):
        if index in want.undecided:
            diffs.append(f"段[{index}] 未判定：{want.undecided[index]}"
                         f"（引擎={theirs[index]}） '{lines[index][:26]}'")
            continue
        if not _same(mine[index], theirs[index]):
            rule = want.rules.get(index, "")
            diffs.append(f"段[{index}] 期望={mine[index]} 引擎={theirs[index]}"
                         f"（{rule or '无规则'}） '{lines[index][:26]}'")
    if not diffs:
        return ("✓一致", group, name, "")
    return ("✘分歧", group, name, "\n".join(diffs[:6]))


def _output_paths(tag: str) -> pathlib.Path:
    suffix = f"-{tag}" if tag else ""
    return HERE / "_out" / f"gen-structure{suffix}.txt"


def main(argv: list[str]) -> int:
    only = argv[argv.index("--only") + 1] if "--only" in argv else None
    jobs = int(argv[argv.index("--jobs") + 1]) if "--jobs" in argv else DEFAULT_JOBS
    limit = int(argv[argv.index("--limit") + 1]) if "--limit" in argv else None
    tag = argv[argv.index("--tag") + 1] if "--tag" in argv else ""

    plan = [(name, group, lines) for name, group, lines in cases(only)
            if not only or only in name or only == group]
    if limit is not None:
        plan = plan[:limit]

    started = time.perf_counter()
    if jobs > 1:
        with concurrent.futures.ProcessPoolExecutor(max_workers=jobs) as pool:
            outcomes = list(pool.map(run_case, plan))
    else:
        outcomes = [run_case(item) for item in plan]
    elapsed = time.perf_counter() - started

    tally: dict[str, int] = {}
    for conclusion, _group, _name, _detail in outcomes:
        tally[conclusion] = tally.get(conclusion, 0) + 1
    print(f"跑完 {len(outcomes)} 条（{jobs} 进程）："
          + "、".join(f"{key} {value}" for key, value in sorted(tally.items())))
    print(f"耗时 {elapsed:.1f} 秒")

    print()
    print("按组：")
    groups: dict[str, dict[str, int]] = {}
    for conclusion, group, _name, _detail in outcomes:
        counter = groups.setdefault(group, {})
        counter[conclusion] = counter.get(conclusion, 0) + 1
    for group in sorted(groups):
        items = groups[group]
        print(f"  {group:<6} " + "、".join(f"{key} {value}"
                                          for key, value in sorted(items.items())))

    print()
    for conclusion, _group, name, detail in outcomes:
        if conclusion.startswith("✓"):
            continue
        print(f"  {conclusion}  {name}")
        for line in (detail or "").splitlines():
            print(f"        {line}")

    report = _output_paths(tag)
    report.write_text("\n".join(f"{conclusion}\t{group}\t{name}\t{detail}"
                                for conclusion, group, name, detail in outcomes),
                      encoding="utf-8")
    print()
    print(f"明细已写出：{report.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
