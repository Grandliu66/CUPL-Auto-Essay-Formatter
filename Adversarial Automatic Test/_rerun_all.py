r"""一次跑完一批夹具的"期望 vs 实际"，**只打结论行**（不打逐段清单）。

存在理由：`aat_expect.py` 一次只吃一个夹具、而且把逐段身份全打出来（十几 KB /
终端一条），
回归时想一眼看清"这次改完，每个夹具还剩几处分歧"非常费事。这里复用它的
部件，只输出每个夹具的一行结论 + 有差异的段落首行。

默认清单 2026-09-19 从 3 份扩到 25 份（见 `FIXTURES`）：`reference/` 与
`reference/用户测试稿/` 里**跑得通的全挂上**（逐份验过：0 份跑失败，
全批 ~12 秒）。某份夹具坏掉只报它自己的"跑失败"一行，不断整层。

用法：.venv\Scripts\python.exe "Adversarial Automatic Test\_rerun_all.py"
      .venv\Scripts\python.exe "Adversarial Automatic Test\_rerun_all.py" 民总.docx
"""
from __future__ import annotations

import pathlib
import shutil
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))

import aat_diff                                                 # noqa: E402
import aat_engine_api as engine_api                            # noqa: E402
import aat_expect as ex                                         # noqa: E402
from aat_docx import read_docx                                  # noqa: E402
from essay_tool.config.preset_io import load_preset             # noqa: E402
from essay_tool.engine import pipeline                          # noqa: E402
from essay_tool.engine.alerts import AlertCollector             # noqa: E402

# 默认清单（2026-09-19 扩）：相对 `reference/` 的路径，子目录直接写成前缀。
# 名字重复的（用户测试稿里的 TOC/一般情景测试/民总-排好版 是给 tools/ 的副本）
# 只取一份。唯一排除的是「中国政法大学硕士、博士学位论文写作规范（更新）.docx」：
# 硕博方向已挂起（docs/硕博规范问题.md）且该文件待移出 reference，挂上只会长期
# 报无关的红；需要时在命令行直接传它的路径。
FIXTURES = [
    "4√正文9.1.docx",
    "PaperStructureTest.docx",
    "PaperTest.docx",
    "test2.docx",
    "TitleIdentify.docx",
    "TOC.docx",
    "一般情景测试.docx",
    "一般情景测试-排好版.docx",
    "中国政法大学本科毕业论文撰写格式规范.docx",
    "中国政法大学本科生毕业论文写作模板.docx",
    "中国政法大学本科生课程考核封面.docx",
    "压力测试_Agent版.docx",
    "压力！.docx",
    "压力！-排好版.docx",
    "图表压力测试.docx",
    "图表脚注测试.docx",
    "平时作业.docx",
    "新建 DOCX 文档.docx",
    "民总.docx",
    "民总-排好版.docx",
    "空白模板.docx",
    "阿拉伯数字.docx",
    "附件七：第二十四届“学术十星”论文大赛论文封面范例.docx",
    "用户测试稿/民总-排好1版.docx",
    "用户测试稿/民总-排好1版-排好版.docx",
]

# 这些夹具按“阿拉伯数字序号层级”跑（界面上的方案单选是用户自己定的；
# 这份文档通篇用 1 / 1.1 / 1.1.1 编号）。2026-09-20 核实：用中文方案跑它
# 出来的是"跑批设置错"——“引言”段带自动编号，中文方案下引擎把反算编号
# 拼进去就成了"1.引言"三级标题；点分方案下两侧都认引言。
DOTTED_FIXTURES = {"阿拉伯数字.docx"}


def one(fixture: pathlib.Path) -> str:
    stem = fixture.stem
    scheme = "dotted" if fixture.name in DOTTED_FIXTURES else "chinese"
    prepared = ex.prepare(fixture, ex.OUT / "_work" / stem)
    shape = read_docx(prepared)
    visible = ex.reverse_numbering(shape.paras, prepared)
    expectation = ex.expect(shape, generated_entries=ex.toc_field_paragraphs(prepared),
                            visible_numbers=visible, scheme=scheme)
    auto_dir = ex.OUT / "_auto" / stem
    want_dir = ex.OUT / "_want" / stem
    shutil.rmtree(auto_dir, ignore_errors=True)
    shutil.rmtree(want_dir, ignore_errors=True)
    auto = pipeline.run(fixture, None, auto_dir, preset_config=load_preset(ex.PRESET),
                        alerts=AlertCollector(), heading_scheme=scheme)
    want = pipeline.run(fixture, None, want_dir, preset_config=load_preset(ex.PRESET),
                        alerts=AlertCollector(), heading_scheme=scheme,
                        index=engine_api.expect_index(prepared, expectation, scheme=scheme))
    outcome = aat_diff.compare_shapes(read_docx(auto), read_docx(want))
    report = aat_diff.render(outcome, max_items=60)
    head = [line for line in report.splitlines() if line.strip().startswith(("差异概览", "结论："))]
    spots = [line.strip() for line in report.splitlines() if line.strip().startswith(("段", "  段"))]
    lines = [f"=== {fixture.name}", f"    {head[0] if head else '（无概览行）'}"]
    lines += [f"    {spot[:80]}" for spot in spots]
    lines += [f"    自动编号：{len(visible)} 段，未判定：{len(expectation.undecided)} 段"]
    return "\n".join(lines)


def main() -> int:
    names = sys.argv[1:] or list(FIXTURES)
    failed = 0
    for name in names:
        given = pathlib.Path(name)
        # 命令行给的名字先在原位找，再按 reference/ 与 reference/用户测试稿/ 找。
        candidates = [given, ex.ROOT / "reference" / name,
                      ex.ROOT / "reference" / "用户测试稿" / name]
        fixture = next((path for path in candidates if path.exists()), None)
        if fixture is None:
            failed += 1
            print(f"=== {pathlib.Path(name).name}")
            print("    结论：跑失败——找不到这份夹具")
            print()
            continue
        try:
            print(one(fixture))
        except Exception as error:  # noqa: BLE001 —— 一份坏掉不能断整层
            failed += 1
            print(f"=== {fixture.name}")
            print(f"    结论：跑失败——{str(error)[:120]}")
        print()
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
