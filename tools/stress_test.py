# -*- coding: utf-8 -*-
"""压力测试程序：拿 tools/make_stress_docx.py 生成的夹具跑管线，逐条验行为。

两类结论，分开报（这是刻意的——判据没定的东西不能算"失败"）：

* **硬断言 H**：不依赖任何"规范理解"，只有两种结果。失败＝真 Bug。
    H01 不崩        —— 各种参数组合下管线不抛异常
    H02 不丢        —— 输入里的每个标记在产物里至少出现一次
    H03 不复制      —— 输入里的每个标记在产物里最多出现一次
    H04 幂等        —— 产物再喂一遍，标记/域/占位语数量都不增加
    H05 数量守恒    —— 产物段落数 <= 输入段落数 + 启用部件数*3
    H06 部件顺序    —— 题目 < 摘要 < 目录 < 引言 < 结论 < 参考文献 < 致谢
    H07 清单域可探  —— 工具自己写的清单域（TOC + \\t）必须被探测端认出来
    H08 已有目录域  —— 真目录域不被替换、旧条目不被套成正文
    H09 删封面彻底  —— 封面区段落 + 封面区**表格**都删掉，区外的不动
    H10 题注编号   —— **不再写任何域**（2026-09-16 题注编号功能整体删除），
                      源文件里手打的编号原样保留
    H11 续行/分图   —— “表 1（续）”与“图 2（a）分图行”都算题注：分图挂样式
                      （进清单），续行不挂样式（不进清单）；两边都不动编号
                      用户 2026-09-15：“分图是两张图一个编号”、“续图压根是同一张图”。
    H18 进清单者   —— 挂不挂图题/表题样式 = 进不进清单：分图挂（进清单），
                         续行与英文对照行不挂（不进清单）
    H12 参数当场生效 —— 传 preset_config 后，产物页面边距 == 传入值
    H13 封面区不报警  —— 用户 2026-09-15 裁决：一旦被判为封面，就不参与识别、
                         不算进清单、也不该出现在告警里
    H14 同类告警合并  —— 用户 2026-09-15 裁决：同一问题只报一条，段号并列
                         （如「孤立：…在段落（38、40、43、45、47）」。）
    H15 自动编号固化  —— Word 自动编号的标题，编号必须变成真文字且 numPr 被摘掉
    H16 修订只在封面  —— 修订标记只在封面区里时不算“未接受的修订”告警
    H17 手打目录条目  —— 紧贴目录/清单标签、带页码尾巴的行不当标题
                         （原为观察项 O10，用户 2026-09-15 拍板修完后升为硬断言）

* **观察项 O**：只报事实、不下判断（判据要用户拍板的那种）。
    O01 识别结果 / O02 告警清单 / O03 导出与打开结构页的口径是否一致
    O04 结构页是否铺出封面区的行 / O05 修订 / O06 自动编号 / O07 未编号图表
    O08 部件顺序 / O09 封面区不参与告警

用法：.venv\\Scripts\\python.exe tools\\stress_test.py
产物：work/_stress/run-*/ 各次导出的成品；work/_stress_report.txt 报告
"""

from __future__ import annotations

import inspect
import json
import pathlib
import re
import shutil
import sys
import traceback

ROOT = pathlib.Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from docx import Document  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

FIXTURE = ROOT / "reference" / "压力测试_Agent版.docx"
MANIFEST = ROOT / "work" / "_stress_manifest.json"
WORK = ROOT / "work" / "_stress"
REPORT = ROOT / "work" / "_stress_report.txt"

MARK_RE = re.compile(r"Z\d\d|zmark[A-Z]\d?")
results: list[tuple[str, str, str, str]] = []   # (id, 标题, PASS/FAIL/SKIP/OBS, 详情)


def record(cid: str, title: str, status: str, detail: str = "") -> None:
    results.append((cid, title, status, detail))


def ok(cid: str, title: str, detail: str = "") -> None:
    record(cid, title, "PASS", detail)


def fail(cid: str, title: str, detail: str) -> None:
    record(cid, title, "FAIL", detail)


def skip(cid: str, title: str, detail: str = "") -> None:
    record(cid, title, "SKIP", detail)


def obs(cid: str, title: str, detail: str) -> None:
    record(cid, title, "OBS", detail)


# ------------------------------------------------------------------ 工具函数
def doc_texts(doc) -> list[str]:
    """所有段落文字（含表格单元格里的），按 body 顺序近似。"""
    texts = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    texts.append(p.text)
    return texts


def markers_of(doc) -> dict[str, int]:
    counts: dict[str, int] = {}
    for text in doc_texts(doc):
        for m in MARK_RE.findall(text):
            counts[m] = counts.get(m, 0) + 1
    return counts


def field_instrs(doc) -> list[str]:
    body = doc._element.body
    out = [el.text or "" for el in body.iter(qn("w:instrText"))]
    out += [el.get(qn("w:instr")) or "" for el in body.iter(qn("w:fldSimple"))]
    return out


def count_fields(doc, needle: str) -> int:
    return sum(1 for s in field_instrs(doc) if needle in s)


def placeholder_count(doc) -> int:
    return sum(1 for p in doc.paragraphs if p.text.strip().startswith("［"))


def squeeze(text: str) -> str:
    """去掉所有空白（含全角空格）——"引　言" 与 "引言" 要能对上。"""
    return re.sub(r"\s|\u3000", "", text)


def first_index_of(doc, needle: str) -> int | None:
    target = squeeze(needle)
    for i, p in enumerate(doc.paragraphs):
        if target in squeeze(p.text):
            return i
    return None


def marker_locations(doc) -> dict[str, object]:
    """标记 -> 所在段落号；只在表格里的记 ('table', 表序号)。"""
    out: dict[str, object] = {}
    for i, p in enumerate(doc.paragraphs):
        for m in MARK_RE.findall(p.text):
            out.setdefault(m, i)
    for ti, table in enumerate(doc.tables):
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for m in MARK_RE.findall(p.text):
                        out.setdefault(m, ("table", ti))
    return out


def call_supported(fn, *args, **kwargs):
    """只传目标函数签名里有的关键字参数（接口在动，脚本得跟着容错）。"""
    sig = inspect.signature(fn)
    accepted = {k: v for k, v in kwargs.items() if k in sig.parameters}
    ignored = sorted(set(kwargs) - set(accepted))
    return fn(*args, **accepted), ignored


# ------------------------------------------------------------------ 准备
lines: list[str] = []


def say(s: str = "") -> None:
    lines.append(str(s))


if not FIXTURE.exists():
    print("夹具不存在，请先跑：python tools\\make_stress_docx.py")
    raise SystemExit(2)

manifest = json.loads(MANIFEST.read_text(encoding="utf-8")) if MANIFEST.exists() else {}
cover_end = manifest.get("cover_end")
title_index = manifest.get("title_index")

src_doc = Document(str(FIXTURE))
src_markers = markers_of(src_doc)
src_locations = marker_locations(src_doc)
# 封面区里的标记：段落号 <= cover_end 的，或落在第一张表（封面表格）里的
def in_cover(mark: str) -> bool:
    loc = src_locations.get(mark)
    if isinstance(loc, tuple):
        return loc[1] == 0          # 第一张表就是封面区那张
    return cover_end is not None and isinstance(loc, int) and loc <= cover_end
src_paras = len(src_doc.paragraphs)
src_toc_fields = count_fields(src_doc, "TOC")
src_seq_fields = count_fields(src_doc, "SEQ")
src_placeholders = placeholder_count(src_doc)
src_toc_entries = sum(1 for p in src_doc.paragraphs
                      if str(p.style.name).lower().startswith("toc "))

say("=" * 100)
say("压力测试报告")
say("夹具：%s" % FIXTURE.name)
say("输入：%d 段 / 标记 %d 个 / TOC 类域 %d 个 / SEQ 域 %d 个 / 占位语 %d 段 / toc 样式条目 %d 段"
    % (src_paras, len(src_markers), src_toc_fields, src_seq_fields, src_placeholders, src_toc_entries))
say("封面范围候选 cover_end=%s（真题目在段 %s）" % (cover_end, title_index))
say("=" * 100)

# 预设
from essay_tool.config.preset_io import load_preset  # noqa: E402
from essay_tool.engine.pipeline import review, run  # noqa: E402

PRESET_PATH = ROOT / "presets" / "中国政法大学-本科毕业论文.json"
base_preset = load_preset(PRESET_PATH)
enabled_parts = sum(1 for v in base_preset.structure_required.values() if v)
say("预设：%s（启用部件 %d 个）" % (PRESET_PATH.name, enabled_parts))
say()

if WORK.exists():
    shutil.rmtree(WORK, ignore_errors=True)
WORK.mkdir(parents=True, exist_ok=True)

# ------------------------------------------------------------------ 各种跑法
RUNS = [
    ("no-cover",  dict(cover_end=None, delete_cover=False)),
    ("cover",     dict(cover_end=cover_end, delete_cover=False)),
    ("del-cover", dict(cover_end=cover_end, delete_cover=True)),
    ("all-parts", dict(cover_end=cover_end, delete_cover=False, override_all=True)),
]

outputs: dict[str, pathlib.Path] = {}
for tag, opts in RUNS:
    preset = load_preset(PRESET_PATH)
    kwargs = {
        "preset_config": preset,
        "cover_end": opts["cover_end"],
        "delete_cover": opts["delete_cover"],
        "check_source_alerts": False,
    }
    if opts.get("override_all"):
        kwargs["structure_override"] = {k: True for k in preset.structure_required}
    out_dir = WORK / ("run-" + tag)
    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        path, ignored = call_supported(run, FIXTURE, None, out_dir, **kwargs)
        outputs[tag] = pathlib.Path(path)
        detail = "产物 %s" % pathlib.Path(path).name
        if ignored:
            detail += "（这些参数管线不认、已跳过：%s）" % ", ".join(ignored)
        ok("H01/" + tag, "不崩：run(%s)" % tag, detail)
    except Exception as exc:                                     # noqa: BLE001
        tb = traceback.format_exc().strip().splitlines()[-3:]
        fail("H01/" + tag, "不崩：run(%s)" % tag,
             "%s: %s\n        %s" % (type(exc).__name__, exc, "\n        ".join(tb)))

# ------------------------------------------------------------------ H02/H03 标记守恒
for tag, path in outputs.items():
    if not path.exists():
        skip("H02/" + tag, "标记守恒：%s" % tag, "产物不存在")
        continue
    out_doc = Document(str(path))
    counts = markers_of(out_doc)
    deleted = (tag == "del-cover")

    missing = [m for m in src_markers if counts.get(m, 0) == 0]
    if deleted:
        # 封面区里的东西本来就该被删掉，不算"丢"
        missing = [m for m in missing if not in_cover(m)]
    if missing:
        detail = "丢了 %d 个标记：%s" % (len(missing), "、".join(sorted(missing)[:8]))
        for m in sorted(missing)[:3]:
            detail += "\n        %s（%s）位置=%s" % (
                m, manifest.get("marks", {}).get(m, "?"), src_locations.get(m))
        fail("H02/" + tag, "不丢：%s 的标记都在产物里" % tag, detail)
    else:
        ok("H02/" + tag, "不丢：%s 的标记都在产物里" % tag,
           "共查 %d 个标记" % len(src_markers))

    dup = {m: c for m, c in counts.items() if c > 1}
    if dup:
        detail = "有 %d 个标记出现多次（输入里各 1 次）：" % len(dup)
        for m, c in sorted(dup.items(), key=lambda kv: -kv[1])[:10]:
            detail += "\n        %s 出现 %d 次  ← %s" % (m, c, manifest.get("marks", {}).get(m, "?"))
        fail("H03/" + tag, "不复制：%s 的标记各出现 1 次" % tag, detail)
    else:
        ok("H03/" + tag, "不复制：%s 的标记各出现 1 次" % tag,
           "最大重复次数 = 1")

# ------------------------------------------------------------------ H05 数量守恒
for tag, path in outputs.items():
    if not path.exists():
        continue
    out_doc = Document(str(path))
    limit = src_paras + enabled_parts * 3
    n = len(out_doc.paragraphs)
    detail = "输入 %d 段 + 启用 %d 部件×3 = 上限 %d；产物 %d 段" % (
        src_paras, enabled_parts, limit, n)
    if n > limit:
        fail("H05/" + tag, "数量守恒：%s" % tag, detail + " ← 超出，可能有复制或重复插入")
    else:
        ok("H05/" + tag, "数量守恒：%s" % tag, detail)

# ------------------------------------------------------------------ H06 部件顺序
ORDER = ["论网络平台侵权责任的边界", "【摘要】", "【目录】", "引　言",
         "结　论", "参考文献", "致　谢"]
for tag, path in outputs.items():
    if not path.exists():
        continue
    out_doc = Document(str(path))
    found = {name: first_index_of(out_doc, name) for name in ORDER}
    worry = [name for name, idx in found.items() if idx is None]
    if worry:
        obs("O08/" + tag, "部件顺序：（%s）里找不到：%s" % (tag, "、".join(worry)),
            "实际位置：" + "、".join("%s=%s" % (k, v) for k, v in found.items()))
        continue
    seq = [found[name] for name in ORDER]
    if seq == sorted(seq):
        ok("H06/" + tag, "部件顺序：%s" % tag,
           "、".join("%s@%d" % (k, v) for k, v in found.items()))
    else:
        fail("H06/" + tag, "部件顺序：%s" % tag,
             "顺序不对：" + "、".join("%s@%d" % (k, v) for k, v in found.items()))

# ------------------------------------------------------------------ H07 清单域可探
try:
    from essay_tool.engine.detect import body_list_field_kinds, body_toc_field_parts
    for tag in ("cover", "no-cover"):
        path = outputs.get(tag)
        if path is None or not path.exists():
            continue
        out_doc = Document(str(path))
        kinds = body_list_field_kinds(out_doc)
        parts = body_toc_field_parts(out_doc)
        n_list = count_fields(out_doc, '\\t "图题')
        if "图" in kinds and "插图清单" in parts and n_list == 1:
            ok("H07/" + tag, "清单域可探：%s" % tag,
               "list_field_kinds=%s；parts=%s；\\t 清单域个数=%d" % (sorted(kinds), sorted(parts), n_list))
        else:
            fail("H07/" + tag, "清单域可探：%s" % tag,
                 "kinds=%s（应含 图）；parts=%s（应含 插图清单）；\\t 清单域个数=%d（应为 1）"
                 % (sorted(kinds), sorted(parts), n_list))
except Exception as exc:                                          # noqa: BLE001
    skip("H07", "清单域可探", "无法验证：%s: %s" % (type(exc).__name__, exc))

# ------------------------------------------------------------------ H08 已有真目录域
for tag in ("cover", "no-cover"):
    path = outputs.get(tag)
    if path is None or not path.exists():
        continue
    out_doc = Document(str(path))
    n_toc = count_fields(out_doc, "TOC \\o")
    entries = sum(1 for p in out_doc.paragraphs
                  if str(p.style.name).lower().startswith("toc "))
    detail = "TOC 域 %d 个（输入 %d）；toc 样式条目 %d 段（输入 %d）" % (
        n_toc, count_fields(src_doc, "TOC \\o"), entries, src_toc_entries)
    if n_toc == count_fields(src_doc, "TOC \\o") and entries >= src_toc_entries:
        ok("H08/" + tag, "已有真目录域不被动：%s" % tag, detail)
    else:
        fail("H08/" + tag, "已有真目录域不被动：%s" % tag,
             detail + " ← 目录域被替换/重复插入，或旧条目被改了样式")

# ------------------------------------------------------------------ H09 删封面
path = outputs.get("del-cover")
if path is not None and path.exists():
    out_doc = Document(str(path))
    counts = markers_of(out_doc)
    left_in_cover = {m: counts.get(m, 0) for m in ("zmarkT1", "zmarkT2")}
    after_cover = counts.get("zmarkA", 0)
    if any(left_in_cover.values()):
        fail("H09", "删封面删干净（含封面里的表格）",
             "封面区表格标记还在：%s" % left_in_cover)
    elif after_cover == 0:
        fail("H09", "删封面删干净（含封面里的表格）",
             "封面区**之外**的标记 zmarkA 也被删了（删多了）")
    else:
        ok("H09", "删封面删干净（含封面里的表格）",
           "封面区表格标记已消失，封面区外的 zmarkA 仍在")
else:
    skip("H09", "删封面删干净", "del-cover 那次导出没成功")

# ------------------------------------------------------------------ H10/H11 题注（只挂样式）
def paragraph_with(doc, marker: str):
    for p in doc.paragraphs:
        if marker in p.text:
            return p
    return None


def seq_fields_of(doc, marker: str) -> list:
    """这一段里的 SEQ 域指令（两种写法都收）——用来验"我们没往文档里塞域"。"""
    p = paragraph_with(doc, marker)
    if p is None:
        return []
    found = [fld.get(qn("w:instr")) or "" for fld in p._p.iter(qn("w:fldSimple"))]
    code = ""
    for el in p._p.iter():
        if el.tag == qn("w:instrText"):
            code += el.text or ""
        elif el.tag == qn("w:fldChar") and el.get(qn("w:fldCharType")) == "separate":
            found.append(" ".join(code.split()))
            code = ""
    return [c for c in found if c.startswith("SEQ")]


def style_of_mark(doc, marker: str) -> str:
    p = paragraph_with(doc, marker)
    if p is None or p.style is None:
        return ""
    return str(p.style.name).strip().lower()


try:
    on_doc = Document(str(outputs["cover"])) if "cover" in outputs else None
    if on_doc is None:
        skip("H10/H11", "题注只挂样式", "cover 那次导出没成功")
    else:
        # H10：**题注化整个功能没有了**（用户 2026-09-16：“我们只管样式、编号与
        # 我们无关”）。夹具里那几条**手打**编号的题注（zmarkM/…)现在一个字
        # 都不该被改成域；图片本身自带的域（夹具里没有）也保持原样。
        stray = {m: seq_fields_of(on_doc, m)
                 for m in ("zmarkM", "zmarkN", "zmarkO", "zmarkQ", "zmarkR")}
        bad = {m: v for m, v in stray.items() if v}
        if bad:
            fail("H10", "题注不再被题注化（不写域、不改号）",
                 "这些段落被塞了域：%s" % bad)
        else:
            ok("H10", "题注不再被题注化（不写域、不改号）",
               "五条手打题注都没有 SEQ 域")

        # H11：手打编号原样保留（不再被重排成连续号）
        keep = [m for m in ("zmarkM", "zmarkO", "zmarkQ", "zmarkR")
                if paragraph_with(on_doc, m) is None]
        if keep:
            fail("H11", "手打编号原样保留", "成品里找不到 %s" % "、".join(keep))
        else:
            texts = {m: paragraph_with(on_doc, m).text.strip()
                     for m in ("zmarkM", "zmarkO", "zmarkQ", "zmarkR")}
            expected = {"zmarkM": "图 1", "zmarkO": "表 1",
                        "zmarkQ": "表 1（续）", "zmarkR": "图 2（a）"}
            wrong = {m: t for m, t in texts.items() if not t.startswith(expected[m])}
            if wrong:
                fail("H11", "手打编号原样保留", "实际 %s" % wrong)
            else:
                ok("H11", "手打编号原样保留",
                   "编号没被动：%s" % {m: t[:14] for m, t in texts.items()})

        # H18：进不进清单 ＝ 挂不挂图题/表题样式（清单域按样式名收）。
        # **不区分主/分**（用户 2026-09-16）：分图也进；续行与英文对照行不进。
        styles = {m: style_of_mark(on_doc, m)
                  for m in ("zmarkM", "zmarkN", "zmarkO", "zmarkQ", "zmarkR")}
        want = {"zmarkM": "图题", "zmarkO": "表题", "zmarkR": "图题"}
        want_not = ("zmarkN", "zmarkQ")
        problems = [m for m, style in want.items() if styles[m] != style]
        problems += [m for m in want_not if styles[m] in ("图题", "表题", "caption", "题注")]
        if problems:
            fail("H18", "题注都进清单、续与英文不进",
                 "样式不对：%s（全部 %s）" % ("、".join(problems), styles))
        else:
            ok("H18", "题注都进清单、续与英文不进",
               "题注（含分图）=%s；英文与续行不挂样式" % styles)
except Exception as exc:                                          # noqa: BLE001
    skip("H10/H11/H18", "题注相关检查", "%s: %s" % (type(exc).__name__, exc))

# ------------------------------------------------------------------ H12 preset_config 当场生效
try:
    from docx.shared import Cm
    preset = load_preset(PRESET_PATH)
    preset.page.margin_left = Cm(3.5)
    preset.page.margin_top = Cm(4.1)
    out_dir = WORK / "run-form"
    out_dir.mkdir(parents=True, exist_ok=True)
    path, _ignored = call_supported(run, FIXTURE, None, out_dir,
                                   preset_config=preset, cover_end=cover_end,
                                   check_source_alerts=False)
    out_doc = Document(str(path))
    want_l, want_t = Cm(3.5), Cm(4.1)
    tol = 2000          # docx 里边距存的是 twips，1 twip = 635 EMU，容这点误差
    got = [(s.left_margin, s.top_margin) for s in out_doc.sections]
    bad = [(l, t) for (l, t) in got
           if abs(l - want_l) > tol or abs(t - want_t) > tol]
    if not bad:
        ok("H12", "传 preset_config 时表单值当场生效",
           "所有节（%d 个）边距 == 传入值 %s（容差 %d EMU）" % (len(got), (want_l, want_t), tol))
    else:
        fail("H12", "传 preset_config 时表单值当场生效",
             "期望 %s，实际 %s" % ((want_l, want_t), bad))
except Exception as exc:                                          # noqa: BLE001
    fail("H12", "传 preset_config 时表单值当场生效",
         "%s: %s" % (type(exc).__name__, exc))

# ------------------------------------------------------------------ H04 幂等
try:
    first = outputs.get("cover")
    if first is not None and first.exists():
        out_dir = WORK / "run-idempotent"
        out_dir.mkdir(parents=True, exist_ok=True)
        path, _ignored = call_supported(run, first, None, out_dir,
                                       preset_config=load_preset(PRESET_PATH),
                                       check_source_alerts=False)
        second = Document(str(path))
        counts = markers_of(second)
        dup = {m: c for m, c in counts.items() if c > 1}
        d1, d2 = Document(str(first)), second
        fields_1, fields_2 = len(field_instrs(d1)), len(field_instrs(d2))
        ph_1, ph_2 = placeholder_count(d1), placeholder_count(d2)
        paras_1, paras_2 = len(d1.paragraphs), len(d2.paragraphs)
        detail = ("标记重复 %s；域 %d→%d；占位语 %d→%d；段落 %d→%d"
                  % (dup or "无", fields_1, fields_2, ph_1, ph_2, paras_1, paras_2))
        if dup or fields_2 > fields_1 + 1 or ph_2 > ph_1 or paras_2 > paras_1 + 2:
            fail("H04", "幂等：成品回炉不再变形", detail)
        else:
            ok("H04", "幂等：成品回炉不再变形", detail)
    else:
        skip("H04", "幂等：成品回炉不再变形", "第一次导出没成功")
except Exception as exc:                                          # noqa: BLE001
    fail("H04", "幂等：成品回炉不再变形", "%s: %s" % (type(exc).__name__, exc))

# ------------------------------------------------------------------ O01 识别结果
try:
    from essay_tool.engine.detect import (detect_captions, detect_document_title,
                                          detect_english_subtitle, detect_english_title,
                                          detect_front_matter_labels, detect_headings)
    doc = Document(str(FIXTURE))
    matches, demoted = detect_headings(doc)
    front = detect_front_matter_labels(doc)
    title = detect_document_title(doc, matches, front)
    en = detect_english_title(doc, front)
    ensub = detect_english_subtitle(doc, en,
                                    {m.paragraph_index for m in matches if m is not None}
                                    | {f.paragraph_index for f in front})
    caps = detect_captions(doc, {m.paragraph_index for m in matches if m is not None})

    def ptext(i):
        return doc.paragraphs[i].text.strip()[:40] if i is not None else "—"

    detail = ["不设封面范围："]
    detail.append("    题目   = 段 %s  %s" % (
        title.title_paragraph_index if title else None,
        ptext(title.title_paragraph_index) if title else "（没识别到）"))
    detail.append("    副标题 = 段 %s  %s" % (
        title.subtitle_paragraph_index if title else None,
        ptext(title.subtitle_paragraph_index) if title else "（没识别到）"))
    detail.append("    英文题目 = 段 %s  %s" % (en, ptext(en)))
    detail.append("    英文副标题 = 段 %s  %s" % (ensub, ptext(ensub)))
    detail.append("    识别到的标题 %d 个；图题/表题 %d 个；前置标签 %d 个；疑似标题 %d 个"
                  % (sum(1 for m in matches if m), len(caps), len(front), len(demoted)))
    for m in matches:
        if m:
            detail.append("      段 %-3d L%d 角色%-6s %s" % (
                m.paragraph_index, m.level, m.role or "-", m.title[:26]))
    for c in caps:
        detail.append("      段 %-3d %s题%s %s" % (
            c.paragraph_index, c.kind, "（英文）" if c.is_english else "", c.text[:26]))
    # 带封面范围再看一次
    if cover_end is not None:
        from essay_tool.engine.pipeline import recognize_source
        doc2 = Document(str(FIXTURE))
        (m2, d2, f2, t2, c2, e2, es2) = recognize_source(doc2, cover_end)
        detail.append("设封面范围 cover_end=%s 之后：" % cover_end)
        detail.append("    题目   = 段 %s  %s" % (
            t2.title_paragraph_index if t2 else None,
            t2.title_paragraph_index and
            Document(str(FIXTURE)).paragraphs[t2.title_paragraph_index].text.strip()[:30]))
        detail.append("    副标题 = 段 %s" % (t2.subtitle_paragraph_index if t2 else None))
        detail.append("    英文题目 = 段 %s；图题/表题 %d 个" % (e2, len(c2)))
    obs("O01", "识别结果（事实，不下判断）", "\n".join(detail))
except Exception as exc:                                          # noqa: BLE001
    skip("O01", "识别结果", "%s: %s" % (type(exc).__name__, exc))

# ------------------------------------------------------------------ O02/O03 告警
try:
    preset = load_preset(PRESET_PATH)
    kw = {"cover_end": cover_end}
    alert_with_cover, ignored_r = call_supported(review, FIXTURE, preset, **kw)
    alert_no_cover, _ = call_supported(review, FIXTURE, preset)
    from essay_tool.engine.alerts import AlertCollector
    collector = AlertCollector()
    out_dir = WORK / "run-alerts"
    out_dir.mkdir(parents=True, exist_ok=True)
    call_supported(run, FIXTURE, None, out_dir, preset_config=preset,
                   cover_end=cover_end, alerts=collector)
    rev_lines = [a.line for a in alert_with_cover.items]
    run_lines = [a.line for a in collector.items]
    doc_for_text = Document(str(FIXTURE))

    def explain(alerts_obj):
        out = []
        for a in alerts_obj.items:
            texts = []
            for idx in a.paragraph_indices[:2]:
                if 0 <= idx < len(doc_for_text.paragraphs):
                    texts.append("段%d=%s" % (idx, doc_for_text.paragraphs[idx].text.strip()[:26]))
            out.append("    · %s%s" % (a.line, ("   ← " + " | ".join(texts)) if texts else ""))
        return "\n".join(out) or "    （无告警）"

    obs("O02", "打开结构页（review）报的告警 —— 设封面范围=%s" % cover_end,
        explain(alert_with_cover))
    obs("O02b", "打开结构页报的告警 —— 不设封面范围", explain(alert_no_cover))

    # 封面区（段号 <= cover_end）里的内容不该出现在告警里——除非用户没设封面范围
    if cover_end is not None:
        inside = [(a.name, a.paragraph_indices) for a in alert_with_cover.items
                  if any(i <= cover_end for i in a.paragraph_indices)]
        if inside:
            obs("O09", "⚠ 封面区里的内容被算进了告警（cover_end=%s）" % cover_end,
                "\n".join("    · %s 报在段 %s" % (n, list(ix)) for n, ix in inside))
        else:
            obs("O09", "封面区里的内容没有被算进告警", "cover_end=%s 生效" % cover_end)
    only_rev = [x for x in rev_lines if x not in run_lines]
    only_run = [x for x in run_lines if x not in rev_lines]
    if only_rev or only_run:
        obs("O03", "review 与 run 的告警口径不一致",
            "只在 review 里：%s\n    只在 run 里：%s" % (only_rev or "—", only_run or "—"))
    else:
        obs("O03", "review 与 run 的告警口径一致", "两边都是 %d 条，逐条相同" % len(rev_lines))
except Exception as exc:                                          # noqa: BLE001
    skip("O02/O03", "告警清单", "%s: %s" % (type(exc).__name__, exc))

# ------------------------------------------------------------------ O04 结构页铺行
try:
    from essay_tool.gui.adjust_preview import build_rows
    rows_out, ignored = call_supported(build_rows, FIXTURE, cover_end=cover_end)
    head = rows_out[:12]
    detail = "共铺 %d 行；前 12 行：" % len(rows_out)
    for r in head:
        detail += "\n      段 %-4s kind=%-10s block=%-12s %s" % (
            getattr(r, "doc_index", "?"), getattr(r, "kind", "?"),
            getattr(r, "block", "") or "-", str(getattr(r, "text", ""))[:34])
    cover_rows = [r for r in rows_out if getattr(r, "doc_index", None) is not None
                  and cover_end is not None and r.doc_index <= cover_end]
    detail += "\n    封面区（段 0~%s）里铺出来的行数 = %d" % (cover_end, len(cover_rows))
    obs("O04", "结构页是否铺出封面区的行", detail)
except Exception as exc:                                          # noqa: BLE001
    skip("O04", "结构页铺行", "%s: %s" % (type(exc).__name__, exc))

# ------------------------------------------------------------------ O05/O06/O07 单项
def has_alert(collector, name: str) -> bool:
    return any(a.name == name for a in collector.items)


try:
    preset = load_preset(PRESET_PATH)
    alerts, _ = call_supported(review, FIXTURE, preset, cover_end=cover_end)
    checks = [
        ("未接受的修订", "未接受的修订告警（夹具里有 w:ins/w:del）"),
        ("未编号的图片/表格", "未编号图片/表格告警（夹具里有 1 张无题注图 + 2 张无题注表）"),
        ("疑似标题", "疑似标题告警（夹具里有第一章/①/1) 三种非规范编号）"),
        ("跳号", "跳号告警（夹具里 三、之后是 六、）"),
        ("未归零", "未归零告警（夹具里 3. 出现在新的二级标题之后）"),
        ("孤立", "孤立告警（夹具里有单条三级/四级标题）"),
        ("多余的正文", "多余的正文告警（目录/清单标签下有正文）"),
        ("重复的部件", "重复的部件告警（夹具里只有一个摘要，预期不报）"),
        ("不存在的清单对象", "不存在的清单对象告警（【附表清单】启用但无表题）"),
    ]
    detail = []
    for name, desc in checks:
        detail.append("    %-14s %s  ← %s" % (name, "报了" if has_alert(alerts, name) else "没报", desc))
    obs("O05", "各类告警的报出情况（事实，需你判断该不该报）", "\n".join(detail))
except Exception as exc:                                          # noqa: BLE001
    skip("O05", "告警报出情况", "%s: %s" % (type(exc).__name__, exc))

try:
    from essay_tool.engine.detect import detect_headings
    doc = Document(str(FIXTURE))
    matches, _d = detect_headings(doc)
    auto = [i for i, p in enumerate(doc.paragraphs) if "自动编号的" in p.text]
    detail = []
    for i in auto:
        m = matches[i]
        pPr = doc.paragraphs[i]._p.find(qn("w:pPr"))
        numpr = pPr is not None and pPr.find(qn("w:numPr")) is not None
        detail.append("    段 %d：%s" % (i, doc.paragraphs[i].text.strip()[:26]))
        detail.append("        识别结果=%s；auto_numbered=%s；段上有 numPr=%s" % (
            ("L%d" % m.level) if m else "不是标题",
            getattr(m, "auto_numbered", "-") if m else "-", numpr))
    styled = [i for i, p in enumerate(doc.paragraphs) if "List Number 样式" in p.text]
    for i in styled:
        pPr = doc.paragraphs[i]._p.find(qn("w:pPr"))
        numpr = pPr is not None and pPr.find(qn("w:numPr")) is not None
        detail.append("    段 %d：%s（样式挂编号）" % (i, doc.paragraphs[i].text.strip()[:26]))
        detail.append("        识别结果=%s；段上有 numPr=%s ← 识别只看段落自己的 numPr"
                      % ("L%d" % matches[i].level if matches[i] else "不是标题", numpr))
    obs("O06", "Word 自动编号段落：识别结果与 numPr 现状", "\n".join(detail) or "    （没找到）")
except Exception as exc:                                          # noqa: BLE001
    skip("O06", "自动编号段落", "%s: %s" % (type(exc).__name__, exc))

# ------------------------------------------------------------------ H13 封面区不参与告警
# 用户 2026-09-15 裁决：「一旦已经被识别为封面，那就不应该管，也不算进清单里，
# 原则上也不应该告警。」
try:
    preset = load_preset(PRESET_PATH)
    alerts_cover, _ignored = call_supported(review, FIXTURE, preset, cover_end=cover_end)
    inside = []
    for a in alerts_cover.items:
        hit = [i for i in a.paragraph_indices
               if cover_end is not None and i <= cover_end]
        if hit:
            inside.append("%s（报在段 %s）" % (a.line, hit))
    if inside:
        fail("H13", "封面区不参与告警 / 不算进清单",
             "下列告警指向了封面区（段 0~%s），按裁决应当整条不报：\n        %s"
             % (cover_end, "\n        ".join(inside)))
    else:
        ok("H13", "封面区不参与告警 / 不算进清单", "cover_end=%s 生效" % cover_end)
except Exception as exc:                                          # noqa: BLE001
    fail("H13", "封面区不参与告警", "%s: %s" % (type(exc).__name__, exc))

# ------------------------------------------------------------------ H14 同类告警合并
# 用户 2026-09-15 裁决：「反正所有问题合着只报一次，段号并列，就不用管了。」
try:
    preset = load_preset(PRESET_PATH)
    alerts_c, _ignored = call_supported(review, FIXTURE, preset, cover_end=cover_end)
    groups: dict[tuple, list] = {}
    for a in alerts_c.items:
        groups.setdefault((a.name, a.message), []).append(a.paragraph_indices)
    dup = {key: value for key, value in groups.items() if len(value) > 1}
    if dup:
        detail = "有 %d 组同类告警没合（同一名称+同一概要出现了多次）：" % len(dup)
        for (name, msg), value in dup.items():
            detail += "\n        %s：%s × %d 条" % (name, msg, len(value))
        detail += "\n    按裁决合并后应当长这样："
        for (name, msg), value in dup.items():
            flat = [i for group in value for i in group]
            shown = "、".join(str(i) for i in flat[:8]) + ("…" if len(flat) > 8 else "")
            detail += "\n        · %s：%s，在段落（%s）。" % (name, msg, shown)
        fail("H14", "同类告警合并成一条、段号并列", detail)
    else:
        ok("H14", "同类告警合并成一条、段号并列",
           "共 %d 条告警，没有同一问题被拆成多条" % len(alerts_c.items))
except Exception as exc:                                          # noqa: BLE001
    fail("H14", "同类告警合并成一条、段号并列", "%s: %s" % (type(exc).__name__, exc))

# ------------------------------------------------------------------ H15 Word 自动编号
# 夹具里两段标题的编号只存在于 numbering.xml（段落文字里什么都没有），
# 导出时必须把编号固化成真文字、并摘掉 numPr，否则清直接格式会把编号弄丢。
try:
    if "cover" in outputs and outputs["cover"].exists():
        src2 = Document(str(FIXTURE))
        out_doc = Document(str(outputs["cover"]))

        def find_para(doc, marker):
            for p in doc.paragraphs:
                if marker in p.text:
                    return p
            return None

        wants = (("zmarkW1", "一、"), ("zmarkW2", "（一）"))
        bad = []
        detail = []
        for marker, want in wants:
            src_p = find_para(src2, marker)
            out_p = find_para(out_doc, marker)
            if src_p is None:
                bad.append("%s 在夹具里就不存在" % marker)
                continue
            if out_p is None:
                bad.append("%s 在产物里丢了" % marker)
                continue
            text = out_p.text.strip()
            pPr = out_p._p.find(qn("w:pPr"))
            numpr = pPr is not None and pPr.find(qn("w:numPr")) is not None
            if not text.startswith(want):
                bad.append("%s 的编号没固化进文字（期望以 %r 开头，实际 %r）"
                           % (marker, want, text[:22]))
            if numpr:
                bad.append("%s 还挂着 numPr（编号会被 Word 再画一遍）" % marker)
            detail.append("    %s → 产物文字 %r；带 numPr=%s" % (marker, text[:26], numpr))
        if bad:
            fail("H15", "Word 自动编号必须固化进文字",
                 "；".join(bad) + "\n" + "\n".join(detail))
        else:
            ok("H15", "Word 自动编号必须固化进文字", "\n" + "\n".join(detail))
    else:
        skip("H15", "Word 自动编号必须固化进文字", "cover 那次导出没成功")
except Exception as exc:                                          # noqa: BLE001
    fail("H15", "Word 自动编号必须固化进文字", "%s: %s" % (type(exc).__name__, exc))

# ------------------------------------------------------------------ H16 修订只在封面区里
# 反向验证 H13 的过滤：封面区里带修订、封面之后不带 → 不报；不设封面范围 → 照报。
try:
    tmp = WORK / "fixture-修订只在封面.docx"
    d = Document()
    p0 = d.add_paragraph("封面里带着未接受的修订的一段")
    ins = OxmlElement("w:ins")
    ins.set(qn("w:id"), "91")
    ins.set(qn("w:author"), "Agent")
    ins.set(qn("w:date"), "2026-09-15T00:00:00Z")
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "封面里的插入内容"
    r.append(t)
    ins.append(r)
    p0._p.append(ins)
    d.add_paragraph("封面之后的一段普通正文，没有任何修订。")
    d.save(str(tmp))
    preset = load_preset(PRESET_PATH)

    def has_revision_alert(alerts_obj):
        return any(a.name == "未接受的修订" for a in alerts_obj.items)

    with_cover, _ig = call_supported(review, tmp, preset, cover_end=0)
    without_cover, _ig2 = call_supported(review, tmp, preset)
    detail = "设 cover_end=0：%s；不设封面范围：%s" % (
        "报了" if has_revision_alert(with_cover) else "没报",
        "报了" if has_revision_alert(without_cover) else "没报")
    if not has_revision_alert(with_cover) and has_revision_alert(without_cover):
        ok("H16", "修订只在封面区里时不报（且不设封面时仍报）", detail)
    else:
        fail("H16", "修订只在封面区里时不报（且不设封面时仍报）",
             detail + " ← 期望：设了封面＝没报、不设封面＝报了")
except Exception as exc:                                          # noqa: BLE001
    fail("H16", "修订只在封面区里时不报", "%s: %s" % (type(exc).__name__, exc))

# ------------------------------------------------------------------ H17 手打目录条目不当标题
# 原来是观察项 O10（"待你拍板"）。用户 2026-09-15 拍板要修，插件已按窄判据
# 实现（`detect.hand_typed_toc_entries()`），所以升成硬断言。
#
# ⚠️ **判据自己重算一遍**：不调 `hand_typed_toc_entries()` 拿现成答案——那样
# 断言只是"函数等于它自己"，实现坏了它跟着一起坏。这里只按用户的话独立推：
# "紧贴【目录】（或清单）标签、连续带页码尾巴的行"不算标题。页码尾巴＝结尾是
# 省略号/三个点/制表符再加数字。
try:
    src2 = Document(str(FIXTURE))
    matches, _demoted = detect_headings(src2)
    paras = src2.paragraphs

    def is_toc_page_label(text: str) -> bool:
        return text.strip().strip("【】[]") in (
            "目录", "插图清单", "附表清单", "插图和附表清单")

    def is_part_label(text: str) -> bool:
        # 下一个部件的开始（前置标签或引言/结论/参考文献/附录/致谢）
        return is_toc_page_label(text) or text.strip().replace(" ", "").replace("\u3000", "") in (
            "摘要", "关键词", "Abstract", "Keywords",
            "引言", "结论", "参考文献", "致谢", "附录")

    def has_page_tail(text: str) -> bool:
        return re.search(r"(?:…{2,}|\.{3,}|\t)\s*\d+\s*$", text) is not None

    entries: list[int] = []
    for i, p in enumerate(paras):
        if not is_toc_page_label(p.text):
            continue
        for j in range(i + 1, len(paras)):
            line = paras[j].text.strip()
            if not line:
                continue                     # 空行只当排版空隙
            if is_part_label(line):
                break                        # 到下一个部件了
            if not has_page_tail(line):
                break                        # 连续页码链断了（真标题就断在这）
            entries.append(j)
    bad_titles = ["    段 %d：%s → 被识别为标题（L%d）"
                  % (i, paras[i].text.strip()[:34], matches[i].level)
                  for i in entries if matches[i] is not None]
    if not entries:
        # 夹具本身没有这种段落时不能算"通过"——那是断言失效，不是行为正确
        fail("H17", "手打目录条目不被当成标题",
             "夹具里找不到「紧贴目录标签 + 连续页码尾巴」的段落，这条断言等于没跑")
    elif bad_titles:
        fail("H17", "手打目录条目不被当成标题",
             "共 %d 段在页码链上，其中 %d 段被当成标题：\n%s"
             % (len(entries), len(bad_titles), "\n".join(bad_titles)))
    else:
        ok("H17", "手打目录条目不被当成标题",
           "紧贴目录标签、连续带页码尾巴的 %d 段（段 %s），全部没有被认成标题"
           % (len(entries), "、".join(str(i) for i in entries)))
except Exception as exc:                                          # noqa: BLE001
    fail("H17", "手打目录条目不被当成标题", "%s: %s" % (type(exc).__name__, exc))

# ------------------------------------------------------------------ 汇总 + 落盘
passed = sum(1 for r in results if r[2] == "PASS")
failed = sum(1 for r in results if r[2] == "FAIL")
skipped = sum(1 for r in results if r[2] == "SKIP")
observed = sum(1 for r in results if r[2] == "OBS")

say()
say("=" * 100)
say("结论汇总：硬断言 通过 %d / 不通过 %d ；跳过 %d ；观察项 %d" % (passed, failed, skipped, observed))
say("=" * 100)
say()
say("【硬断言】失败＝真 Bug")
for cid, title, status, detail in results:
    if status in ("PASS", "FAIL", "SKIP"):
        mark = {"PASS": "PASS", "FAIL": "FAIL", "SKIP": "SKIP"}[status]
        say("[%s] %-8s %s" % (cid, mark, title))
        if detail:
            say("        " + detail)
say()
say("【观察项】只报事实，该不该修由你判断")
for cid, title, status, detail in results:
    if status == "OBS":
        say("[%s] %s" % (cid, title))
        say(detail)

REPORT.write_text("\n".join(lines), encoding="utf-8")

print("REPORT: %s" % REPORT)
print("HARD: pass=%d fail=%d skip=%d   OBS=%d" % (passed, failed, skipped, observed))
for cid, title, status, detail in results:
    if status == "FAIL":
        first = detail.splitlines()[0] if detail else ""
        ascii_detail = first.encode("ascii", "replace").decode("ascii")
        print("  FAIL [%s] %s :: %s" % (cid, ascii_detail[:80], ""))
