"""清理规则的引擎检查：清除手打分页符、清除空行（严格空段、表格边界、图片/编号/分节符/换行不动）、
封面范围不受影响、标签前另起一页覆盖（含结构页路径）、成品回炉稳定。

跑法：.venv\\Scripts\\python.exe tools\\check_cleanup_engine_20260920.py    （退出码 0＝全部通过）
"""

import sys, pathlib, tempfile, dataclasses
sys.stdout.reconfigure(encoding="utf-8")
ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from docx import Document
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import qn, nsdecls
from PIL import Image
from essay_tool.engine.pipeline import run
from essay_tool.engine.alerts import AlertCollector
from essay_tool.config.preset_io import load_preset

bad = 0
def check(n, c, extra=""):
    global bad
    print(("OK   " if c else "FAIL ") + n, "" if c else extra); bad += (not c)

T = pathlib.Path(tempfile.mkdtemp())
BASE = load_preset(ROOT / "presets" / "中国政法大学-本科毕业论文.json")
W = nsdecls("w")

def brk(p, first=False):
    r = parse_xml(f'<w:r {W}><w:br w:type="page"/></w:r>')
    runs = p._p.findall(qn("w:r"))
    (runs[0].addprevious(r) if first and runs else p._p.append(r))

def build(name="rules"):
    """返回 (路径, 关键段落序号表)"""
    img = T / "px.png"; Image.new("RGB", (40, 20), "red").save(img)
    d = Document()
    P = lambda t="": d.add_paragraph(t)
    P("论文题目")                          # 0  ┐ 封面范围（cover_end=3）
    P("")                                 # 1  │ 封面里的空行
    b = P(""); brk(b)                     # 2  │ 封面里的分页符
    P("封面最后一行")                      # 3  ┘
    P("摘要"); P("摘要正文。")              # 4 5
    P("")                                 # 6  空行（应删）
    P("")                                 # 7  空行（应删）
    P("\u3000 ")                          # 8  空白字符（严格：不删）
    b = P(""); brk(b)                     # 9  只有分页符的段
    P("一、绪论")                          # 10
    p = P("正文一。"); brk(p)              # 11 文字+分页符
    pi = P(""); pi.add_run().add_picture(str(img))   # 12 图片独占段（不删）
    P("")                                 # 13 空行（应删）
    pn = P(""); ppr = pn._p.get_or_add_pPr()
    ppr.append(parse_xml(f'<w:numPr {W}><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'))   # 14 空编号项（不删）
    pl = P("含换行"); pl.add_run().add_break()       # 15
    P("二、分析")                          # 16
    t1 = d.add_table(rows=1, cols=1); t1.cell(0, 0).text = "表一"
    P("")                                 # 表间空行（两表之间：留）
    t2 = d.add_table(rows=1, cols=1); t2.cell(0, 0).text = "表二"
    P("")                                 # 表后空行（还有后文：可删）
    P("正文二。")
    P("参考文献"); P("[1] 某某. 论文. 2020.")
    P("")                                 # 参考文献里的空行（应删）
    P("[2] 某某. 论文二. 2021.")
    t3 = d.add_table(rows=1, cols=1); t3.cell(0, 0).text = "文末表"
    P("")                                 # 文末表格之后的空段（必须留）
    src = T / f"{name}.docx"; d.save(str(src)); return src

def go(src, tag, **cfg):
    kw = dict(cover_end=3) if cfg.pop("with_cover", None) is not None else {}
    preset = dataclasses.replace(BASE, **cfg)
    return Document(str(run(src, None, T / f"w_{tag}", alerts=AlertCollector(), preset_config=preset,
                            check_source_alerts=False, **kw)))

def stats(doc):
    ps = doc.paragraphs
    n_br = sum(1 for p in ps for b in p._p.iter(qn("w:br")) if b.get(qn("w:type")) == "page")
    n_br += sum(1 for t in doc.tables for row in t.rows for c in row.cells for p in c.paragraphs
                for b in p._p.iter(qn("w:br")) if b.get(qn("w:type")) == "page")
    blanks = sum(1 for p in ps if not p.text and not p._p.findall(".//" + qn("w:drawing"))
                 and p._p.find(qn("w:pPr")) is not None or (not p.text and p._p.find(qn("w:pPr")) is None))
    return len(ps), n_br

src = build()
base = Document(str(run(src, None, T / "w0", alerts=AlertCollector(), preset_config=BASE, check_source_alerts=False)))
n0, br0 = stats(base)
print(f"基线（两项都关）：{n0} 段，手打分页符 {br0} 个")
check("基线：分页符都在（源文档 3 个）", br0 == 3, br0)

# ---------- A 只清分页符
a = go(src, "A", clear_page_breaks=True)
nA, brA = stats(a)
check("A 清除分页符：成品里一个分页符都不剩", brA == 0, brA)
check("A 段落数不变（分页符单独成段的那些变成了空段，仍在）", nA == n0, (nA, n0))
check("A 文字内容一个字没少", [p.text for p in a.paragraphs if p.text.strip()] == [p.text for p in base.paragraphs if p.text.strip()])

# ---------- B 只清空行
b = go(src, "B", clear_blank_lines=True)
nB, brB = stats(b)
texts_b = [p.text for p in b.paragraphs]
check("B 清除空行：段落变少了", nB < n0, (nB, n0))
check("B 分页符没被清（勾的是空行）", brB == br0, brB)
check("B 只含分页符的段落保留（它不是空的）", any((not p.text) and any(x.get(qn('w:type')) == 'page' for x in p._p.iter(qn('w:br'))) for p in b.paragraphs))
check("B 全角空格+空格的段落保留（严格：有字符就不算空）", any(t == "\u3000 " for t in texts_b), texts_b)
check("B 图片独占段保留", sum(1 for p in b.paragraphs if p._p.findall('.//' + qn('w:drawing'))) == 1)
check("B 空编号项保留", any(p._p.find(qn('w:pPr')) is not None and p._p.find(qn('w:pPr')).find(qn('w:numPr')) is not None and not p.text for p in b.paragraphs))
check("B 含换行(Shift+Enter)的段落保留", any(p.text.startswith("含换行") for p in b.paragraphs))
# 参考文献里的空行：三条之间不再有空段
refs = [i for i, p in enumerate(b.paragraphs) if "某某" in p.text]
check("B 参考文献条目之间的空行被清掉", len(refs) == 2 and refs[1] - refs[0] == 1, refs)
# 表格结构：两表之间仍隔着一个段落；文末表格后仍有段落
body = list(b.element.body)
tags = [("t" if e.tag == qn("w:tbl") else "p" if e.tag == qn("w:p") else "s") for e in body]
seq = "".join(tags)
check("B 两张表格之间没有被并成一张（中间仍有段落）", "tt" not in seq, seq)
check("B 文末表格之后仍有段落（正文以段落收尾）", seq.rstrip("s")[-1] == "p", seq[-8:])

# ---------- C 两项都开
c = go(src, "C", clear_page_breaks=True, clear_blank_lines=True)
nC, brC = stats(c)
check("C 两项都开：分页符全清、且单独成段的分页符段也被当空行删掉", brC == 0 and nC < nB, (nC, nB))

# ---------- D 封面范围不受影响
d_ = go(src, "D", clear_page_breaks=True, clear_blank_lines=True, with_cover=True)
ps = d_.paragraphs
check("D 封面范围（前 4 段）原样：空行还在", ps[0].text == "论文题目" and ps[1].text == "" and ps[3].text == "封面最后一行", [p.text for p in ps[:4]])
check("D 封面里的分页符还在", any(x.get(qn('w:type')) == 'page' for x in ps[2]._p.iter(qn('w:br'))))
_, brD = stats(d_)
check("D 封面之外的分页符仍被清掉（只剩封面里那 1 个）", brD == 1, brD)
check("D 封面之外的空行仍被清掉", len(d_.paragraphs) < n0)

# ---------- E 标签另起一页覆盖
def pbb_direct(p):
    ppr = p._p.find(qn("w:pPr"))
    e = ppr.find(qn("w:pageBreakBefore")) if ppr is not None else None
    return None if e is None else (e.get(qn("w:val")) not in ("0", "false", "off"))
def find(doc, text): return next(p for p in doc.paragraphs if p.text.strip() == text)
e0 = base
print("基线：摘要标签直接段前分页 =", pbb_direct(find(e0, "【摘要】")), "；参考文献标签直接 =", pbb_direct(find(e0, "参考文献")), "；关键词(生成)=", pbb_direct(find(e0, "【关键词】")))
e = go(src, "E", label_page_break={"摘要": True, "参考文献": False, "关键词": True})
check("E 摘要标签（原本不分页）→ 另起一页", pbb_direct(find(e, "【摘要】")) is True)
check("E 参考文献标签（一级标题样式本来分页）→ 显式不分页", pbb_direct(find(e, "参考文献")) is False)
check("E 生成出来的关键词标签（原文没有）也吃到覆盖", pbb_direct(find(e, "【关键词】")) is True)
check("E 没写覆盖的部件不受影响（英文摘要标签）", pbb_direct(find(e, "【ABSTRACT】")) == pbb_direct(find(base, "【ABSTRACT】")))

# ---------- F 结构页路径（rows→ops）同样生效
sys.path.insert(0, str(ROOT / "tools"))
from essay_tool.gui import adjust_preview as ap
from essay_tool.engine.doc_index import build_ops, body_list_field_kinds, body_toc_field_parts
from essay_tool.engine.pipeline import collect_source_alerts
def gui_run(src, preset):
    rows, index = ap.build_rows(src)
    ap.seed_enabled_from(preset.structure_required, rows)
    doc = Document(str(src))
    _c, plan, _cap = collect_source_alerts(document=doc, index=index, structure_required=preset.structure_required,
        existing_list_kinds=body_list_field_kinds(doc), cover_end=None, regenerate_fields=frozenset(), include_source_checks=False)
    ops = build_ops(rows, index, enabled=dict(ap.ENABLED), insert_toc=("目录" not in body_toc_field_parts(doc)), list_kinds=plan)
    return Document(str(run(src, None, T / "wg", alerts=AlertCollector(), index=index, ops=ops, preset_config=preset, check_source_alerts=False)))
g = gui_run(src, dataclasses.replace(BASE, clear_page_breaks=True, clear_blank_lines=True, label_page_break={"摘要": True}))
check("F 结构页路径：分页符全清、空行被清", stats(g)[1] == 0 and len(g.paragraphs) < n0)
check("F 结构页路径：标签另起一页覆盖生效", pbb_direct(find(g, "【摘要】")) is True)
check("F 结构页路径：占位语没有重复", len([p for p in g.paragraphs if p.text.strip() == "［请在此处填写摘要］"]) <= 1)
# ---------- G 再导一遍稳定
g2 = Document(str(run(T / "w_C" / next(iter((T / "w_C").glob("*.docx"))).name, None, T / "w_C2", alerts=AlertCollector(),
                      preset_config=dataclasses.replace(BASE, clear_page_breaks=True, clear_blank_lines=True), check_source_alerts=False)))
check("G 成品再导一遍（两项都开）：段落数不变", len(g2.paragraphs) == len(c.paragraphs), (len(g2.paragraphs), len(c.paragraphs)))
print("失败:", bad)

sys.exit(1 if bad else 0)
