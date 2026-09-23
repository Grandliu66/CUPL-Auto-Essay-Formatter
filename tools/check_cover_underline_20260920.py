"""封面超长值折行时，下划线右端是否齐平的检查（需要本机装有 Word；用 Word 隐藏导出 PDF 后
读矢量线段的坐标，不弹界面）。

论文封面：学院栏填 10～45 个汉字，折出的每一行下划线右端必须相同，且行数不许比旧做法更多。
课程封面：课程名称超长时，两行下划线右端必须相同。

跑法：.venv\\Scripts\\python.exe tools\\check_cover_underline_20260920.py    （退出码 0＝全部通过；没有 Word 则跳过）
"""

import sys, tempfile, pathlib
sys.stdout.reconfigure(encoding="utf-8")
ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
try:
    import win32com.client, pymupdf
except ImportError:
    print("缺 pywin32/pymupdf，跳过"); sys.exit(0)
from essay_tool.engine import cover as C

bad = 0
def check(n, c, extra=""):
    global bad
    print(("OK   " if c else "FAIL ") + n, "" if c else extra); bad += (not c)

T = pathlib.Path(tempfile.mkdtemp())
try:
    word = win32com.client.DispatchEx("Word.Application"); word.Visible = False; word.DisplayAlerts = 0
except Exception:
    print("没有 Word，跳过"); sys.exit(0)

def underlines(docx, x_max, y_min, y_max=10_000):
    doc = word.Documents.Open(str(docx), ReadOnly=True, AddToRecentFiles=False, Visible=False)
    pdf = docx.with_suffix(".pdf"); doc.ExportAsFixedFormat(str(pdf), 17); doc.Close(False)
    pg = pymupdf.open(str(pdf))[0]
    return sorted((round(d["rect"].y0), round(d["rect"].x1, 1)) for d in pg.get_drawings()
                  if d["rect"].height < 2 and d["rect"].width > 20 and d["rect"].x0 < x_max
                  and y_min < d["rect"].y0 < y_max)

BASE = "国际法学院涉外法治研究所培养项目试点计划综合改革试验区示范班次国际法学院涉外法治"
orig = C._flush_underline_right
try:
    for n in (9, 10, 13, 17, 18, 19, 24, 30, 45):
        f = C.CoverFields(title="题", student_name="张三", student_id="2025030001", college=BASE[:n],
                          grade="2025级", date="2026年9月", major="法学", advisor_name="张三", advisor_title="教授")
        res = {}
        for mode in ("old", "new"):
            C._flush_underline_right = (lambda *a: None) if mode == "old" else orig
            p = T / f"t{n}_{mode}.docx"; C.build_thesis_cover_document(f).save(str(p))
            res[mode] = underlines(p, 300, 620)[:-2]        # 左栏：学院各行（去掉其后的年级、日期）
        ends = {x for _, x in res["new"]}
        check(f"论文封面 学院{n}字：各行下划线右端齐平、行数不多于旧做法",
              len(ends) == 1 and len(res["new"]) <= len(res["old"]), f"新={res['new']} 旧={res['old']}")
finally:
    C._flush_underline_right = orig

f = C.CoverFields(course_name="民法学原理与比较法视野下的合同制度研究专题讨论课程综合考核（春季学期）",
                  course_code="LAW2025001", course_seq="01", teacher_name="张老师", student_name="李四", student_id="2025030001")
p = T / "course.docx"; C.build_course_cover_document(f).save(str(p))
seg = underlines(p, 10_000, 200, 300)          # 课程名称占页面 y≈258、≈290 两行
check("课程封面 课程名称超长：两行下划线右端齐平", len(seg) >= 2 and seg[0][1] == seg[1][1], str(seg))
word.Quit()
print("全部通过" if not bad else f"{bad} 项失败")
sys.exit(1 if bad else 0)
