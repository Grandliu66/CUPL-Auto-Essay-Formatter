"""生成覆盖「表格文字 + 图题/表题 + 脚注」的测试文件。

为什么拿 `PaperTest.docx` 当底、而不是从零新建：**脚注**要 `word/footnotes.xml`
这个独立部件，外加 [Content_Types]、关系项、`w:footnoteReference` 一整套样板，
从零构造纯属自找麻烦。`PaperTest.docx` 里本来就有**真的脚注部件**和**真的嵌入
图片**，在它基础上补表格和图题最省事。

产出里刻意埋了几个陷阱（对应识别要能"宽进严出"）：
  * 手打的图题 / 表题（弱信号，靠文字特征认）
  * 用"正规做法"插的题注（题注样式 + SEQ 域，强信号）
  * 以标点结尾的「图 1 显示了这个结果，……」——**明显是正文**，不能认成图题
  * 超长的「表 1 …」——**明显是正文**，不能认成表题
  * 「表 1（续）」「图 1（a）分图标记测试」——**不是新题注**，不能自动编号
    （编错号会把"（续）"顶成"表 2（续）"）、不能进清单

用法：.venv\\Scripts\\python.exe tools\\make_fixture_charts.py
产物：reference\\图表脚注测试.docx
（这个文件是我们自己生成的测试样本，可以随代码重新生成；reference 下别的
文件仍是只读的。）
"""

import io
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# 路径按**脚本位置**推（审计 2026-09-15 核实：原来是相对路径 `reference/…`，
# 换个工作目录跑就找不到夹具了；同目录别的工具脚本都是这么写的）。
SRC = ROOT / "reference" / "PaperTest.docx"
OUT = ROOT / "reference" / "图表脚注测试.docx"


def add_seq_caption(document, label: str, number: str, text: str) -> None:
    """模拟 Word「插入题注」：题注样式 + SEQ 域。"""
    paragraph = document.add_paragraph()
    try:
        paragraph.style = document.styles["Caption"]
    except KeyError:
        paragraph.style = document.styles.add_style("Caption", WD_STYLE_TYPE.PARAGRAPH)
    paragraph.add_run(label + " ")
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), " SEQ %s \\* ARABIC " % label)
    inner_run = OxmlElement("w:r")
    inner_t = OxmlElement("w:t")
    inner_t.text = number
    inner_run.append(inner_t)
    fld.append(inner_run)
    paragraph._p.append(fld)
    paragraph.add_run(" " + text)


def image_parts(document):
    return [p for p in document.part.package.iter_parts() if p.content_type.startswith("image/")]


def main() -> None:
    document = Document(SRC)
    document.add_paragraph()

    document.add_paragraph("下面开始是图题/表题/表格的测试区。")

    # ① 手打图题 + 真图片
    document.add_paragraph("图 1 手打的图题")
    parts = image_parts(document)
    if parts:
        document.add_picture(io.BytesIO(parts[0].blob))
    else:
        print("  ！PaperTest 里没有图片部件，跳过插图")

    # ①b 英文图题：规范要求"英文图序与图名另起一行放中文下方"。注意它下面
    #     隔了一段图片（无文字的段落），识别时会跳过空段往上找。
    document.add_paragraph("Fig 1 Manual caption in English")

    # ② 手打表题 + 表格
    document.add_paragraph("表 1 手打的表题")
    table = document.add_table(rows=2, cols=3)
    for col, head in enumerate(("表头一", "表头二", "表头三")):
        table.cell(0, col).text = head
    for col, body in enumerate(("内容一", "内容二", "内容三")):
        table.cell(1, col).text = body

    # ②b 英文表题
    document.add_paragraph("Table 1 Manual table caption")

    # ③ 正规做法：题注样式 + SEQ 域
    add_seq_caption(document, "图", "2", "用题注插的图名")

    # ④ 陷阱：以标点结尾 → 明显是正文
    document.add_paragraph("图 1 显示了这个结果，但这一句是正文，因为它以标点结尾。")

    # ⑤ 陷阱：过长 → 明显是正文
    document.add_paragraph("表 1 " + "这一段其实是正文，只是碰巧以表加数字开头。" * 12)

    # ⑤b 陷阱：英文正文。`Table 2 gives the summary of all measured values`
    #     不以标点结尾、也不过长，光靠中文字题那两条排除拦不住——靠的是
    #     "上面必须是中文表题"这条硬条件（它上面不是）。
    document.add_paragraph("Table 2 gives the summary of all measured values")

    # ⑤c 陷阱：「（续）」和「分图」标记行——它们**不是新的一张图/表**：
    #     不识别、不套样式、不自动编号、不进清单。2026-09-12 用户问
    #     "手动（续）/分图还能不能用"时实测到的坑：不排除的话它们会被编成
    #     "表 2（续）""图 3（a）"，还会误报"编号不连续"。**必须保持正文**。
    document.add_paragraph("表 1（续）")
    document.add_paragraph("图 1（a）分图标记测试")

    # ⑥ 附录：完整的三段结构（标题 / 序号题名 / 正文）。两个目的：
    #    a) 验规范 [225] 的「附录标题」是不是建成了**独立样式**——以前把
    #       "黑体"直接写在段落上，Word 生成目录时连字体一起复制，目录里
    #       那条也变黑体（用户发现的 bug）。
    #    b) **底稿 PaperTest.docx 里本来就有一个附录**，所以成品里会有两个
    #       ——正好当"多个相同结构"的用例（规范 [226] 明确允许多个附录，
    #       且要求序列靠后的另起一页）。看到两条"附　录"不是重复插入。
    document.add_paragraph("附　录")
    document.add_paragraph("附录A  关于图题表题的说明")
    document.add_paragraph("这里是附录正文，格式应当跟正文一致。")

    document.save(OUT)
    print("已生成：%s" % OUT)


if __name__ == "__main__":
    main()
