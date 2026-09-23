"""生成封面样张：调用 engine/cover.py 的组装函数，输出到 work\\ 供肉眼验收。

- 论文：`reference\\中国政法大学本科生毕业论文写作模板.docx` 第 1 页
  （文字区 + 两栏信息区）
- 课程：`reference\\中国政法大学本科生课程考核封面.docx` 第 1 页
  （"书写和装订要求"页不搬）

组装/清理/字面填写规则全在 `essay_tool/engine/cover.py`（学号不校验、姓名
两字空格、指导教师"姓名 职称"两个量、悬挂缩进……）。本脚本只负责摆样张：

1/2 = 论文 空白/填写；3/4 = 课程考核 空白/填写；
5   = 超长值压力测试（少数民族长名字 + 超长学院/专业，看悬挂缩进）。

用法：.venv\\Scripts\\python.exe tools\\make_cover_samples.py
"""

import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from essay_tool.engine.cover import (
    CoverFields,
    build_course_cover_document,
    build_thesis_cover_document,
)

OUT_DIR = ROOT / "work"


def write_cover(document, path: pathlib.Path) -> None:
    document.save(str(path))
    print("已写出", path.relative_to(ROOT))


def main() -> None:
    OUT_DIR.mkdir(exist_ok=True)

    # 演示数据（就是 GUI 以后让用户填的那些量）。填的都是**原始值**：
    # "张三"会按"两个字空格"规则写成"张 三"；日期留空 → 系统日期
    thesis_filled = CoverFields(
        number="2026-BK-001",
        title="论违约精神损害赔偿的适用边界",
        subtitle="以《民法典》第九百九十六条为视角",
        student_name="张三",
        student_id="2022101042",
        college="法学院",
        grade="2022级",
        date="",  # 留空 → 生成时取系统日期（二〇二六年九月）
        major="法学",
        advisor_name="李四",
        advisor_title="副教授",
    )
    course_filled = CoverFields(
        course_name="民法学原理",
        course_code="307010113",
        course_seq="42",
        teacher_name="李四",  # 课程考核默认不带职称
        student_name="张三",
        student_id="2022101042",
    )
    # 5 号：超长值／少数民族姓名压力测试——专看"值超过线宽"时的换行、
    # 悬挂缩进（续行对齐提示词右侧），以及带"·"名字的表现
    thesis_long = CoverFields(
        number="2026-BK-001",
        title="论违约精神损害赔偿的适用边界",
        subtitle="以《民法典》第九百九十六条为视角",
        student_name="阿依古丽·阿不都热依木",  # 22 个半角位：超出线宽 → 换行
        student_id="2022101042",
        college="国际法学院/涉外法治学院",  # 23 个半角位：超出线宽 → 换行
        grade="2022级",
        date="",
        major="法学（涉外法治人才培养实验班）",  # 30 个半角位：大幅超出 → 换行
        advisor_name="李四",
        advisor_title="副教授",
    )

    write_cover(
        build_thesis_cover_document(CoverFields(), verbose=True),
        OUT_DIR / "封面样例_1_论文_空白.docx",
    )
    write_cover(
        build_thesis_cover_document(thesis_filled, verbose=True),
        OUT_DIR / "封面样例_2_论文_填写.docx",
    )
    write_cover(
        build_course_cover_document(CoverFields(), verbose=True),
        OUT_DIR / "封面样例_3_课程考核_空白.docx",
    )
    write_cover(
        build_course_cover_document(course_filled, verbose=True),
        OUT_DIR / "封面样例_4_课程考核_填写.docx",
    )
    write_cover(
        build_thesis_cover_document(thesis_long, verbose=True),
        OUT_DIR / "封面样例_5_论文_填写_长值测试.docx",
    )


if __name__ == "__main__":
    main()
