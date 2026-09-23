"""程序入口（PyInstaller 打包的入口脚本；开发时也可以 `python main.py` 直接启动）。

`--selftest <结果文件>`：不弹出任何窗口，检查"打包后该有的东西是不是都在、能不能用"
（内置预设、官方文件校验、封面样张与图标、python-docx/openpyxl/lxml/Pillow 能否工作、
空白模板能否生成、主窗口能否构造），把逐项结果写进结果文件，全部通过退出码为 0。
打包脚本在出包后用它自动验收，安装到别的机器上出问题时也可以用它先自查。
"""

from __future__ import annotations

import sys


def _selftest(out_path: str) -> int:
    import pathlib
    import tempfile
    import traceback

    lines: list[str] = []
    bad = 0

    def check(name: str, fn) -> None:
        nonlocal bad
        try:
            detail = fn()
            lines.append(f"OK    {name}" + (f"：{detail}" if detail else ""))
        except Exception:                                    # noqa: BLE001
            bad += 1
            lines.append(f"FAIL  {name}\n" + traceback.format_exc())

    from essay_tool import __version__
    from essay_tool.config import paths
    lines.append(f"版本 {__version__}；frozen={paths.is_frozen()}；资源根 {paths.resource_dir()}")
    tmp = pathlib.Path(tempfile.mkdtemp(prefix="aet-selftest-"))

    def presets():
        from essay_tool.config.preset_io import load_preset
        files = sorted(paths.builtin_presets_dir().glob("*.json"))
        assert len(files) >= 4, f"内置预设只有 {len(files)} 份"
        for f in files:
            load_preset(f)
        return f"{len(files)} 份可读"

    def official():
        from essay_tool.config.integrity import verify_official
        from essay_tool.engine import cover
        from essay_tool.gui import spec_viewer
        files = [cover.THESIS_TEMPLATE, cover.COURSE_TEMPLATE, cover._tenstar_template(),
                 spec_viewer.find_default_spec(), spec_viewer.find_star10_spec()]
        for f in files:
            assert f is not None, "缺官方文件"
            verify_official(f)
        return f"{len(files)} 份校验通过"

    def assets():
        from PIL import Image
        d = paths.resource_dir() / "essay_tool" / "gui" / "assets"
        for n in ("cover_thesis.png", "cover_year.png", "cover_course.png",
                  "cover_tenstar.png", "app.ico"):
            with Image.open(d / n) as im:
                im.load()
        return "封面样张 4 张 + 图标"

    def covers():
        from essay_tool.engine import cover
        f = cover.CoverFields(title="测试题目", student_name="张三", student_id="1",
                              college="法学院", major="法学", grade="2022",
                              advisor_name="李四", course_name="课", course_code="1",
                              course_seq="1", teacher_name="王", session="二十四")
        for fn in (cover.build_thesis_cover_document, cover.build_course_cover_document,
                   cover.build_tenstar_cover_document):
            fn(f)
        return "三种封面可生成"

    def blank_template():
        from essay_tool.config.preset_io import load_preset
        from essay_tool.engine.template import build_template
        preset = load_preset(next(iter(sorted(paths.builtin_presets_dir().glob("*.json")))))
        out = build_template(None, tmp / "空白模板.docx", work_dir=tmp / "w", with_cover=True,
                             preset_config=preset)
        from docx import Document
        assert len(Document(str(out)).paragraphs) > 5
        return "空白模板（含封面）可生成并可读回"

    def excel_roundtrip():
        from essay_tool.config.preset_io import export_to_excel, load_preset
        cfg = load_preset(next(iter(sorted(paths.builtin_presets_dir().glob("*.json")))))
        export_to_excel(cfg, tmp / "p.xlsx")
        load_preset(tmp / "p.xlsx")
        return "预设 ↔ Excel 往返正常（openpyxl）"

    def user_dirs():
        for d in (paths.user_data_dir(), paths.user_presets_dir(), paths.import_backup_dir(),
                  paths.session_work_dir()):
            probe = d / ".selftest"
            probe.write_text("x", encoding="utf-8")
            probe.unlink()
        paths.cleanup_session_work_dir()
        return "用户数据目录与临时目录可写"

    def gui():
        # 构造主窗口但不让它显示：拦掉 state()/deiconify，构造完立刻销毁
        import tkinter as tk
        orig_init, orig_state = tk.Tk.__init__, tk.Tk.state

        def quiet_init(self, *a, **k):
            orig_init(self, *a, **k)
            self.withdraw()
        tk.Tk.__init__, tk.Tk.state = quiet_init, (lambda self, *a: "withdrawn")
        try:
            from essay_tool.gui import main_window as mw
            app = mw.MainWindow()
            try:
                from essay_tool.config.preset_io import load_preset
                app._load_formats_from(
                    load_preset(next(iter(sorted(paths.builtin_presets_dir().glob("*.json"))))), "selftest")
                app.update_idletasks()
                assert len(app.fmt_rows) >= 20, "格式表没有灌进数据"
                assert app._collect_current_preset_config().styles, "收不出预设"
            finally:
                app.destroy()
        finally:
            tk.Tk.__init__, tk.Tk.state = orig_init, orig_state
        return "主窗口可构造，格式表可灌数据，预设可收集"

    for name, fn in (("内置预设", presets), ("官方文件校验", official), ("图片与图标", assets),
                     ("封面生成", covers), ("空白模板", blank_template),
                     ("Excel 往返", excel_roundtrip), ("数据目录", user_dirs), ("主窗口", gui)):
        check(name, fn)
    lines.append("结果：" + ("全部通过" if not bad else f"{bad} 项失败"))
    pathlib.Path(out_path).write_text("\n".join(lines), encoding="utf-8")
    return 1 if bad else 0


if __name__ == "__main__":
    if len(sys.argv) >= 2 and sys.argv[1] == "--selftest":
        sys.exit(_selftest(sys.argv[2] if len(sys.argv) >= 3 else "selftest.txt"))
    from essay_tool.gui.main_window import main
    main()
