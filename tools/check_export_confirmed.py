"""导出只认结构页"确认"过的启用状态（含清单域）——不碰窗口、不动用户设置文件的回归检查。

背景（2026-10-11 审计）：`MainWindow._build_ops_for_export()` 里"部件启用"取确认时的快照，
但"要不要插清单域"（`list_plan`）由实时的 `structure_required` 算出，同一次导出里两处来源不一致：
  * 确认时开、事后取消勾选（没确认）⇒ 清单域消失；
  * 确认时关、事后勾选（没确认）⇒ 清单域凭空多出来。
界面弹窗承诺"这次改动不会应用"，必须以确认时为准。

用法：.venv/Scripts/python.exe tools/check_export_confirmed.py
"""

from __future__ import annotations

import pathlib
import sys
import tempfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import tkinter as tk                                                  # noqa: E402
from docx import Document                                             # noqa: E402

fails: list[str] = []


def check(name: str, ok: bool, detail: str = "") -> None:
    print(f"   {'OK  ' if ok else 'FAIL'}  {name}" + (f"  → {detail}" if detail and not ok else ""))
    if not ok:
        fails.append(name)


def main() -> int:
    orig_init, orig_state = tk.Tk.__init__, tk.Tk.state

    def quiet_init(self, *a, **k):
        orig_init(self, *a, **k)
        self.withdraw()

    tk.Tk.__init__, tk.Tk.state = quiet_init, (lambda self, *a: "withdrawn")
    try:
        from essay_tool.gui import adjust_preview as ap, main_window as mw

        mw.save_settings = lambda s: None                              # 不动用户真实的设置文件
        with tempfile.TemporaryDirectory() as tmp:
            src = pathlib.Path(tmp) / "doc.docx"
            doc = Document()
            for text in ["论文题目", "【插图清单】", "一、引言", "正文第一段。", "图 1 示意图", "正文第二段。"]:
                doc.add_paragraph(text)
            doc.save(str(src))

            def list_fields(ops):
                return [op.part for op in ops if op.action == "generate" and op.kind == "list_field"]

            app = mw.MainWindow()
            try:
                app.doc_path = src
                app._rebuild_struct_adjust()
                inst = app._adjust_instance
                saved = dict(ap.ENABLED)

                print("== 结构页勾选 vs 导出 ==")
                ap.ENABLED["插图清单"] = True
                inst.on_confirm()
                ap.ENABLED["插图清单"] = False                          # 取消勾选，但没点确认
                _index, ops = app._build_ops_for_export()
                check("确认时开、事后取消（没确认）⇒ 导出仍有插图清单域", list_fields(ops) == ["插图清单"],
                      str(list_fields(ops)))
                cfg = app._collect_current_preset_config(for_export=True)
                check("导出用的配置：启用状态取确认时的快照", cfg.structure_required["插图清单"] is True)
                check("保存预设用的配置：仍是实时值", app._collect_current_preset_config()
                      .structure_required["插图清单"] is False)

                ap.ENABLED["插图清单"] = False
                inst.on_confirm()
                ap.ENABLED["插图清单"] = True                           # 勾选了，但没点确认
                _index, ops = app._build_ops_for_export()
                check("确认时关、事后勾选（没确认）⇒ 导出里没有清单域", list_fields(ops) == [], str(list_fields(ops)))

                ap.ENABLED["插图清单"] = True
                inst.on_confirm()                                       # 确认后再导出，当然要有
                _index, ops = app._build_ops_for_export()
                check("确认之后 ⇒ 导出有插图清单域", list_fields(ops) == ["插图清单"], str(list_fields(ops)))
                ap.ENABLED.clear()
                ap.ENABLED.update(saved)
            finally:
                app.destroy()
    finally:
        tk.Tk.__init__, tk.Tk.state = orig_init, orig_state
    print()
    print(f"失败 {len(fails)} 条")
    return 1 if fails else 0


if __name__ == "__main__":
    raise SystemExit(main())
