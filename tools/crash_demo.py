# -*- coding: utf-8 -*-
"""可控崩溃演示：正常启动主窗口，1.5 秒后故意在界面回调里抛一个异常。

走的是**真正的崩溃链路**（跟日常意外崩溃完全同一条）：
界面异常 → 生成错误代码 → 写入 crash.log → 弹出「哎呀！出现了一点小问题」
崩溃弹窗（含「联系开发者」「关闭程序」两个按钮）。

用法：.venv\\Scripts\\python.exe tools\\crash_demo.py
看完就关：点弹窗里的「关闭程序」，或者直接关掉。
"""

import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import essay_tool.gui.main_window as mw  # noqa: E402


def main() -> None:
    app = mw.MainWindow()

    def boom() -> None:
        raise RuntimeError("可控崩溃演示：这是一次故意触发的异常，用来看崩溃界面。")

    app.after(1500, boom)
    print("主窗口已打开；1.5 秒后演示一次崩溃。", flush=True)
    app.mainloop()


if __name__ == "__main__":
    main()
