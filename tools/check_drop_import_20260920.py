"""拖入文件导入的检查（隐藏窗口，不弹界面）：向主窗口的真实窗口过程发 WM_DROPFILES，
验证：首次拖入直接导入；已有导入时先弹确认（文案固定），取消则原状不动、确定才覆盖；
不支持的扩展名被忽略；一次拖多个只取第一个可用的；按钮导入同样走确认。

跑法：.venv\\Scripts\\python.exe tools\\check_drop_import_20260920.py    （退出码 0＝全部通过）
"""

import ctypes, sys, tempfile, pathlib, tkinter as tk
from ctypes import wintypes
sys.stdout.reconfigure(encoding="utf-8")
ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
orig_init, orig_state = tk.Tk.__init__, tk.Tk.state
tk.Tk.__init__ = lambda self, *a, **k: (orig_init(self, *a, **k), self.withdraw())[0]
tk.Tk.state = lambda self, *a: "withdrawn"
from tkinter import filedialog, messagebox
from docx import Document
from essay_tool.gui import main_window as mw

bad = 0
def check(n, c, extra=""):
    global bad
    print(("OK   " if c else "FAIL ") + n, "" if c else extra); bad += (not c)

T = pathlib.Path(tempfile.mkdtemp())
def mk_docx(name):
    p = T / name; d = Document(); d.add_paragraph("正文"); d.save(str(p)); return p
A, B = mk_docx("甲.docx"), mk_docx("乙.docx")
TXT = T / "丙.txt"; TXT.write_text("第一段\n第二段", encoding="utf-8")
PNG = T / "图.png"; PNG.write_bytes(b"x")

user32, kernel32 = ctypes.windll.user32, ctypes.windll.kernel32
kernel32.GlobalAlloc.restype = ctypes.c_void_p
kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
kernel32.GlobalLock.restype = ctypes.c_void_p
kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.SendMessageW.restype = ctypes.c_ssize_t
user32.GetParent.argtypes = [wintypes.HWND]; user32.GetParent.restype = wintypes.HWND

def drop(app, *paths):
    """造一个真的 HDROP（DROPFILES 头 + 宽字符路径串）发给主窗口的窗口过程。"""
    body = ("\0".join(str(p) for p in paths) + "\0\0").encode("utf-16-le")
    header = ctypes.create_string_buffer(20)
    ctypes.memmove(header, (ctypes.c_uint32 * 5)(20, 0, 0, 0, 1), 20)
    h = kernel32.GlobalAlloc(0x0042, 20 + len(body))     # GHND
    ptr = kernel32.GlobalLock(h)
    ctypes.memmove(ptr, header.raw, 20); ctypes.memmove(ptr + 20, body, len(body))
    kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]; kernel32.GlobalUnlock(h)
    user32.SendMessageW(app._drop_hwnd, 0x0233, h, 0)

app = mw.MainWindow()
def settle():
    import time                                  # after(10) 要真等到点才会触发
    end = time.time() + 0.4
    while time.time() < end:
        app.update_idletasks(); app.update(); time.sleep(0.02)
asked = []
answer = {"v": True}
def fake_ask(title, message, **k):
    asked.append((title, message)); return answer["v"]
try:
    messagebox.askokcancel = fake_ask
    messagebox.showinfo = messagebox.showwarning = messagebox.showerror = lambda *a, **k: None
    check("窗口已启用拖入（拿到了句柄）", getattr(app, "_drop_hwnd", None) is not None)

    drop(app, A); settle()
    check("首次拖入：直接导入、不弹确认", app.doc_path == A and not asked, f"{app.doc_path} {asked}")

    answer["v"] = False
    drop(app, B); settle()
    check("已有导入：弹确认，文案正确", len(asked) == 1 and asked[0][1] == "继续导入将覆盖已有文件及其编辑，确定继续吗？", str(asked))
    check("选取消：原文件原状不动", app.doc_path == A and app.lbl_doc.cget("text") == "甲.docx")

    answer["v"] = True; asked.clear()
    drop(app, B); settle()
    check("选确定：覆盖成新文件", app.doc_path == B and app.lbl_doc.cget("text") == "乙.docx" and len(asked) == 1)

    asked.clear()
    drop(app, PNG); settle()
    check("不支持的扩展名：忽略，不弹确认、不改导入", app.doc_path == B and not asked)

    drop(app, PNG, TXT, A); settle()
    check("多个文件：取第一个可用的（.txt 转成 docx 副本）",
          app.src_path == TXT and app.doc_path is not None and app.doc_path.suffix == ".docx", f"{app.src_path} {app.doc_path}")

    asked.clear()
    filedialog.askopenfilename = lambda *a, **k: str(A)
    app.on_import_file(); settle()
    check("按钮导入同样先确认", len(asked) == 1 and app.doc_path == A)
    answer["v"] = False; asked.clear()
    app.on_import_file(); settle()
    check("按钮导入选取消：不变", len(asked) == 1 and app.doc_path == A)

    # 取消导入之后再拖：不该再弹确认
    answer["v"] = True; asked.clear()
    mw.messagebox.askokcancel = lambda *a, **k: True
    app.on_cancel_import(); settle()
    mw.messagebox.askokcancel = fake_ask
    drop(app, B); settle()
    check("取消导入后再拖：不弹确认", app.doc_path == B and not asked)
finally:
    app.destroy()
print("全部通过" if not bad else f"{bad} 项失败")
sys.exit(1 if bad else 0)
