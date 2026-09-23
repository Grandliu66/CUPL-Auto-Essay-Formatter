# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 配置（onedir：一个文件夹，含 exe + _internal）。

用法：不要直接手敲，跑 `packaging\\build.ps1`——它先建专用环境、更新官方文件校验清单、
再调用本文件、最后自检出包结果。

随程序分发的只读资源（进 `_internal`，目录结构与仓库一致，运行时由
`essay_tool.config.paths.resource_dir()` 定位）：
  reference/   —— 只带运行时真会读的 5 份官方文件（规则同 tools/update_integrity.py），
                  其余 reference 里的测试稿、流程图、规范草稿一概不带
  presets/     —— 4 份内置预设 .json（.xlsx 只是开发工具的生成物，运行时不读）
  essay_tool/gui/assets/ —— 封面样张 PNG 与程序图标
"""

import pathlib
import sys

ROOT = pathlib.Path(SPECPATH).resolve().parent          # noqa: F821  (SPECPATH 由 PyInstaller 注入)
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tools"))

import update_integrity                                   # noqa: E402  只用它的文件名规则
from essay_tool import __version__                        # noqa: E402

APP_NAME = "CUPL-Undergraduate-Thesis-Auto-Formatter"
ICON = str(ROOT / "essay_tool" / "gui" / "assets" / "app.ico")

# ---- 随包资源
datas = []
for _label, pattern in update_integrity.PATTERNS:
    found = sorted(p for p in (ROOT / "reference").glob(pattern) if not p.name.startswith("~$"))
    if not found:
        raise SystemExit(f"打包失败：reference/ 里找不到「{_label}」（{pattern}）")
    datas.append((str(found[0]), "reference"))            # 与运行时一致：取排序第一个
for p in sorted((ROOT / "presets").glob("*.json")):
    datas.append((str(p), "presets"))
for p in sorted((ROOT / "essay_tool" / "gui" / "assets").iterdir()):
    datas.append((str(p), "essay_tool/gui/assets"))

# ---- Windows 版本信息（资源管理器"详细信息"页里看到的；也让杀毒软件少一点误报）
_v = tuple((__version__.split(".") + ["0", "0", "0"])[:3]) + ("0",)
_vt = ", ".join(str(int(x)) for x in _v)
version_file = ROOT / "build" / "version_info.txt"
version_file.parent.mkdir(exist_ok=True)
version_file.write_text(f"""VSVersionInfo(
  ffi=FixedFileInfo(filevers=({_vt}), prodvers=({_vt}), mask=0x3f, flags=0x0, OS=0x40004,
                    fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[
    StringFileInfo([StringTable('080404b0', [
      StringStruct('FileDescription', 'CUPL Undergraduate Thesis Auto-Formatter'),
      StringStruct('FileVersion', '{__version__}'),
      StringStruct('InternalName', '{APP_NAME}'),
      StringStruct('OriginalFilename', '{APP_NAME}.exe'),
      StringStruct('ProductName', 'CUPL Undergraduate Thesis Auto-Formatter'),
      StringStruct('ProductVersion', '{__version__}')])]),
    VarFileInfo([VarStruct('Translation', [2052, 1200])])
  ])
""", encoding="utf-8")

a = Analysis(                                             # noqa: F821
    [str(ROOT / "main.py")],
    pathex=[str(ROOT)],
    datas=datas,
    hiddenimports=[],
    # 开发/测试用的东西一律不进包（专用打包环境里本来也没装，这里是双保险）
    excludes=["fitz", "pymupdf", "win32com", "win32api", "win32con", "pythoncom",
              "pywintypes", "numpy", "matplotlib", "pytest", "IPython"],
    noarchive=False,
)
pyz = PYZ(a.pure)                                         # noqa: F821
exe = EXE(                                                # noqa: F821
    pyz, a.scripts, [],
    exclude_binaries=True,
    name=APP_NAME,
    console=False,                                        # 不弹黑色控制台窗口
    icon=ICON,
    version=str(version_file),
    upx=False,                                            # 不压壳：UPX 压过的 exe 更容易被杀软误报
)
coll = COLLECT(                                           # noqa: F821
    exe, a.binaries, a.datas,
    strip=False, upx=False,
    name=APP_NAME,
)
