@echo off
rem ============================================================
rem  论文排版工具 —— 双击这个文件即可启动（不会留黑色控制台窗口）
rem
rem  想看到报错信息时，改用下面这行（会开一个控制台窗口）：
rem    ".venv\Scripts\python.exe" "essay_tool\gui\main_window.py"
rem ============================================================
cd /d "%~dp0"
start "" ".venv\Scripts\pythonw.exe" "essay_tool\gui\main_window.py"
