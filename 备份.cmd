@echo off
rem 双击本文件即可给项目打一份快照备份（不需要 git、不需要联网）。
chcp 65001 >nul
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo 找不到 .venv\Scripts\python.exe，请先确认项目环境完整。
  pause
  exit /b 1
)
".venv\Scripts\python.exe" "tools\backup_snapshot.py"
echo.
echo 备份完成后可以关掉本窗口。
pause
