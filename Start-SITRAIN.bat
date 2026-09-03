@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Virtual environment not found. Run Setup-SITRAIN.bat first.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" "scripts\run_sitrain.py"
