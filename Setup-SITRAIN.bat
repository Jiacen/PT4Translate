@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" python -m venv .venv
".venv\Scripts\python.exe" -m pip install -r requirements\base.txt -r requirements\web.txt pymupdf rapidocr onnxruntime opencv-python-headless "numpy<2.5"
echo Setup complete. Run Start-SITRAIN.bat.
pause
