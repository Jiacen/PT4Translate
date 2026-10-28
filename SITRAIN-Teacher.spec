"""Build on Windows with: .venv\\Scripts\\python.exe -m PyInstaller SITRAIN-Teacher.spec"""
from pathlib import Path

from PyInstaller.utils.hooks import collect_submodules, collect_data_files

root = Path(SPEC).resolve().parent
datas = [
    (str(root / "sitrain_app" / "static"), "sitrain_app/static"),
    (str(root / "config" / "system_config.default.json"), "config"),
    (str(root / "config" / "translation_modes.json"), "config"),
    (str(root / "config" / "prompts"), "config/prompts"),
    (str(root / "config" / "locales"), "config/locales"),
    (str(root / "assets" / "models" / "tiktoken"), "assets/models/tiktoken"),
    (str(root / "LICENSE"), "."),
]
datas += collect_data_files("rapidocr")
hidden = collect_submodules("sitrain_app") + [
    "core.pipelines.ppt_translation_pipeline",
    "core.pipelines.word_translation_pipeline",
    "core.pipelines.excel_translation_pipeline",
]

a = Analysis(
    [str(root / "scripts" / "run_sitrain.py")],
    pathex=[str(root)],
    binaries=[],
    datas=datas,
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["PySide6", "gradio", "torch", "paddle"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name="SITRAIN-Teacher",
          debug=False, bootloader_ignore_signals=False, strip=False, upx=False,
          console=True, disable_windowed_traceback=False)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name="SITRAIN-Teacher")
