"""Build on Windows with: .venv\\Scripts\\python.exe -m PyInstaller SITRAIN-Teacher.spec"""
from pathlib import Path
import re

from PyInstaller.utils.hooks import collect_submodules, collect_data_files
from PyInstaller.utils.win32.versioninfo import (VSVersionInfo, FixedFileInfo, StringFileInfo,
                                                StringTable, StringStruct, VarFileInfo, VarStruct)

root = Path(SPEC).resolve().parent
version_source = (root / "sitrain_app" / "__init__.py").read_text(encoding="utf-8")
version = re.search(r'^__version__ = "([^"]+)"', version_source, re.M).group(1)
build_match = re.search(r'^__build__ = (\d+)', version_source, re.M)
build_number = int(build_match.group(1)) if build_match else 0
release = f"{version}-fix{build_number}" if build_number else version
version_tuple = tuple(int(part) for part in version.split(".")) + (build_number,)
version_info = VSVersionInfo(
    ffi=FixedFileInfo(filevers=version_tuple, prodvers=version_tuple, fileType=1),
    kids=[StringFileInfo([StringTable("040904B0", [
        StringStruct("FileDescription", "SITRAIN 教材翻译平台"),
        StringStruct("FileVersion", '.'.join(map(str, version_tuple))),
        StringStruct("InternalName", "SITRAIN-Teacher"),
        StringStruct("OriginalFilename", "SITRAIN-Teacher.exe"),
        StringStruct("ProductName", "SITRAIN Teacher"),
        StringStruct("ProductVersion", release),
    ])]), VarFileInfo([VarStruct("Translation", [0x0409, 1200])])],
)
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
          console=True, disable_windowed_traceback=False, version=version_info)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name="SITRAIN-Teacher")
