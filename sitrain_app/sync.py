from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from . import db
from .pipeline import add_version
from .storage import file_dir


def _latest(file_id: str, *types: str) -> dict:
    for kind in types:
        row = db.row("SELECT * FROM file_versions WHERE file_id=? AND version_type=? ORDER BY sequence DESC LIMIT 1",(file_id,kind))
        if row:
            return row
    raise RuntimeError(f"文件 {file_id} 没有可用版本")


def _slide_count(path: Path) -> int:
    with zipfile.ZipFile(path) as archive:
        return sum(1 for n in archive.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml",n))


def _word_slide_media(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as archive:
        names = [n for n in archive.namelist() if re.fullmatch(r"word/media/image\d+\.(?:emf|wmf|png|jpe?g)",n,re.I)]
    return sorted(names,key=lambda n:int(re.search(r"image(\d+)",n,re.I).group(1)))


def create_mappings(chapter_id: str) -> list[dict]:
    files = db.rows("SELECT * FROM files WHERE chapter_id=?",(chapter_id,))
    ppt = next((f for f in files if f["extension"] in {".pptx",".ppt"}),None)
    word = next((f for f in files if f["extension"] in {".docx",".doc"}),None)
    if not ppt or not word:
        raise RuntimeError("该章节必须同时包含 PPT 和 Word 才能建立同步映射")
    ppt_path = Path(_latest(ppt["id"],"published","working","original")["path"])
    word_path = Path(_latest(word["id"],"working","original")["path"])
    media = _word_slide_media(word_path)
    slides = _slide_count(ppt_path)
    if not media:
        raise RuntimeError("Word 中未发现可映射的幻灯片图片")
    offset = max(0,slides-len(media))
    now = db.utcnow()
    with db.connect() as conn:
        for index,part in enumerate(media,1):
            slide_no = min(slides,index+offset)
            confidence = 0.9 if slides-len(media) <= 3 else 0.7
            conn.execute("INSERT INTO ppt_word_mappings(id,chapter_id,ppt_file_id,word_file_id,slide_number,media_part,confidence,confirmed,updated_at) VALUES(?,?,?,?,?,?,?,?,?) ON CONFLICT(word_file_id,media_part) DO UPDATE SET slide_number=excluded.slide_number,confidence=excluded.confidence,updated_at=excluded.updated_at",
                         (db.new_id(),chapter_id,ppt["id"],word["id"],slide_no,part,confidence,1 if confidence>=0.85 else 0,now))
    return db.rows("SELECT * FROM ppt_word_mappings WHERE chapter_id=? ORDER BY slide_number",(chapter_id,))


def apply_sync(chapter_id: str) -> str:
    mappings = db.rows("SELECT * FROM ppt_word_mappings WHERE chapter_id=? ORDER BY slide_number",(chapter_id,))
    if not mappings:
        mappings = create_mappings(chapter_id)
    unconfirmed = [m for m in mappings if not m["confirmed"]]
    if unconfirmed:
        raise RuntimeError(f"仍有 {len(unconfirmed)} 条低置信度映射需要确认")
    ppt_path = Path(_latest(mappings[0]["ppt_file_id"],"published","working","original")["path"])
    word_path = Path(_latest(mappings[0]["word_file_id"],"published","working","original")["path"])
    export_dir = file_dir(mappings[0]["word_file_id"])/"slide_exports"
    export_dir.mkdir(parents=True,exist_ok=True)
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        app = win32com.client.DispatchEx("PowerPoint.Application")
        deck = app.Presentations.Open(str(ppt_path.resolve()),WithWindow=False)
        replacements = {}
        for mapping in mappings:
            ext = Path(mapping["media_part"]).suffix.lstrip(".").upper()
            export_format = "EMF" if ext in {"EMF","WMF"} else ("JPG" if ext in {"JPG","JPEG"} else "PNG")
            target = export_dir/f"slide_{mapping['slide_number']}.{ext.lower()}"
            deck.Slides(mapping["slide_number"]).Export(str(target.resolve()),export_format)
            replacements[mapping["media_part"]] = target.read_bytes()
        deck.Close(); app.Quit(); pythoncom.CoUninitialize()
    except Exception as exc:
        try:
            deck.Close(); app.Quit()
        except Exception:
            pass
        try: pythoncom.CoUninitialize()
        except Exception: pass
        raise RuntimeError(f"PowerPoint 页面渲染失败：{exc}") from exc
    output = file_dir(mappings[0]["word_file_id"])/"published"/f"{word_path.stem}_synced.docx"
    output.parent.mkdir(parents=True,exist_ok=True)
    if output.exists():
        output = output.with_name(f"{output.stem}_{db.new_id()[:8]}.docx")
    with zipfile.ZipFile(word_path) as src, zipfile.ZipFile(output,"x",zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            dst.writestr(info,replacements.get(info.filename,src.read(info.filename)))
    return add_version(mappings[0]["word_file_id"],"published",output,{"ppt_sync":True,"mapping_count":len(mappings)})
