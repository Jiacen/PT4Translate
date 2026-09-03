from __future__ import annotations

import csv
import json
import mimetypes
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Any

from bs4 import BeautifulSoup, NavigableString
from PIL import Image, ImageDraw, ImageFont

from .config import LEGACY_OFFICE

_ocr_lock = threading.RLock()
_ocr_engine = None
_RASTER_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tif", ".tiff"}
OFFICE_OCR_MIN_BYTES = 25 * 1024
OFFICE_OCR_MIN_PIXELS = 80_000
# Training materials often include screenshots of Siemens applications.  Keep those
# screenshots intact by default: OCR/replacement makes the UI inaccurate and harms
# the visual layout.  This is intentionally a server-side policy, not a user task.
TRANSLATE_OFFICE_IMAGE_TEXT = False


def _get_ocr():
    global _ocr_engine
    with _ocr_lock:
        if _ocr_engine is None:
            try:
                from rapidocr import RapidOCR
            except ImportError as exc:
                raise RuntimeError("检测到需要 OCR 的内容，但未安装 RapidOCR/PP-OCRv6") from exc
            _ocr_engine = RapidOCR()
        return _ocr_engine


def _ocr_image(content: bytes) -> list[dict]:
    result = _get_ocr()(content)
    if result is None or result.boxes is None:
        return []
    return [{"text":text,"confidence":float(score),"box":[[float(x),float(y)] for x,y in box]}
            for box,text,score in zip(result.boxes,result.txts,result.scores) if str(text).strip()]


def _office_ocr_units(path: Path, start: int) -> list[dict]:
    units = []
    with zipfile.ZipFile(path) as archive:
        media = []
        for info in archive.infolist():
            if "/media/" not in info.filename or Path(info.filename).suffix.lower() not in _RASTER_EXTS:
                continue
            if info.file_size >= OFFICE_OCR_MIN_BYTES:
                media.append(info.filename)
                continue
            try:
                with Image.open(BytesIO(archive.read(info.filename))) as image:
                    if image.width * image.height >= OFFICE_OCR_MIN_PIXELS:
                        media.append(info.filename)
            except Exception:
                continue
        media.sort()
        for part in media:
            try:
                regions = _ocr_image(archive.read(part))
            except Exception:
                continue
            for region_no,region in enumerate(regions,1):
                start += 1
                units.append({"stable_key":f"ocr:{start}","page":None,"object_ref":part,"unit_type":"office_image_ocr",
                              "source_text":region["text"],"confidence":region["confidence"],
                              "source_payload":{"media_part":part,"box":region["box"],"region":region_no}})
    return units


def _kind(ext: str) -> str:
    return {".pptx":"powerpoint", ".ppt":"powerpoint", ".docx":"word", ".doc":"word",
            ".xlsx":"excel", ".xls":"excel", ".pdf":"pdf"}.get(ext, "text")


def _office_counts(path: Path) -> tuple[int, int]:
    pages, images = 0, 0
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        images = sum(1 for n in names if re.search(r"/(media)/", n) and not n.endswith("/"))
        if path.suffix.lower() == ".pptx":
            pages = sum(1 for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n))
        elif path.suffix.lower() == ".docx":
            pages = 0
        elif path.suffix.lower() == ".xlsx":
            pages = sum(1 for n in names if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n))
    return pages, images


def preflight(path: Path) -> dict[str, Any]:
    ext = path.suffix.lower()
    result = {"kind":_kind(ext),"extension":ext,"size":path.stat().st_size,"pages":0,
              "text_units":0,"images":0,"needs_ocr":False,"legacy":ext in LEGACY_OFFICE,"risks":[]}
    if ext in {".pptx", ".docx", ".xlsx"}:
        result["pages"], result["images"] = _office_counts(path)
        result["image_text_translation"] = "enabled" if TRANSLATE_OFFICE_IMAGE_TEXT else "disabled"
    elif ext == ".pdf":
        try:
            import fitz
            doc = fitz.open(path)
            result["pages"] = len(doc)
            digital = sum(1 for p in doc if len(p.get_text().strip()) >= 20)
            if digital == 0:
                result["pdf_mode"] = "scanned"
            elif digital == len(doc):
                result["pdf_mode"] = "digital"
            else:
                result["pdf_mode"] = "mixed"
            result["needs_ocr"] = digital < len(doc)
            result["images"] = sum(len(p.get_images(full=True)) for p in doc)
            doc.close()
        except ImportError:
            result["pdf_mode"] = "unknown"
            result["risks"].append("未安装 PyMuPDF，PDF 预检和回填不可用")
    if ext in LEGACY_OFFICE:
        result["risks"].append("旧 Office 格式将转换为现代内部工作副本；回导可能改变动画、字体或版式")
    return result


def convert_legacy(source: Path, work_dir: Path) -> tuple[Path, list[str]]:
    ext = source.suffix.lower()
    if ext not in LEGACY_OFFICE:
        return source, []
    target = work_dir / f"{source.stem}{LEGACY_OFFICE[ext]}"
    work_dir.mkdir(parents=True, exist_ok=True)
    warnings = ["源文件已保留；内部使用现代 Office 工作副本", "默认发布现代格式，旧格式回导不保证无损"]
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        if ext == ".ppt":
            app = win32com.client.DispatchEx("PowerPoint.Application")
            app.Visible = 0
            doc = app.Presentations.Open(str(source.resolve()), WithWindow=False)
            doc.SaveAs(str(target.resolve()), 24)
        elif ext == ".doc":
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            doc = app.Documents.Open(str(source.resolve()), ReadOnly=True)
            doc.SaveAs2(str(target.resolve()), FileFormat=16)
        else:
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            doc = app.Workbooks.Open(str(source.resolve()), ReadOnly=True)
            doc.SaveAs(str(target.resolve()), FileFormat=51)
        doc.Close(False)
        app.Quit()
        pythoncom.CoUninitialize()
    except Exception as exc:
        try:
            doc.Close(False)
            app.Quit()
        except Exception:
            pass
        try: pythoncom.CoUninitialize()
        except Exception: pass
        raise RuntimeError(f"旧 Office 转换失败，原件已保留：{exc}") from exc
    if not target.exists() or target.stat().st_size == 0:
        raise RuntimeError("旧 Office 转换未生成有效工作副本，原件已保留")
    return target, warnings


def _office_extract(path: Path, temp_dir: Path) -> tuple[Path, list[dict]]:
    ext = path.suffix.lower()
    if ext == ".pptx":
        from core.pipelines.ppt_translation_pipeline import extract_ppt_content_to_json
        json_path = Path(extract_ppt_content_to_json(str(path), str(temp_dir)))
    elif ext == ".docx":
        from core.pipelines.word_translation_pipeline import extract_word_content_to_json
        json_path = Path(extract_word_content_to_json(str(path), str(temp_dir)))
    elif ext == ".xlsx":
        from core.pipelines.excel_translation_pipeline import extract_excel_content_to_json
        json_path = Path(extract_excel_content_to_json(str(path), str(temp_dir)))
    else:
        raise ValueError(ext)
    data = json.loads(json_path.read_text(encoding="utf-8"))
    units = []
    for item in data:
        page = item.get("slide_index") or item.get("page")
        if ext == ".xlsx":
            page = item.get("sheet")
        units.append({"stable_key":str(item["count_src"]),"page":page if isinstance(page,int) else None,
                      "object_ref":str(page or item.get("type", "")),"unit_type":item.get("type", "text"),
                      "source_text":str(item.get("value", "")),"confidence":None,"source_payload":item})
    if TRANSLATE_OFFICE_IMAGE_TEXT:
        units.extend(_office_ocr_units(path,len(units)))
    return json_path, units


def _walk_json(value: Any, path: tuple = ()): 
    if isinstance(value, dict):
        for key, child in value.items():
            yield from _walk_json(child, path + (key,))
    elif isinstance(value, list):
        for index, child in enumerate(value):
            yield from _walk_json(child, path + (index,))
    elif isinstance(value, str) and re.search(r"[A-Za-z\u4e00-\u9fff]", value):
        yield path, value


def extract_units(path: Path, temp_dir: Path) -> tuple[Path | None, list[dict]]:
    ext = path.suffix.lower()
    if ext in {".pptx", ".docx", ".xlsx"}:
        return _office_extract(path, temp_dir)
    units: list[dict] = []
    if ext == ".pdf":
        try:
            import fitz
        except ImportError as exc:
            raise RuntimeError("PDF 处理需要安装 PyMuPDF") from exc
        doc = fitz.open(path)
        count = 0
        for page_no, page in enumerate(doc, 1):
            page_units = 0
            for block_no, block in enumerate(page.get_text("blocks")):
                text = block[4].strip()
                if not text:
                    continue
                count += 1
                units.append({"stable_key":str(count),"page":page_no,"object_ref":f"block:{block_no}","unit_type":"pdf_text",
                              "source_text":text,"confidence":1.0,"source_payload":{"rect":list(block[:4]),"block":block_no}})
                page_units += 1
            if page_units == 0:
                pix = page.get_pixmap(matrix=fitz.Matrix(2,2),alpha=False)
                for region_no,region in enumerate(_ocr_image(pix.tobytes("png")),1):
                    count += 1
                    xs=[p[0]/2 for p in region["box"]]; ys=[p[1]/2 for p in region["box"]]
                    units.append({"stable_key":str(count),"page":page_no,"object_ref":f"ocr:{region_no}","unit_type":"pdf_ocr",
                                  "source_text":region["text"],"confidence":region["confidence"],
                                  "source_payload":{"rect":[min(xs),min(ys),max(xs),max(ys)],"box":region["box"],"region":region_no}})
        doc.close()
        return None, units
    if ext == ".json":
        obj = json.loads(path.read_text(encoding="utf-8-sig"))
        for idx, (json_path, text) in enumerate(_walk_json(obj), 1):
            units.append({"stable_key":str(idx),"page":None,"object_ref":json.dumps(json_path,ensure_ascii=False),
                          "unit_type":"json_string","source_text":text,"confidence":None,"source_payload":{"path":json_path}})
    elif ext in {".csv", ".tsv"}:
        delimiter = "\t" if ext == ".tsv" else ","
        with path.open("r", encoding="utf-8-sig", newline="") as stream:
            for row_no, row in enumerate(csv.reader(stream, delimiter=delimiter)):
                for col_no, text in enumerate(row):
                    if re.search(r"[A-Za-z\u4e00-\u9fff]", text):
                        units.append({"stable_key":f"{row_no}:{col_no}","page":None,"object_ref":f"R{row_no+1}C{col_no+1}",
                                      "unit_type":"cell","source_text":text,"confidence":None,"source_payload":{"row":row_no,"col":col_no}})
    elif ext in {".html", ".htm"}:
        soup = BeautifulSoup(path.read_text(encoding="utf-8-sig"), "html.parser")
        for idx, node in enumerate([n for n in soup.find_all(string=True) if isinstance(n, NavigableString) and n.parent.name not in {"script","style"} and n.strip()], 1):
            units.append({"stable_key":str(idx),"page":None,"object_ref":f"text:{idx}","unit_type":"html_text",
                          "source_text":str(node),"confidence":None,"source_payload":{"index":idx-1}})
    else:
        text = path.read_text(encoding="utf-8-sig")
        for idx, line in enumerate(text.splitlines(keepends=False), 1):
            if line.strip() and re.search(r"[A-Za-z\u4e00-\u9fff]", line):
                units.append({"stable_key":str(idx),"page":None,"object_ref":f"line:{idx}","unit_type":"text",
                              "source_text":line,"confidence":None,"source_payload":{"line":idx-1}})
    return None, units


def _set_json_path(obj: Any, path: list, value: str) -> None:
    cur = obj
    for key in path[:-1]:
        cur = cur[key]
    cur[path[-1]] = value


def _font(size: int):
    for candidate in (Path(os.environ.get("WINDIR","C:/Windows"))/"Fonts"/"msyh.ttc",
                      Path(os.environ.get("WINDIR","C:/Windows"))/"Fonts"/"simhei.ttf"):
        if candidate.exists():
            return ImageFont.truetype(str(candidate),size)
    return ImageFont.load_default()


def _redraw_media(content: bytes, regions: list[tuple[dict,str]], suffix: str) -> bytes:
    image = Image.open(BytesIO(content)).convert("RGB")
    draw = ImageDraw.Draw(image)
    for payload,text in regions:
        points=payload["box"]; xs=[p[0] for p in points]; ys=[p[1] for p in points]
        box=(int(min(xs)),int(min(ys)),int(max(xs)),int(max(ys)))
        draw.rectangle(box,fill="white")
        size=max(8,min(32,int((box[3]-box[1])*0.7)))
        font=_font(size)
        while size>7 and draw.textbbox((0,0),text,font=font)[2] > max(1,box[2]-box[0]):
            size-=1; font=_font(size)
        draw.text((box[0],box[1]),text,fill="black",font=font)
    formats={".jpg":"JPEG",".jpeg":"JPEG",".bmp":"BMP",".webp":"WEBP",".tif":"TIFF",".tiff":"TIFF",".png":"PNG"}
    output=BytesIO(); image.save(output,format=formats.get(suffix.lower(),"PNG"),quality=95)
    return output.getvalue()


def _apply_office_ocr(output: Path, units: list[dict]) -> None:
    by_part={}
    for unit in units:
        if unit.get("unit_type") != "office_image_ocr" or unit.get("status") == "skipped_image":
            continue
        translated = unit.get("human_text") or unit.get("machine_text")
        if not translated:
            continue
        by_part.setdefault(unit["source_payload"]["media_part"],[]).append((unit["source_payload"],translated))
    if not by_part:
        return
    temp=output.with_suffix(output.suffix+".ocr.tmp")
    with zipfile.ZipFile(output) as src,zipfile.ZipFile(temp,"x",zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            content=src.read(info.filename)
            if info.filename in by_part:
                content=_redraw_media(content,by_part[info.filename],Path(info.filename).suffix)
            dst.writestr(info,content)
    shutil.move(temp,output)


def rebuild(source: Path, source_json: Path | None, units: list[dict], result_dir: Path,
            source_lang: str, target_lang: str) -> Path:
    result_dir.mkdir(parents=True, exist_ok=True)
    ext = source.suffix.lower()
    translated = [{"count_src":u["stable_key"],"type":u.get("unit_type","text"),"original":u["source_text"],
                   "translated":u.get("human_text") or u.get("machine_text") or u["source_text"]} for u in units]
    if ext in {".pptx", ".docx", ".xlsx"}:
        dst_json = source_json.parent / "dst_translated.json"
        dst_json.write_text(json.dumps(translated,ensure_ascii=False,indent=2),encoding="utf-8")
        if ext == ".pptx":
            from core.pipelines.ppt_translation_pipeline import write_translated_content_to_ppt
            output=Path(write_translated_content_to_ppt(str(source),str(source_json),str(dst_json),str(source_json.parent.parent),str(result_dir),source_lang,target_lang))
            _apply_office_ocr(output,units); return output
        if ext == ".docx":
            from core.pipelines.word_translation_pipeline import write_translated_content_to_word
            output=Path(write_translated_content_to_word(str(source),str(source_json),str(dst_json),str(source_json.parent.parent),str(result_dir),False,source_lang,target_lang))
            _apply_office_ocr(output,units); return output
        from core.pipelines.excel_translation_pipeline import write_translated_content_to_excel
        output=Path(write_translated_content_to_excel(str(source),str(source_json),str(dst_json),str(result_dir),False,False,source_lang,target_lang))
        _apply_office_ocr(output,units); return output
    target_suffix = {"Chinese": "zh", "English": "en"}.get(target_lang, target_lang.casefold().replace(" ", "_"))
    output = result_dir / f"{source.stem}_{target_suffix}{ext}"
    values = {u["stable_key"]:u.get("human_text") or u.get("machine_text") or u["source_text"] for u in units}
    if ext == ".pdf":
        import fitz
        doc = fitz.open(source)
        for unit in units:
            page = doc[int(unit["page"])-1]
            rect = fitz.Rect(unit["source_payload"]["rect"])
            text = values[unit["stable_key"]]
            page.add_redact_annot(rect, fill=(1,1,1))
            page.apply_redactions()
            page.insert_textbox(rect, text, fontsize=max(6,min(11,rect.height/2)),fontname="china-s",color=(0,0,0))
        doc.save(output)
        doc.close()
    elif ext == ".json":
        obj = json.loads(source.read_text(encoding="utf-8-sig"))
        for u in units:
            _set_json_path(obj,u["source_payload"]["path"],values[u["stable_key"]])
        output.write_text(json.dumps(obj,ensure_ascii=False,indent=2),encoding="utf-8")
    elif ext in {".csv", ".tsv"}:
        delimiter = "\t" if ext == ".tsv" else ","
        with source.open("r",encoding="utf-8-sig",newline="") as stream:
            table = list(csv.reader(stream,delimiter=delimiter))
        for u in units:
            table[u["source_payload"]["row"]][u["source_payload"]["col"]] = values[u["stable_key"]]
        with output.open("x",encoding="utf-8-sig",newline="") as stream:
            csv.writer(stream,delimiter=delimiter).writerows(table)
    elif ext in {".html", ".htm"}:
        soup = BeautifulSoup(source.read_text(encoding="utf-8-sig"),"html.parser")
        nodes = [n for n in soup.find_all(string=True) if isinstance(n,NavigableString) and n.parent.name not in {"script","style"} and n.strip()]
        for u in units:
            nodes[u["source_payload"]["index"]].replace_with(values[u["stable_key"]])
        output.write_text(str(soup),encoding="utf-8")
    else:
        lines = source.read_text(encoding="utf-8-sig").splitlines(keepends=True)
        for u in units:
            idx = u["source_payload"]["line"]
            ending = "\n" if lines[idx].endswith("\n") else ""
            lines[idx] = values[u["stable_key"]] + ending
        output.write_text("".join(lines),encoding="utf-8")
    return output


def _render_pptx_with_powerpoint(source: Path, page_number: int, target: Path) -> bool:
    """Use a short-lived helper: Office COM must not live in the web worker."""
    helper = Path(__file__).resolve().parent.parent / "scripts" / "render_ppt_preview.py"
    if not helper.exists() or os.name != "nt":
        return False
    try:
        run = subprocess.run([sys.executable, str(helper), str(source), str(target), str(page_number)],
                             capture_output=True, text=True, timeout=120, check=False)
    except (OSError, subprocess.TimeoutExpired):
        return False
    return run.returncode == 0 and target.exists() and target.stat().st_size > 0


def _render_pptx_deck_with_powerpoint(source: Path, cache_dir: Path) -> bool:
    """Render every slide in one isolated Office process, not one process/page."""
    helper = Path(__file__).resolve().parent.parent / "scripts" / "render_ppt_preview.py"
    if not helper.exists() or os.name != "nt":
        return False
    try:
        run = subprocess.run([sys.executable, str(helper), str(source), str(cache_dir), "all"],
                             capture_output=True, text=True, timeout=300, check=False)
    except (OSError, subprocess.TimeoutExpired):
        return False
    return run.returncode == 0 and (cache_dir / "page-1-rendered.png").exists()


def _render_pptx_with_libreoffice(source: Path, page_number: int, cache_dir: Path) -> Path | None:
    """Cloud-safe renderer used on Ubuntu VPS deployments when LibreOffice exists."""
    soffice = shutil.which("soffice")
    if not soffice:
        return None
    pdf = cache_dir / f"{source.stem}-rendered.pdf"
    if not pdf.exists():
        try:
            run = subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(cache_dir), str(source)],
                                 capture_output=True, text=True, timeout=120, check=False)
        except (OSError, subprocess.TimeoutExpired):
            return None
        generated = cache_dir / f"{source.stem}.pdf"
        if run.returncode != 0 or not generated.exists():
            return None
        generated.replace(pdf)
    return render_preview(pdf, page_number, cache_dir / "pdf-render")


def prepare_preview_cache(source: Path, cache_dir: Path, pages: int) -> None:
    """Pre-render the deliverable once so review-page navigation is instant.

    On Windows the helper opens PowerPoint only once for the entire deck. On a
    VPS, LibreOffice converts the deck to a PDF once and PyMuPDF renders each
    page locally. Failures are intentionally non-fatal: on-demand preview still
    remains available as a fallback.
    """
    cache_dir.mkdir(parents=True, exist_ok=True)
    if pages < 1 or source.suffix.lower() != ".pptx":
        return
    if all((cache_dir / f"page-{page}-rendered.png").exists() for page in range(1, pages + 1)):
        return
    if _render_pptx_deck_with_powerpoint(source, cache_dir):
        return
    soffice = shutil.which("soffice")
    if not soffice:
        return
    try:
        run = subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(cache_dir), str(source)],
                             capture_output=True, text=True, timeout=300, check=False)
    except (OSError, subprocess.TimeoutExpired):
        return
    pdf = cache_dir / f"{source.stem}.pdf"
    if run.returncode != 0 or not pdf.exists():
        return
    try:
        import fitz
        doc = fitz.open(pdf)
        for page in range(1, min(pages, len(doc)) + 1):
            doc[page - 1].get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False).save(cache_dir / f"page-{page}-rendered.png")
        doc.close()
    except Exception:
        return


def render_preview(source: Path, page_number: int, cache_dir: Path) -> Path:
    """Render one page without executing document macros or embedded links."""
    cache_dir.mkdir(parents=True,exist_ok=True)
    target=cache_dir/f"page-{page_number}-rendered.png"
    if target.exists(): return target
    ext=source.suffix.lower()
    if ext == ".pdf":
        import fitz
        doc=fitz.open(source)
        if page_number < 1 or page_number > len(doc):
            doc.close(); raise RuntimeError("预览页码超出范围")
        pix=doc[page_number-1].get_pixmap(matrix=fitz.Matrix(1.5,1.5),alpha=False)
        pix.save(target); doc.close(); return target
    if ext == ".pptx":
        # The native/LibreOffice paths render the actual deck: backgrounds, theme,
        # shapes, font metrics and positions match the downloadable PPTX.
        if _render_pptx_with_powerpoint(source, page_number, target):
            return target
        rendered = _render_pptx_with_libreoffice(source, page_number, cache_dir)
        if rendered:
            return rendered
        # Last-resort preview for developer environments without a renderer.
        # It is intentionally not used on a properly provisioned VPS.
        try:
            from pptx import Presentation
            from pptx.enum.shapes import MSO_SHAPE_TYPE
            deck=Presentation(source)
            if page_number < 1 or page_number > len(deck.slides): raise RuntimeError("预览页码超出范围")
            width,height=1280,720; sx=width/deck.slide_width; sy=height/deck.slide_height
            canvas=Image.new("RGB",(width,height),"white"); draw=ImageDraw.Draw(canvas)
            for shape in deck.slides[page_number-1].shapes:
                x,y,w,h=[int(v) for v in (shape.left*sx,shape.top*sy,shape.width*sx,shape.height*sy)]
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    picture=Image.open(BytesIO(shape.image.blob)).convert("RGB").resize((max(1,w),max(1,h)))
                    canvas.paste(picture,(x,y)); continue
                text=""
                if getattr(shape,"has_text_frame",False): text=shape.text
                elif getattr(shape,"has_table",False): text="\n".join(" | ".join(c.text for c in row.cells) for row in shape.table.rows)
                if text:
                    draw.multiline_text((x,y),text,fill="black",font=_font(max(10,min(24,int(h/max(1,text.count(chr(10))+1)*0.45)))),spacing=4)
            canvas.save(target,"PNG"); return target
        except Exception as exc:
            raise RuntimeError(f"PowerPoint 预览渲染失败：{exc}") from exc
    if ext in {".docx", ".xlsx"}:
        pdf=cache_dir/f"{source.stem}.pdf"
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            if ext == ".docx":
                app=win32com.client.DispatchEx("Word.Application"); app.Visible=False
                doc=app.Documents.Open(str(source.resolve()),ReadOnly=True); doc.ExportAsFixedFormat(str(pdf.resolve()),17)
            else:
                app=win32com.client.DispatchEx("Excel.Application"); app.Visible=False; app.DisplayAlerts=False
                doc=app.Workbooks.Open(str(source.resolve()),ReadOnly=True); doc.ExportAsFixedFormat(0,str(pdf.resolve()))
            doc.Close(False); app.Quit(); pythoncom.CoUninitialize(); return render_preview(pdf,page_number,cache_dir)
        except Exception as exc:
            try: doc.Close(False); app.Quit()
            except Exception: pass
            try: pythoncom.CoUninitialize()
            except Exception: pass
            raise RuntimeError(f"Office 预览渲染失败：{exc}") from exc
    raise RuntimeError("该格式没有页面预览；请使用右侧结构化文本校对")
