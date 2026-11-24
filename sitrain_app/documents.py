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
PPT_RENDER_CACHE_VERSION = "ppt-native-v2"


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


def document_page_count(source: Path, cache_dir: Path) -> int:
    """Return a Word document's real printed-page count.

    A DOCX package has no dependable page-count field.  Word is the layout
    engine that defines pagination, so export a read-only PDF once and count
    its pages.  The PDF is retained in the preview cache for later rendering.
    """
    if source.suffix.lower() != ".docx":
        return 0
    cache_dir.mkdir(parents=True, exist_ok=True)
    pdf = cache_dir / f"{source.stem}.pdf"
    app = doc = None
    initialized = False
    try:
        import fitz
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        initialized = True
        if not pdf.exists() or pdf.stat().st_size == 0:
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            try:
                app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            except Exception:
                pass
            doc = app.Documents.Open(str(source.resolve()), ReadOnly=True,
                                     AddToRecentFiles=False, ConfirmConversions=False)
            doc.ExportAsFixedFormat(str(pdf.resolve()), 17)
        if not pdf.exists() or pdf.stat().st_size == 0:
            return 0
        rendered = fitz.open(pdf)
        try:
            return len(rendered)
        finally:
            rendered.close()
    except Exception:
        return 0
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        if initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


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
        source_text = str(item.get("value", ""))
        if ext == ".pptx":
            source_text = source_text.replace("␊", "\n").replace("␍", "\n")
        units.append({"stable_key":str(item["count_src"]),"page":page if isinstance(page,int) else None,
                      "object_ref":str(page or item.get("type", "")),"unit_type":item.get("type", "text"),
                      "source_text":source_text,"confidence":None,"source_payload":item})
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
            from .ppt_fidelity import verify_pptx_preservation
            output=Path(write_translated_content_to_ppt(str(source),str(source_json),str(dst_json),str(source_json.parent.parent),str(result_dir),source_lang,target_lang,preserve_formatting=True))
            verify_pptx_preservation(source, output)
            _apply_office_ocr(output, units)
            return output
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


def _ppt_render_command(source: Path, target: Path, page: str) -> list[str] | None:
    if os.name != "nt":
        return None
    if getattr(sys, "frozen", False):
        return [sys.executable, "--render-ppt-preview", str(source), str(target), page]
    helper = Path(__file__).resolve().parent.parent / "scripts" / "render_ppt_preview.py"
    if not helper.exists():
        return None
    return [sys.executable, str(helper), str(source), str(target), page]


def _render_pptx_with_powerpoint(source: Path, page_number: int, target: Path) -> bool:
    """Use a short-lived helper: Office COM must not live in the web worker."""
    command = _ppt_render_command(source, target, str(page_number))
    if command is None:
        return False
    try:
        run = subprocess.run(command, capture_output=True, text=True, timeout=120, check=False,
                             creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except (OSError, subprocess.TimeoutExpired):
        return False
    return run.returncode == 0 and target.exists() and target.stat().st_size > 0


def _render_pptx_deck_with_powerpoint(source: Path, cache_dir: Path) -> bool:
    """Render every slide in one isolated Office process, not one process/page."""
    command = _ppt_render_command(source, cache_dir, "all")
    if command is None:
        return False
    try:
        run = subprocess.run(command, capture_output=True, text=True, timeout=300, check=False,
                             creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
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
    if pages < 1:
        return
    if source.suffix.lower() == ".docx":
        # Word startup and PDF export are expensive.  Do both once when the
        # machine version is produced, then serve every review page as a local
        # PNG just like a PPT deck.
        actual_pages = document_page_count(source, cache_dir)
        pdf = cache_dir / f"{source.stem}.pdf"
        if actual_pages < 1 or not pdf.exists():
            return
        try:
            import fitz
            rendered = fitz.open(pdf)
            for page in range(1, min(actual_pages, len(rendered)) + 1):
                target = cache_dir / f"page-{page}-rendered.png"
                if not target.exists():
                    rendered[page - 1].get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False).save(target)
            rendered.close()
        except Exception:
            return
        return
    if source.suffix.lower() != ".pptx":
        return
    cache_dir = cache_dir / PPT_RENDER_CACHE_VERSION
    cache_dir.mkdir(parents=True, exist_ok=True)
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
    ext=source.suffix.lower()
    if ext == ".pptx":
        cache_dir = cache_dir / PPT_RENDER_CACHE_VERSION
    cache_dir.mkdir(parents=True,exist_ok=True)
    target=cache_dir/f"page-{page_number}-rendered.png"
    if target.exists(): return target
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
        raise RuntimeError("无法生成保留原版式的 PPT 预览。请确认已安装 Microsoft PowerPoint，"
                           "并重试。原始文件和翻译文件均已保留。")
    if ext in {".docx", ".xlsx"}:
        pdf=cache_dir/f"{source.stem}.pdf"
        try:
            # A completed Word/PDF export is immutable for this version.  Do
            # not start Office again when users simply turn to another page.
            if pdf.exists() and pdf.stat().st_size > 0:
                return render_preview(pdf,page_number,cache_dir)
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


def word_page_objects(source: Path, page_number: int, cache_dir: Path, units: list[dict]) -> dict:
    """Locate editable Word text on a rendered page without treating images as text.

    Word supplies the authoritative pagination through its cached PDF.  We map a
    translation unit only when its current target text occurs in exactly one PDF
    text block; ambiguous or image-only content gets no click target.
    """
    if source.suffix.lower() != ".docx":
        raise ValueError("仅支持 Word 文档")
    pdf = cache_dir / f"{source.stem}.pdf"
    if not pdf.exists():
        document_page_count(source, cache_dir)
    if not pdf.exists():
        raise RuntimeError("Word 页面定位缓存尚未准备好")
    import fitz
    doc = fitz.open(pdf)
    try:
        if page_number < 1 or page_number > len(doc):
            raise ValueError("页码超出范围")
        page = doc[page_number - 1]
        blocks = page.get_text("blocks")
        def compact(value: str) -> str:
            return re.sub(r"[\s\u240a\u240d]+", "", value or "")
        page_text = ""
        indexed_blocks = []
        for block_no, block in enumerate(blocks):
            text = compact(block[4])
            start = len(page_text)
            page_text += text
            indexed_blocks.append((block_no, block, start, len(page_text)))
        objects = []
        for unit in units:
            target = compact(unit.get("human_text") or unit.get("machine_text") or "")
            # Two-character Chinese labels such as “安全” are legitimate Word
            # text and must remain editable.  Only ignore empty/single-glyph
            # layout noise.
            if len(target) < 2:
                continue
            start = page_text.find(target)
            if start < 0:
                continue
            end = start + len(target)
            matching = [item for item in indexed_blocks
                        if item[3] > start and item[2] < end and item[3] > item[2]]
            if not matching:
                continue
            x0 = min(item[1][0] for item in matching)
            y0 = min(item[1][1] for item in matching)
            x1 = max(item[1][2] for item in matching)
            y1 = max(item[1][3] for item in matching)
            objects.append({"id": f"word-unit:{unit['id']}", "unit_id": unit["id"],
                            "name": (unit.get("source_text") or "Word 文本")[:80],
                            "box": [x0 / page.rect.width, y0 / page.rect.height,
                                    (x1 - x0) / page.rect.width, (y1 - y0) / page.rect.height]})
        return {"width": page.rect.width, "height": page.rect.height, "objects": objects}
    finally:
        doc.close()
