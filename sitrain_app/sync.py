from __future__ import annotations

import io
import json
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from PIL import Image, ImageChops, ImageOps, ImageStat

from . import db
from .documents import document_page_count, render_preview
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
        # SITRAIN source material uses full-slide EMF/WMF media with a VML
        # crop.  Small PNG/JPEG illustrations are not slide substitutes.
        names = [n for n in archive.namelist() if re.fullmatch(r"word/media/image\d+\.(?:emf|wmf)",n,re.I)]
    return sorted(names,key=lambda n:int(re.search(r"image(\d+)",n,re.I).group(1)))


def _is_cover_slide(path: Path, page: int) -> bool:
    """Recognise a title/cover page so sequence matching can skip it."""
    try:
        with zipfile.ZipFile(path) as archive:
            xml = archive.read(f"ppt/slides/slide{page}.xml").decode("utf-8", errors="ignore")
    except (KeyError, zipfile.BadZipFile):
        return False
    text = " ".join(re.findall(r"<a:t>(.*?)</a:t>", xml, re.S)).lower()
    # A course cover normally identifies its chapter but contains none of the
    # body-page cues.  This intentionally remains conservative: it is used
    # only alongside a reliable following Word/PPT match.
    return bool(re.search(r"\b(chapter|module|lesson)\b", text)) and not bool(
        re.search(r"\b(objective|content|overview|you\.{3})\b", text)
    )


def _fingerprint(image: Image.Image) -> list[float]:
    """A size-independent visual fingerprint that tolerates text reflow."""
    grayscale = ImageOps.grayscale(image).resize((48, 27), Image.Resampling.LANCZOS)
    pixels = list(grayscale.getdata())
    mean = sum(pixels) / len(pixels)
    return [pixel - mean for pixel in pixels]


def _distance(left: list[float], right: list[float]) -> float:
    return sum((a - b) ** 2 for a, b in zip(left, right)) / len(left)


def _page_fingerprints(image: Image.Image) -> list[list[float]]:
    """Return whole-page and content-area visual evidence.

    PPT renderers can turn an otherwise transparent logo or decorative group
    into an opaque rectangle.  A whole-slide comparison then incorrectly
    considers two identical teaching pages unrelated.  The upper and left
    regions preserve the actual title/body/screenshots while excluding the
    usual lower-right decoration.  We still retain the full-page evidence for
    slides without this renderer artefact.
    """
    width, height = image.size
    regions = [
        image,
        image.crop((0, 0, int(width * 0.70), int(height * 0.72))),
        image.crop((0, 0, width, int(height * 0.58))),
        image.crop((0, 0, int(width * 0.65), height)),
    ]
    return [_fingerprint(region) for region in regions]


def _visual_distance(left: list[list[float]], right: list[list[float]]) -> float:
    # A strong agreement in one information-rich content region is better
    # evidence than a damaged decorative corner.  The matching region is kept
    # positionally aligned, so blank-area coincidences cannot match unrelated
    # layouts across different parts of a page.
    return min(_distance(a, b) for a, b in zip(left, right))


def _visual_mappings(word_path: Path, ppt_path: Path, cache_dir: Path) -> list[tuple[str, int, float, bool, list[dict]]]:
    """Match full-slide Word media to the original PPT by visual evidence."""
    parts = _word_slide_media(word_path)
    slides = _slide_count(ppt_path)
    cache_dir.mkdir(parents=True, exist_ok=True)
    slide_fingerprints = [_page_fingerprints(Image.open(render_preview(ppt_path, page, cache_dir)).convert("RGB"))
                          for page in range(1, slides + 1)]
    candidates = []
    with zipfile.ZipFile(word_path) as archive:
        for part in parts:
            try:
                fingerprint = _page_fingerprints(Image.open(io.BytesIO(archive.read(part))).convert("RGB"))
            except Exception:
                candidates.append((part, 1, 0.0, False, []))
                continue
            ranked = sorted((_visual_distance(fingerprint, slide), page)
                            for page, slide in enumerate(slide_fingerprints, 1))
            best_score, best_page = ranked[0]
            next_score = ranked[1][0] if len(ranked) > 1 else float("inf")
            # A clear margin proves the visual match.  The absolute ceiling
            # rejects unrelated diagrams even if they happen to be the least
            # different image on the page.
            margin = next_score - best_score
            reliable = best_score <= 450 and margin >= max(30, best_score * 0.5)
            confidence = min(0.99, max(0.05, margin / max(1.0, next_score)))
            choices = [{"slide_number": page, "score": round(score, 2)} for score, page in ranked[:3]]
            candidates.append((part, best_page, confidence, reliable, choices))
    # Slides normally appear once in a Word chapter.  Do not silently assign
    # two images to one slide; retain the stronger visual match only.
    accepted_pages: set[int] = set()
    result = []
    for part, page, confidence, reliable, choices in sorted(candidates, key=lambda item: item[2], reverse=True):
        accepted = reliable and page not in accepted_pages
        if accepted:
            accepted_pages.add(page)
        result.append((part, page, confidence, accepted, choices))
    result.sort(key=lambda item: int(re.search(r"image(\d+)", item[0], re.I).group(1)))

    # Office sometimes renders a transparent SVG/EMF group in a source PPT
    # with a black rectangle.  That corrupts pixel similarity even though the
    # Word image is visibly the same slide.  For the first Word slide only,
    # use the document order as a guarded fallback: after a detected course
    # cover, the first content image is slide 2 when the next image already
    # has a reliable, later visual match.  This avoids asking teachers to
    # confirm an otherwise obvious mapping while never guessing a middle page.
    if len(result) >= 2 and not result[0][3] and result[1][3] and result[1][1] >= 4 and _is_cover_slide(ppt_path, 1):
        part, _page, _confidence, _accepted, choices = result[0]
        if all(choice["slide_number"] != 2 for choice in choices):
            choices = [{"slide_number": 2, "score": None, "reason": "sequence_after_cover"}, *choices]
        result[0] = (part, 2, 0.8, True, choices)
    return result


def create_mappings(chapter_id: str) -> list[dict]:
    files = db.rows("SELECT * FROM files WHERE chapter_id=?",(chapter_id,))
    ppt = next((f for f in files if f["extension"] in {".pptx",".ppt"}),None)
    word = next((f for f in files if f["extension"] in {".docx",".doc"}),None)
    if not ppt or not word:
        raise RuntimeError("该章节必须同时包含 PPT 和 Word 才能建立同步映射")
    if ppt["status"] != "published" or word["status"] != "published":
        raise RuntimeError("请先分别确认发布本章节的 PPT 和 Word；同步会以两份已确认版本作为输入")
    ppt_path = Path(_latest(ppt["id"],"published")["path"])
    original_ppt_path = Path(_latest(ppt["id"],"working","original")["path"])
    word_path = Path(_latest(word["id"],"working","original")["path"])
    visual = _visual_mappings(word_path, original_ppt_path, file_dir(word["id"]) / "mapping_previews")
    if not visual:
        raise RuntimeError("Word 中未发现可映射的幻灯片图片")
    now = db.utcnow()
    with db.connect() as conn:
        # A new visual pass supersedes any older sequence-based guesses,
        # including media that is not a full-slide candidate.
        conn.execute("DELETE FROM ppt_word_mappings WHERE chapter_id=? AND word_file_id=?", (chapter_id, word["id"]))
        for part, slide_no, confidence, confirmed, choices in visual:
            decision = "automatic" if confirmed else "pending"
            conn.execute("INSERT INTO ppt_word_mappings(id,chapter_id,ppt_file_id,word_file_id,slide_number,media_part,confidence,confirmed,candidates_json,decision,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                         (db.new_id(),chapter_id,ppt["id"],word["id"],slide_no,part,confidence,int(confirmed),json.dumps(choices,ensure_ascii=False),decision,now))
    return db.rows("SELECT * FROM ppt_word_mappings WHERE chapter_id=? ORDER BY slide_number",(chapter_id,))


def _changed_word_pages(before: Path, after: Path, file_id: str, review_version_id: str) -> list[int]:
    """Render both immutable Word versions and identify pages changed by sync."""
    root = file_dir(file_id) / "previews"
    before_cache = root / "sync-base"
    after_cache = root / review_version_id
    pages = max(document_page_count(before, before_cache), document_page_count(after, after_cache))
    changed = []
    for page in range(1, pages + 1):
        try:
            left = Image.open(render_preview(before, page, before_cache)).convert("RGB")
            right = Image.open(render_preview(after, page, after_cache)).convert("RGB")
            if left.size != right.size:
                changed.append(page)
                continue
            difference = ImageChops.difference(left, right)
            # Anti-aliasing can change a handful of pixels during a Word PDF
            # export.  Require a visible average difference before marking it.
            if sum(ImageStat.Stat(difference).mean) / 3 > 1.0:
                changed.append(page)
        except Exception:
            # Never pretend that a page was verified if its preview could not
            # be rendered.  It remains visible in the review screen.
            changed.append(page)
    return changed


def apply_sync(chapter_id: str) -> str:
    mappings = db.rows("SELECT * FROM ppt_word_mappings WHERE chapter_id=? ORDER BY slide_number",(chapter_id,))
    if not mappings:
        mappings = create_mappings(chapter_id)
    resolved = [mapping for mapping in mappings if mapping["confirmed"]]
    if not resolved:
        raise RuntimeError("没有可可靠识别的 Word 幻灯片图片，未生成同步版本")
    ppt = db.row("SELECT * FROM files WHERE id=?", (resolved[0]["ppt_file_id"],))
    word = db.row("SELECT * FROM files WHERE id=?", (resolved[0]["word_file_id"],))
    if not ppt or not word or ppt["status"] != "published" or word["status"] != "published":
        raise RuntimeError("PPT 和 Word 都必须先确认发布，才能生成同步 Word 交付版")
    # Synchronization must use the completed translation/review versions.  The
    # original pair is immutable input and would otherwise put English slides
    # back into an already translated Word chapter.
    ppt_path = Path(_latest(resolved[0]["ppt_file_id"],"published")["path"])
    published_word_versions = db.rows("SELECT * FROM file_versions WHERE file_id=? AND version_type='published' ORDER BY sequence DESC",
                                      (resolved[0]["word_file_id"],))
    word_version = next((version for version in published_word_versions if not (version.get("metadata") or {}).get("ppt_sync")), None)
    if not word_version:
        raise RuntimeError("未找到 Word 的独立已发布版本，不能用已同步文件再次覆盖")
    word_path = Path(word_version["path"])
    export_dir = file_dir(resolved[0]["word_file_id"])/"slide_exports"
    export_dir.mkdir(parents=True,exist_ok=True)
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        app = win32com.client.DispatchEx("PowerPoint.Application")
        deck = app.Presentations.Open(str(ppt_path.resolve()),WithWindow=False)
        replacements = {}
        for mapping in resolved:
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
    output = file_dir(resolved[0]["word_file_id"])/"published"/f"{word_path.stem}_synced.docx"
    output.parent.mkdir(parents=True,exist_ok=True)
    if output.exists():
        output = output.with_name(f"{output.stem}_{db.new_id()[:8]}.docx")
    with zipfile.ZipFile(word_path) as src, zipfile.ZipFile(output,"x",zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            dst.writestr(info,replacements.get(info.filename,src.read(info.filename)))
    pending_count = sum(1 for mapping in mappings if not mapping["confirmed"] and mapping.get("decision") != "keep_original")
    version_id = add_version(resolved[0]["word_file_id"], "sync_review", output, {"ppt_sync": True,
                             "sync_review": True, "mapping_count": len(resolved),
                             "unmapped_count": len(mappings) - len(resolved), "pending_count": pending_count,
                             "ppt_version": _latest(resolved[0]["ppt_file_id"], "published")["id"],
                             "word_base_version": word_version["id"]})
    changed_pages = _changed_word_pages(word_path, output, resolved[0]["word_file_id"], version_id)
    version = db.row("SELECT metadata_json FROM file_versions WHERE id=?", (version_id,))
    metadata = version.get("metadata") or {}
    metadata["changed_pages"] = changed_pages
    db.execute("UPDATE file_versions SET metadata_json=? WHERE id=?", (json.dumps(metadata, ensure_ascii=False), version_id))
    return version_id
