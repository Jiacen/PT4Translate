from __future__ import annotations

import io
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import zipfile
from pathlib import Path

from PIL import Image, ImageChops, ImageOps, ImageStat

from . import db
from .documents import document_page_count, render_preview
from .pipeline import add_version
from .storage import file_dir

_mapping_lock = threading.RLock()
_page_index_lock = threading.RLock()


def _pair(chapter_id: str) -> tuple[dict, dict]:
    files = db.rows("SELECT * FROM files WHERE chapter_id=? ORDER BY created_at,id", (chapter_id,))
    ppt = next((f for f in files if f["extension"] in {".pptx", ".ppt"}), None)
    word = next((f for f in files if f["extension"] in {".docx", ".doc"}), None)
    if not ppt or not word:
        raise RuntimeError("该章节必须同时包含 PPT 和 Word 才能建立同步映射")
    return ppt, word


def word_base_version(file_id: str) -> dict:
    versions = db.rows("SELECT * FROM file_versions WHERE file_id=? AND version_type='published' ORDER BY sequence DESC", (file_id,))
    version = next((v for v in versions if not (v.get("metadata") or {}).get("ppt_sync")), None)
    if not version:
        raise RuntimeError("未找到 Word 的独立已发布版本")
    return version


def _source_signature(ppt: dict, word: dict) -> str:
    versions = [_latest(f["id"], "working", "original") for f in (ppt, word)]
    return json.dumps([(f["id"], v["checksum"]) for f, v in zip((ppt, word), versions)], separators=(",", ":"))


def _ensure_batch(chapter_id: str, mappings: list[dict]) -> dict | None:
    if not mappings:
        return None
    ppt, word = _pair(chapter_id)
    if any(m["ppt_file_id"] != ppt["id"] or m["word_file_id"] != word["id"] for m in mappings):
        raise RuntimeError("本章文件配对已变化，请先核对；原有人工确认已保留")
    signature = _source_signature(ppt, word)
    initial = sum(m["decision"] != "automatic" for m in mappings)
    with db.connect() as conn:
        conn.execute("INSERT OR IGNORE INTO mapping_batches VALUES(?,?,?,?,?,?)",
                     (chapter_id, ppt["id"], word["id"], signature, initial, db.utcnow()))
    batch = db.row("SELECT * FROM mapping_batches WHERE chapter_id=?", (chapter_id,))
    if batch["source_signature"] != signature or any(m["ppt_file_id"] != ppt["id"] or m["word_file_id"] != word["id"] for m in mappings):
        raise RuntimeError("本章源文件已变化，请先核对文件配对；原有人工确认已保留")
    return batch


def mapping_signature(mappings: list[dict]) -> str:
    payload = sorted((m["id"], m["slide_number"], m["confirmed"], m["decision"]) for m in mappings)
    return hashlib.sha256(json.dumps(payload, separators=(",", ":")).encode()).hexdigest()


def sync_needs_apply(chapter_id: str, version: dict | None = None) -> bool:
    mappings = db.rows("SELECT * FROM ppt_word_mappings WHERE chapter_id=?", (chapter_id,))
    if not mappings:
        return version is None
    if any(not m["confirmed"] and m["decision"] != "keep_original" for m in mappings):
        return True
    if version is None:
        version = db.row("SELECT * FROM file_versions WHERE file_id=? AND version_type='sync_review' ORDER BY sequence DESC LIMIT 1", (mappings[0]["word_file_id"],))
    if not version:
        return True
    metadata = version.get("metadata") or {}
    if metadata.get("mapping_choices_changed"):
        return True
    if metadata.get("mapping_signature") and metadata["mapping_signature"] != mapping_signature(mappings):
        return True
    try:
        return (metadata.get("ppt_version") != _latest(mappings[0]["ppt_file_id"], "published")["id"] or
                metadata.get("word_base_version") != word_base_version(mappings[0]["word_file_id"])["id"])
    except RuntimeError:
        return True


def _page_index_path(word_id: str) -> tuple[dict, Path]:
    version = word_base_version(word_id)
    return version, file_dir(word_id) / "mapping_pages" / f"word-pages-v1-{version['checksum']}.json"


def word_page_index(word_id: str, *, generate: bool = False) -> dict:
    if not generate:
        return _word_page_index(word_id)
    with _page_index_lock:
        return _word_page_index(word_id, generate=generate)


def _word_page_index(word_id: str, *, generate: bool = False) -> dict:
    version, target = _page_index_path(word_id)
    if target.exists():
        try:
            data = json.loads(target.read_text(encoding="utf-8"))
            if data.get("status") in {"ready", "partial"} or not generate:
                return data
        except (ValueError, OSError):
            pass
    if not generate:
        return {"status": "not_loaded", "pages": {}, "locations": {}}
    target.parent.mkdir(parents=True, exist_ok=True)
    if getattr(sys, "frozen", False):
        command = [sys.executable, "--index-word-images", version["path"], str(target)]
    else:
        helper = Path(__file__).resolve().parents[1] / "scripts" / "index_word_images.py"
        command = [sys.executable, str(helper), version["path"], str(target)]
    try:
        run = subprocess.run(command, capture_output=True, timeout=180, check=False,
                             creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0)
        if run.returncode == 0 and target.exists():
            return json.loads(target.read_text(encoding="utf-8"))
    except (OSError, ValueError, subprocess.TimeoutExpired):
        pass
    data = {"status": "unavailable", "pages": {}, "locations": {},
            "message": "未能读取 Word 实际页码，请确认本机已安装 Word 后重试。"}
    target.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return data


def pending_review(chapter_id: str, limit: int | None = None) -> dict:
    mappings = db.rows("SELECT * FROM ppt_word_mappings WHERE chapter_id=?", (chapter_id,))
    batch = _ensure_batch(chapter_id, mappings)
    pending = [m for m in mappings if not m["confirmed"] and m["decision"] != "keep_original"]
    completed = sum(m["decision"] in {"manual", "keep_original"} for m in mappings)
    result = {"initialized": bool(batch), "mappings": [], "progress": {
        "initial_pending": batch["initial_pending"] if batch else 0,
        "completed": completed, "remaining": len(pending),
        "automatic": sum(m["decision"] == "automatic" for m in mappings)}}
    if not batch:
        return result
    ppt, word = _pair(chapter_id)
    index = word_page_index(word["id"])
    for mapping in pending:
        mapping["word_pages"] = index.get("pages", {}).get(mapping["media_part"], [])
        mapping["word_locations"] = index.get("locations", {}).get(mapping["media_part"], [])
    pending.sort(key=lambda m: (min(m["word_pages"], default=10**9),
                               min((loc["position"] for loc in m["word_locations"]), default=10**9),
                               int(re.search(r"image(\d+)", m["media_part"], re.I).group(1)) if re.search(r"image(\d+)", m["media_part"], re.I) else 0))
    result.update(mappings=pending[:limit] if limit else pending, word_file_id=word["id"],
                  word_name=word["original_name"], ppt_name=ppt["original_name"],
                  ppt_file_id=ppt["id"], ppt_version_id=_latest(ppt["id"], "published")["id"],
                  ppt_pages=_slide_count(Path(_latest(ppt["id"], "published")["path"])),
                  page_index_status=index["status"], page_index_message=index.get("message"),
                  needs_apply=sync_needs_apply(chapter_id))
    return result


def save_mapping_decision(mapping_id: str, slide_number: int | None, keep_original: bool) -> dict:
    if bool(slide_number) == bool(keep_original):
        raise RuntimeError("请选择一个 PPT 页，或选择保留 Word 原图")
    mapping = db.row("SELECT * FROM ppt_word_mappings WHERE id=?", (mapping_id,))
    if not mapping:
        raise RuntimeError("图片映射不存在")
    if slide_number is not None and not 1 <= slide_number <= _slide_count(Path(_latest(mapping["ppt_file_id"], "published")["path"])):
        raise RuntimeError("PPT 页码超出文件的实际范围")
    decision = "keep_original" if keep_original else "manual"
    with db.connect() as conn:
        conn.execute("BEGIN IMMEDIATE")
        existing = dict(conn.execute("SELECT * FROM ppt_word_mappings WHERE id=?", (mapping_id,)).fetchone())
        if existing["confirmed"] or existing["decision"] == "keep_original":
            if existing["decision"] == decision and (keep_original or existing["slide_number"] == slide_number):
                return db.row("SELECT * FROM ppt_word_mappings WHERE id=?", (mapping_id,))
            raise RuntimeError("这张图片已确认，请关闭后重新打开，查看最新进度")
        conn.execute("UPDATE ppt_word_mappings SET slide_number=?,confirmed=?,decision=?,updated_at=? WHERE id=?",
                     (slide_number or existing["slide_number"], int(not keep_original), decision, db.utcnow(), mapping_id))
        version = conn.execute("SELECT id,metadata_json FROM file_versions WHERE file_id=? AND version_type='sync_review' ORDER BY sequence DESC LIMIT 1", (mapping["word_file_id"],)).fetchone()
        if version:
            metadata = json.loads(version["metadata_json"] or "{}")
            metadata.update(mapping_choices_changed=True, sync_review_confirmed=False)
            conn.execute("UPDATE file_versions SET metadata_json=? WHERE id=?", (json.dumps(metadata, ensure_ascii=False), version["id"]))
    return db.row("SELECT * FROM ppt_word_mappings WHERE id=?", (mapping_id,))


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
    with _mapping_lock:
        return _create_mappings(chapter_id)


def _create_mappings(chapter_id: str) -> list[dict]:
    files = db.rows("SELECT * FROM files WHERE chapter_id=?",(chapter_id,))
    ppt = next((f for f in files if f["extension"] in {".pptx",".ppt"}),None)
    word = next((f for f in files if f["extension"] in {".docx",".doc"}),None)
    if not ppt or not word:
        raise RuntimeError("该章节必须同时包含 PPT 和 Word 才能建立同步映射")
    if ppt["status"] != "published" or word["status"] != "published":
        raise RuntimeError("请先分别确认发布本章节的 PPT 和 Word；同步会以两份已确认版本作为输入")
    existing = db.rows("SELECT * FROM ppt_word_mappings WHERE chapter_id=?", (chapter_id,))
    if existing:
        _ensure_batch(chapter_id, existing)
        return existing
    ppt_path = Path(_latest(ppt["id"],"published")["path"])
    original_ppt_path = Path(_latest(ppt["id"],"working","original")["path"])
    word_path = Path(_latest(word["id"],"working","original")["path"])
    visual = _visual_mappings(word_path, original_ppt_path, file_dir(word["id"]) / "mapping_previews")
    if not visual:
        raise RuntimeError("Word 中未发现可映射的幻灯片图片")
    now = db.utcnow()
    with db.connect() as conn:
        for part, slide_no, confidence, confirmed, choices in visual:
            decision = "automatic" if confirmed else "pending"
            conn.execute("INSERT INTO ppt_word_mappings(id,chapter_id,ppt_file_id,word_file_id,slide_number,media_part,confidence,confirmed,candidates_json,decision,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                         (db.new_id(),chapter_id,ppt["id"],word["id"],slide_no,part,confidence,int(confirmed),json.dumps(choices,ensure_ascii=False),decision,now))
        conn.execute("INSERT INTO mapping_batches VALUES(?,?,?,?,?,?)",
                     (chapter_id, ppt["id"], word["id"], _source_signature(ppt, word),
                      sum(not confirmed for _, _, _, confirmed, _ in visual), now))
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
    with _mapping_lock:
        return _apply_sync(chapter_id)


def _apply_sync(chapter_id: str) -> str:
    mappings = db.rows("SELECT * FROM ppt_word_mappings WHERE chapter_id=? ORDER BY slide_number",(chapter_id,))
    if not mappings:
        mappings = create_mappings(chapter_id)
    _ensure_batch(chapter_id, mappings)
    if any(not m["confirmed"] and m["decision"] != "keep_original" for m in mappings):
        raise RuntimeError("请先逐项确认所有待确认图片，再生成同步 Word")
    if not sync_needs_apply(chapter_id):
        return _latest(mappings[0]["word_file_id"], "sync_review")["id"]
    resolved = [mapping for mapping in mappings if mapping["confirmed"]]
    ppt = db.row("SELECT * FROM files WHERE id=?", (mappings[0]["ppt_file_id"],))
    word = db.row("SELECT * FROM files WHERE id=?", (mappings[0]["word_file_id"],))
    if not ppt or not word or ppt["status"] != "published" or word["status"] != "published":
        raise RuntimeError("PPT 和 Word 都必须先确认发布，才能生成同步 Word 交付版")
    # Synchronization must use the completed translation/review versions.  The
    # original pair is immutable input and would otherwise put English slides
    # back into an already translated Word chapter.
    ppt_path = Path(_latest(ppt["id"],"published")["path"])
    published_word_versions = db.rows("SELECT * FROM file_versions WHERE file_id=? AND version_type='published' ORDER BY sequence DESC",
                                      (word["id"],))
    word_version = next((version for version in published_word_versions if not (version.get("metadata") or {}).get("ppt_sync")), None)
    if not word_version:
        raise RuntimeError("未找到 Word 的独立已发布版本，不能用已同步文件再次覆盖")
    word_path = Path(word_version["path"])
    export_dir = file_dir(word["id"])/"slide_exports"
    export_dir.mkdir(parents=True,exist_ok=True)
    replacements = _export_replacements(ppt_path, resolved, export_dir) if resolved else {}
    output = file_dir(word["id"])/"published"/f"{word_path.stem}_synced.docx"
    output.parent.mkdir(parents=True,exist_ok=True)
    if output.exists():
        output = output.with_name(f"{output.stem}_{db.new_id()[:8]}.docx")
    with zipfile.ZipFile(word_path) as src, zipfile.ZipFile(output,"x",zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            dst.writestr(info,replacements.get(info.filename,src.read(info.filename)))
    version_id = add_version(word["id"], "sync_review", output, {"ppt_sync": True,
                             "sync_review": True, "mapping_count": len(resolved),
                             "unmapped_count": len(mappings) - len(resolved), "pending_count": 0,
                             "mapping_signature": mapping_signature(mappings),
                             "ppt_version": _latest(ppt["id"], "published")["id"],
                             "word_base_version": word_version["id"]})
    changed_pages = _changed_word_pages(word_path, output, word["id"], version_id) if resolved else []
    version = db.row("SELECT metadata_json FROM file_versions WHERE id=?", (version_id,))
    metadata = version.get("metadata") or {}
    metadata["changed_pages"] = changed_pages
    db.execute("UPDATE file_versions SET metadata_json=? WHERE id=?", (json.dumps(metadata, ensure_ascii=False), version_id))
    return version_id


def _export_replacements(ppt_path: Path, resolved: list[dict], export_dir: Path) -> dict:
    app = deck = None
    initialized = False
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        initialized = True
        app = win32com.client.DispatchEx("PowerPoint.Application")
        deck = app.Presentations.Open(str(ppt_path.resolve()),WithWindow=False)
        replacements = {}
        for mapping in resolved:
            ext = Path(mapping["media_part"]).suffix.lstrip(".").upper()
            export_format = "EMF" if ext in {"EMF","WMF"} else ("JPG" if ext in {"JPG","JPEG"} else "PNG")
            target = export_dir/f"slide_{mapping['slide_number']}.{ext.lower()}"
            deck.Slides(mapping["slide_number"]).Export(str(target.resolve()),export_format)
            replacements[mapping["media_part"]] = target.read_bytes()
        return replacements
    except Exception as exc:
        raise RuntimeError(f"PowerPoint 页面渲染失败：{exc}") from exc
    finally:
        if deck is not None:
            try: deck.Close()
            except Exception: pass
        if app is not None:
            try: app.Quit()
            except Exception: pass
        if initialized:
            pythoncom.CoUninitialize()
