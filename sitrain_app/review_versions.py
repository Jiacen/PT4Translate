"""Immutable review versions and page-scoped preview cache selection."""
from functools import wraps
from pathlib import Path
import threading
import shutil

from fastapi import HTTPException

from . import db
from .storage import file_dir
from .ppt_review import edit_pptx, edit_pptx_unit_texts, page_objects
from .terminology import normalize

edit_lock = threading.RLock()


def serialized(fn):
    @wraps(fn)
    def call(*args, **kwargs):
        with edit_lock:
            return fn(*args, **kwargs)
    return call


def current_version(file_id):
    file = db.row("SELECT status FROM files WHERE id=?", (file_id,))
    # A sync candidate is stored as an immutable published artifact, but it
    # must not replace the active review document until the file is explicitly
    # published.  This preserves a safe, editable machine/review version when
    # an unverified mapping needs correction.
    if not file or file["status"] != "published":
        order = "CASE version_type WHEN 'review' THEN 0 WHEN 'sync_review' THEN 1 WHEN 'machine' THEN 2 WHEN 'published' THEN 3 WHEN 'working' THEN 4 ELSE 5 END"
    else:
        # A PPT→Word result is a distinct, immutable review candidate.  It
        # becomes the active preview immediately so teachers inspect exactly
        # the Word file that would be delivered, before accepting it.
        order = "CASE version_type WHEN 'sync_review' THEN 0 WHEN 'published' THEN 1 WHEN 'review' THEN 2 WHEN 'machine' THEN 3 WHEN 'working' THEN 4 ELSE 5 END"
    version = db.row("""SELECT * FROM file_versions WHERE file_id=?
        AND version_type IN ('review','sync_review','published','machine','working','original')
        ORDER BY """ + order + ",sequence DESC LIMIT 1", (file_id,))
    if not version:
        raise HTTPException(404, "文件版本不存在")
    return version


def preview_source(version, page):
    # An unchanged page reuses its ancestor's cache; changed pages render from the
    # exact immutable version. Never reuse a previous rendering for edited content.
    seen = set()
    while version["id"] not in seen:
        seen.add(version["id"])
        meta = version.get("metadata") or {}
        if meta.get("edited_page") == page or page in meta.get("edited_pages", []) or not meta.get("base_version"):
            break
        parent = db.row("SELECT * FROM file_versions WHERE id=? AND file_id=?",
                        (meta["base_version"], version["file_id"]))
        if not parent:
            break
        version = parent
    root = file_dir(version["file_id"]) / "previews"
    if Path(version["path"]).suffix.lower() == ".docx":
        # Each Word revision has a separate PDF/page cache.  This keeps the
        # preview faithful immediately after a text-unit edit.
        return Path(version["path"]), root / version["id"]
    # Only the page modified by this review version needs a new render. All other
    # pages keep the completed machine-version cache, including legacy published
    # files that predate review-version lineage metadata.
    meta = version.get("metadata") or {}
    changed_page = meta.get("edited_page") == page or page in meta.get("edited_pages", [])
    cache = root / version["id"] if changed_page else root
    return Path(version["path"]), cache


def _payload_matches(reference, unit):
    """Match only the structural locator emitted by the PPT extractor.

    Visible source text is intentionally not involved: common labels such as
    "Start" or "OK" can appear many times in one deck.
    """
    payload = unit.get("source_payload") or {}
    if payload.get("type") != reference.get("type"):
        return False
    fields = ("text_box_index", "table_index", "row_index", "cell_index", "paragraph_index")
    return all(payload.get(field) == reference.get(field)
               for field in fields if field in reference)


def object_unit_matches(file_id, version, page, object_id):
    """Return unambiguous unit-to-paragraph matches for one PPT object."""
    source = Path(version["path"])
    data = page_objects(str(source), source.stat().st_mtime_ns, page)
    obj = next((item for item in data["objects"] if item["id"] == object_id), None)
    if not obj or obj["kind"] != "text":
        return obj, []
    units = db.rows("SELECT * FROM translation_units WHERE file_id=? AND page=?", (file_id, page))
    matches = []
    for reference in obj.get("unit_refs", []):
        candidates = [unit for unit in units if _payload_matches(reference, unit)]
        if len(candidates) == 1:
            matches.append({"paragraph_index": reference["paragraph_index"], "unit": candidates[0]})
    return obj, matches


def review_context(file_id, page, object_id):
    version = current_version(file_id)
    obj, matches = object_unit_matches(file_id, version, page, object_id)
    if not obj:
        raise HTTPException(404, "未找到页面对象")
    # The first version intentionally requires a one-paragraph object.  This
    # makes an "apply everywhere" correction a precise unit-level operation,
    # instead of accidentally applying one paragraph to an entire text box.
    trusted = len(obj.get("paragraphs", [])) == 1 and len(matches) == 1
    unit = matches[0]["unit"] if trusted else None
    return {"version_id": version["id"], "trusted": trusted,
            "reason": "" if trusted else "此对象无法唯一对应一个原始翻译单元，不能安全批量应用。",
            "source": unit.get("source_text") if unit else "",
            "target": obj["paragraphs"][0] if trusted else "",
            "unit_id": unit.get("id") if unit else None,
            "memory_scope": "project" if unit and db.scalar("SELECT project_id FROM files WHERE id=?", (file_id,)) else "local"}


def _record_text_changes(file_id, page, object_id, version, paragraphs):
    """Keep direct PPT edits and the audit/translation-unit state aligned."""
    obj, matches = object_unit_matches(file_id, version, page, object_id)
    if not obj:
        return
    for match in matches:
        index, unit = match["paragraph_index"], match["unit"]
        if index >= len(paragraphs):
            continue
        target = paragraphs[index]
        before = obj["paragraphs"][index]
        if target == before:
            continue
        db.execute("UPDATE translation_units SET human_text=?,status='modified',revision=revision+1,updated_at=? WHERE id=?",
                   (target, db.utcnow(), unit["id"]))
        db.execute("INSERT INTO review_changes(id,file_id,page,object_id,unit_id,source_text,before_text,after_text,change_type,created_at) VALUES(?,?,?,?,?,?,?,?,?,?)",
                   (db.new_id(), file_id, page, object_id, unit["id"], unit["source_text"], before, target, "text", db.utcnow()))


def _write_text_edit(file_id, page, object_id, base_version, paragraphs):
    from .pipeline import add_version
    version = current_version(file_id)
    if version["id"] != base_version:
        raise HTTPException(409, "文件已被修改，请重新选择页面后再保存")
    source = Path(version["path"])
    output = file_dir(file_id) / "review" / db.new_id() / source.name
    try:
        edit_pptx(source, output, page, object_id, paragraphs=paragraphs)
    except (ValueError, OSError) as exc:
        raise HTTPException(422, str(exc)) from exc
    _record_text_changes(file_id, page, object_id, version, paragraphs)
    return add_version(file_id, "review", output, {"base_version": base_version,
                       "edited_page": page, "object_id": object_id, "edit_kind": "text"})


def _object_for_unit(file_id, version, unit):
    """Find a PPT canvas object from an extraction locator, if it is unique."""
    page = unit.get("page")
    if not page:
        return None
    try:
        source = Path(version["path"])
        data = page_objects(str(source), source.stat().st_mtime_ns, page)
    except (OSError, ValueError):
        return None
    found = []
    for obj in data["objects"]:
        for reference in obj.get("unit_refs", []):
            if _payload_matches(reference, unit):
                found.append(obj["id"])
    return found[0] if len(found) == 1 else None


@serialized
def apply_safe_machine_term_repairs(file_id, units):
    """Merge repaired machine terms into a reviewed PPT without touching edits.

    A review version is the teacher's working copy.  We only patch paragraph
    locators that can be proved and whose *entire canvas object* has never been
    edited.  Speaker notes have no browser editing surface, so their structural
    paragraphs can be safely refreshed too.  Everything uncertain is skipped
    and remains visible to QA instead of risking a silent overwrite.
    """
    if not units:
        return {"version_id": None, "applied": 0, "skipped": 0}
    version = current_version(file_id)
    source = Path(version["path"])
    if source.suffix.lower() != ".pptx":
        return {"version_id": None, "applied": 0, "skipped": len(units)}
    edited = {(item["page"], item["object_id"])
              for item in db.rows("SELECT page,object_id FROM review_changes WHERE file_id=?", (file_id,))}
    updates, skipped = [], 0
    for unit in units:
        kind = unit.get("unit_type")
        if kind == "notes":
            updates.append({**unit, "text": unit.get("machine_text")})
            continue
        if kind not in {"text_paragraph", "table_cell_paragraph"}:
            skipped += 1
            continue
        object_id = _object_for_unit(file_id, version, unit)
        if not object_id or (unit.get("page"), object_id) in edited:
            skipped += 1
            continue
        updates.append({**unit, "text": unit.get("machine_text")})
    if not updates:
        return {"version_id": None, "applied": 0, "skipped": skipped}
    output = file_dir(file_id) / "review" / db.new_id() / source.name
    try:
        pages = edit_pptx_unit_texts(source, output, updates)
    except (ValueError, OSError) as exc:
        raise HTTPException(422, f"术语修正未写入 PPT：{exc}") from exc
    if not pages:
        return {"version_id": None, "applied": 0, "skipped": skipped + len(updates)}
    from .pipeline import add_version
    version_id = add_version(file_id, "review", output, {"base_version": version["id"],
                              "edited_pages": pages, "edit_kind": "safe_term_repair"})
    db.execute("UPDATE files SET updated_at=? WHERE id=?", (db.utcnow(), file_id))
    return {"version_id": version_id, "applied": len(updates), "skipped": skipped}


@serialized
def apply_same_source(file_id, page, object_id, base_version):
    """Apply the saved correction to structurally locatable exact-source units.

    This is deliberately a document-only helper.  It never changes a global
    term or another file, and it skips units whose PPT location cannot be
    proven from the original extraction payload.
    """
    file = db.row("SELECT * FROM files WHERE id=?", (file_id,))
    if not file or file["status"] != "review":
        raise HTTPException(409, "仅可对待校对文件执行此操作")
    version = current_version(file_id)
    if version["id"] != base_version:
        raise HTTPException(409, "文件已被修改，请刷新后再应用")
    context = review_context(file_id, page, object_id)
    if not context["trusted"]:
        raise HTTPException(422, context["reason"])
    desired = context["target"]
    source_text = context["source"]
    candidates = db.rows("SELECT * FROM translation_units WHERE file_id=? AND source_text=? ORDER BY page,stable_key",
                         (file_id, source_text))
    applied, skipped = 0, 0
    active_version = version["id"]
    for unit in candidates:
        unit_page = unit.get("page")
        if not unit_page:
            skipped += 1
            continue
        current = current_version(file_id)
        try:
            data = page_objects(str(Path(current["path"])), Path(current["path"]).stat().st_mtime_ns, unit_page)
        except (ValueError, OSError):
            skipped += 1
            continue
        locator = None
        for candidate_object in data["objects"]:
            for reference in candidate_object.get("unit_refs", []):
                if _payload_matches(reference, unit):
                    locator = (candidate_object, reference["paragraph_index"])
                    break
            if locator:
                break
        if not locator:
            skipped += 1
            continue
        target_object, paragraph_index = locator
        paragraphs = list(target_object["paragraphs"])
        if paragraphs[paragraph_index] == desired:
            continue
        paragraphs[paragraph_index] = desired
        active_version = _write_text_edit(file_id, unit_page, target_object["id"], current["id"], paragraphs)
        applied += 1
    if not applied:
        raise HTTPException(422, "没有可安全应用的同源文本；原有校对结果未变更")
    db.execute("UPDATE files SET updated_at=? WHERE id=?", (db.utcnow(), file_id))
    return {"version_id": active_version, "applied": applied, "skipped": skipped, "source": source_text}


@serialized
def save_review_memory(file_id, page, object_id, base_version):
    version = current_version(file_id)
    if version["id"] != base_version:
        raise HTTPException(409, "文件已被修改，请刷新后再加入翻译记忆")
    context = review_context(file_id, page, object_id)
    if not context["trusted"]:
        raise HTTPException(422, context["reason"])
    file = db.row("SELECT project_id FROM files WHERE id=?", (file_id,))
    db.execute("INSERT INTO translation_memory(id,project_id,source_norm,source,target,origin_unit_id,approved_at) VALUES(?,?,?,?,?,?,?) ON CONFLICT(project_id,source_norm) DO UPDATE SET target=excluded.target,origin_unit_id=excluded.origin_unit_id,approved_at=excluded.approved_at",
               (db.new_id(), file.get("project_id"), normalize(context["source"]), context["source"],
                context["target"], context["unit_id"], db.utcnow()))
    return {"source": context["source"], "target": context["target"], "scope": context["memory_scope"]}


@serialized
def reopen_review(file_id):
    from .pipeline import add_version
    file = db.row("SELECT * FROM files WHERE id=?", (file_id,))
    if not file:
        raise HTTPException(404, "文件不存在")
    if file['status'] == 'review':
        return {"version_id": current_version(file_id)['id'], "status": "review"}
    if file['status'] != 'published':
        raise HTTPException(409, "文件尚未发布或仍在处理中")
    source = db.row("SELECT * FROM file_versions WHERE file_id=? AND version_type='published' ORDER BY sequence DESC LIMIT 1", (file_id,))
    if not source or Path(source['path']).suffix.lower() != '.pptx':
        raise HTTPException(415, "当前继续校对仅支持 PPTX")
    output = file_dir(file_id)/'review'/db.new_id()/Path(source['path']).name
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source['path'], output)
    vid = add_version(file_id, 'review', output, {'base_version': source['id'], 'reopened': True})
    db.execute("UPDATE files SET status='review',updated_at=? WHERE id=?", (db.utcnow(),file_id))
    return {"version_id": vid, "status": "review"}


@serialized
def save_edit(file_id, page, object_id, base_version, paragraphs=None, image=None):
    from .pipeline import add_version
    file = db.row("SELECT * FROM files WHERE id=?", (file_id,))
    if not file:
        raise HTTPException(404, "文件不存在")
    if file["status"] != "review":
        raise HTTPException(409, "仅可修改待校对文件")
    if db.row("SELECT id FROM tasks WHERE file_id=? AND status IN ('queued','running')", (file_id,)):
        raise HTTPException(409, "文件正在处理，请完成后再校对")
    version = current_version(file_id)
    if paragraphs is not None:
        vid = _write_text_edit(file_id, page, object_id, base_version, paragraphs)
        db.execute("UPDATE files SET updated_at=? WHERE id=?", (db.utcnow(), file_id))
        return {"version_id": vid, "page": page}
    if version["id"] != base_version:
        raise HTTPException(409, "文件已被修改，请重新选择页面后再保存")
    source = Path(version["path"])
    if source.suffix.lower() != ".pptx":
        raise HTTPException(415, "当前原位编辑仅支持 PowerPoint 文件")
    output = file_dir(file_id) / "review" / db.new_id() / source.name
    try:
        edit_pptx(source, output, page, object_id, paragraphs, image)
    except (ValueError, OSError) as exc:
        raise HTTPException(422, str(exc)) from exc
    vid = add_version(file_id, "review", output, {"base_version": base_version,
                      "edited_page": page, "object_id": object_id, "edit_kind": "image" if image is not None else "text"})
    db.execute("UPDATE files SET updated_at=? WHERE id=?", (db.utcnow(), file_id))
    return {"version_id": vid, "page": page}
