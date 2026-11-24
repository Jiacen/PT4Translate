from __future__ import annotations

import json
import shutil
import threading
import traceback
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from . import db
from .documents import (TRANSLATE_OFFICE_IMAGE_TEXT, convert_legacy, document_page_count,
                        extract_units, preflight, prepare_preview_cache, rebuild)
from .providers import TranslationRequest, translate
from .ppt_fidelity import paragraph_line_groups
from .storage import checksum, file_dir
from .review_versions import serialized
from .terminology import Term, TerminologyEngine, normalize, residual_source_issues


# Large教材包 may contain dozens of Office files. Keep expensive API translation
# bounded while allowing every accepted file to remain durably queued in SQLite.
_translation_executor = ThreadPoolExecutor(max_workers=2, thread_name_prefix="sitrain-translate")
_project_queue_lock = threading.RLock()


def _version(file_id: str, version_type: str) -> dict | None:
    return db.row("SELECT * FROM file_versions WHERE file_id=? AND version_type=? ORDER BY sequence DESC LIMIT 1",(file_id,version_type))


def add_version(file_id: str, version_type: str, path: Path, metadata: dict | None = None) -> str:
    sequence = int(db.scalar("SELECT COALESCE(MAX(sequence),0)+1 FROM file_versions WHERE file_id=? AND version_type=?",(file_id,version_type)))
    vid = db.new_id()
    db.execute("INSERT INTO file_versions(id,file_id,version_type,sequence,path,checksum,size,metadata_json,created_at) VALUES(?,?,?,?,?,?,?,?,?)",
               (vid,file_id,version_type,sequence,str(path.resolve()),checksum(path),path.stat().st_size,json.dumps(metadata or {},ensure_ascii=False),db.utcnow()))
    return vid


def prepare_file(file_id: str) -> None:
    file = db.row("SELECT * FROM files WHERE id=?",(file_id,))
    source_v = _version(file_id,"original")
    source = Path(source_v["path"])
    try:
        work_dir = file_dir(file_id)/"working"
        working,warnings = convert_legacy(source,work_dir)
        if working != source:
            add_version(file_id,"working",working,{"conversion_warnings":warnings,"source_extension":source.suffix.lower()})
        else:
            working = source
        pf = preflight(working)
        if working.suffix.lower() == ".docx":
            pf["pages"] = document_page_count(working, file_dir(file_id) / "previews")
            if not pf["pages"]:
                pf["risks"].append("无法获得 Word 的实际分页；预览暂时只显示第一页")
        pf["conversion_warnings"] = warnings
        temp_dir = file_dir(file_id)/"pipeline"
        source_json, units = extract_units(working,temp_dir)
        pf["text_units"] = len(units)
        if source_json:
            pf["source_json"] = str(source_json.resolve())
        now = db.utcnow()
        with db.connect() as conn:
            for unit in units:
                uid = db.new_id()
                conn.execute("INSERT OR REPLACE INTO translation_units(id,file_id,stable_key,page,object_ref,unit_type,source_text,status,confidence,source_payload_json,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                             (uid,file_id,unit["stable_key"],unit.get("page"),unit.get("object_ref"),unit.get("unit_type","text"),unit["source_text"],"pending_translation",unit.get("confidence"),json.dumps(unit.get("source_payload",{}),ensure_ascii=False),now))
            conn.execute("UPDATE files SET status='ready',preflight_json=?,error=NULL,updated_at=? WHERE id=?",(json.dumps(pf,ensure_ascii=False),now,file_id))
        # A standalone file remains one-click: it can start as soon as its
        # preflight is complete.  Project files deliberately stop here until
        # the teacher has checked the detected chapter structure and explicitly
        # starts the project.
        if file.get("project_id"):
            schedule_project_translation(file["project_id"])
        else:
            start_automatic_translation(file_id)
    except Exception as exc:
        db.execute("UPDATE files SET status='failed',error=?,updated_at=? WHERE id=?",(str(exc),db.utcnow(),file_id))
        raise


def start_automatic_translation(file_id: str) -> None:
    task_id = _enqueue_translation(file_id)
    if task_id:
        schedule_translation(task_id)


def schedule_translation(task_id: str, provider_id: str | None = None) -> None:
    """Submit one persisted task to the bounded local translation worker."""
    _translation_executor.submit(run_translation, task_id, provider_id)


def _enqueue_translation(file_id: str, provider_id: str | None = None) -> str | None:
    """Create the durable task for one ready file, without starting it twice."""
    revision = db.scalar("SELECT COALESCE(MAX(revision),0) FROM translation_units WHERE file_id=?", (file_id,))
    key = f"translate:{file_id}:{revision}:{provider_id or 'default'}"
    existing = db.row("SELECT id FROM tasks WHERE idempotency_key=?", (key,))
    if existing:
        return None
    task_id, now = db.new_id(), db.utcnow()
    db.execute("INSERT INTO tasks(id,file_id,task_type,status,idempotency_key,created_at,updated_at) VALUES(?,?,?,?,?,?,?)",
               (task_id, file_id, "translate", "queued", key, now, now))
    return task_id


def start_project_translation(project_id: str) -> None:
    """Start a teacher-confirmed project at its first runnable chapter."""
    project = db.row("SELECT * FROM projects WHERE id=?", (project_id,))
    if not project:
        raise RuntimeError("教材项目不存在")
    if project.get("workflow_status") == "translating":
        return
    incomplete = db.scalar("SELECT COUNT(*) FROM files WHERE project_id=? AND status='preflight'", (project_id,))
    if incomplete:
        raise RuntimeError("教材文件仍在整理中，请完成后再开始翻译")
    db.execute("UPDATE projects SET workflow_status='translating',updated_at=? WHERE id=?", (db.utcnow(), project_id))
    schedule_project_translation(project_id)


def schedule_project_translation(project_id: str) -> None:
    """Run exactly one chapter at a time, while allowing its Word/PPT files
    to occupy the two bounded file workers together.

    The next chapter is released only after every task in the current chapter
    has reached a terminal state. Failed chapters are left visible for the
    teacher but do not silently block the rest of the textbook.
    """
    with _project_queue_lock:
        project = db.row("SELECT * FROM projects WHERE id=?", (project_id,))
        if not project or project.get("workflow_status") != "translating":
            return
        active = db.row("""SELECT t.id FROM tasks t
            JOIN files f ON f.id=t.file_id
            WHERE f.project_id=? AND t.task_type='translate'
              AND t.status IN ('queued','running') LIMIT 1""", (project_id,))
        if active:
            return
        chapters = db.rows("SELECT * FROM chapters WHERE project_id=? ORDER BY position,created_at", (project_id,))
        for chapter in chapters:
            files = db.rows("SELECT * FROM files WHERE chapter_id=? ORDER BY created_at", (chapter["id"],))
            if not files or any(file["status"] == "preflight" for file in files):
                return
            ready = [file for file in files if file["status"] == "ready"]
            if not ready:
                # A failed chapter is intentionally skipped so that a malformed
                # legacy document cannot stall every later chapter.
                continue
            task_ids = [_enqueue_translation(file["id"]) for file in ready]
            for task_id in task_ids:
                if task_id:
                    schedule_translation(task_id)
            return
        db.execute("UPDATE projects SET workflow_status='completed',updated_at=? WHERE id=?", (db.utcnow(), project_id))


@serialized
def refresh_machine_version(file_id: str) -> str | None:
    if _version(file_id, "review"):
        raise RuntimeError("已有原位校对版本，不能用重新翻译覆盖人工修改")
    file = db.row("SELECT * FROM files WHERE id=?", (file_id,))
    if not file or file["status"] not in {"review", "published"}:
        return None
    source_v = _version(file_id, "working") or _version(file_id, "original")
    source_json_path = (file.get("preflight") or {}).get("source_json")
    units = db.rows("SELECT * FROM translation_units WHERE file_id=? ORDER BY CAST(stable_key AS INTEGER),stable_key", (file_id,))
    output = rebuild(Path(source_v["path"]), Path(source_json_path) if source_json_path else None, units,
                     file_dir(file_id) / "machine", file["source_language"], file["target_language"])
    version_id = add_version(file_id, "machine", output, {"generated_from": "machine_and_review"})
    try:
        # A revised Word is a new immutable version.  Its page images must not
        # reuse a prior version's cache, otherwise saved text appears unchanged.
        cache = file_dir(file_id) / "previews"
        if output.suffix.lower() == ".docx":
            cache /= version_id
        prepare_preview_cache(output, cache, int((file.get("preflight") or {}).get("pages") or 0))
    except Exception:
        pass
    return version_id


def _terms(project_id: str | None, source_language: str, target_language: str) -> tuple[list[dict],TerminologyEngine]:
    if project_id:
        records = db.rows("SELECT * FROM glossary_terms WHERE project_id IS NULL OR project_id=?",(project_id,))
    else:
        records = db.rows("SELECT * FROM glossary_terms WHERE project_id IS NULL")
    if source_language == "Chinese" and target_language == "English":
        records = [{**r, "source": r["target"], "target": r["source"], "aliases": r.get("chinese_aliases") or []} for r in records]
    terms = [Term(r["id"],r["source"],r["target"],tuple(r.get("aliases") or []),r["mode"],bool(r["case_sensitive"])) for r in records]
    return records,TerminologyEngine(terms)


@serialized
def revalidate_terminology_qa(project_id: str | None = None) -> dict[str, int]:
    """Refresh editable units' QA after a glossary change, without altering text."""
    file_sql = "SELECT * FROM files WHERE status IN ('review','published')"
    params: tuple = ()
    if project_id is not None:
        file_sql += " AND project_id=?"
        params = (project_id,)
    files = db.rows(file_sql, params)
    checked = blockers = 0
    now = db.utcnow()
    with db.connect() as conn:
        for file in files:
            _records, engine = _terms(file.get("project_id"), file["source_language"], file["target_language"])
            units = db.rows("SELECT * FROM translation_units WHERE file_id=? AND status NOT IN ('confirmed','skipped_image')", (file["id"],))
            for unit in units:
                translated = unit.get("human_text") or unit.get("machine_text")
                if translated is None:
                    continue
                issues = _qa_issues(unit, engine, translated)
                has_term_blocker = any(issue["code"] == "TERM_MISSING" and issue["severity"] == "blocker" for issue in issues)
                status = "terminology_issue" if has_term_blocker else ("pending_review" if unit["status"] == "terminology_issue" else unit["status"])
                conn.execute("UPDATE translation_units SET qa_json=?,status=?,updated_at=? WHERE id=?",
                             (json.dumps(issues, ensure_ascii=False), status, now, unit["id"]))
                checked += 1
                blockers += sum(1 for issue in issues if issue.get("severity") == "blocker")
    return {"files": len(files), "units": checked, "blockers": blockers}


def _qa_issues(unit: dict, engine: TerminologyEngine, translated: str) -> list[dict]:
    issues = engine.validate(unit["source_text"], translated) + residual_source_issues(unit["source_text"], translated)
    # Slide-layout alternative descriptions are non-visual accessibility
    # metadata.  They are not editable in the slide review canvas, so they
    # must be reported without making a teacher unable to publish a deck.
    payload = unit.get("source_payload") or {}
    is_word_field = unit.get("unit_type") == "header_footer" and bool(payload.get("field_info"))
    if unit.get("unit_type") == "ppt_alttext" or is_word_field:
        for issue in issues:
            if issue.get("severity") == "blocker":
                issue["severity"] = "warning"
                label = "PPT 辅助描述" if unit.get("unit_type") == "ppt_alttext" else "Word 自动字段"
                issue["message"] = label + "：" + issue["message"]
    return issues


def _has_required_term_blocker(issues: list[dict]) -> bool:
    return any(issue.get("code") == "TERM_MISSING" and issue.get("severity") == "blocker" for issue in issues)


def _retry_missing_terms(provider: dict, file: dict, engine: TerminologyEngine,
                         term_records: list[dict], memory: list[dict], units: list[dict],
                         translations: dict[str, str], progress=lambda _done, _total: None) -> tuple[dict[str, str], int]:
    """Retry only failed machine-translation units with protected terminology.

    This is deliberately a targeted retranslation rather than a blind text
    replacement: a missing brand/product term may have been translated into a
    natural-language word, so appending or replacing a guessed substring could
    corrupt the sentence. Human edits never enter this function.
    """
    retry_units, mappings = [], {}
    for unit in units:
        translated = translations.get(unit["stable_key"], "")
        if not _has_required_term_blocker(_qa_issues(unit, engine, translated)):
            continue
        protected, mapping = engine.protect(unit["source_text"])
        retry_units.append((unit["stable_key"], protected))
        mappings[unit["stable_key"]] = mapping
    if not retry_units:
        return translations, 0
    retried = translate(provider, TranslationRequest(retry_units, file["source_language"], file["target_language"],
                                                     term_records, memory, preserve_line_breaks=file.get("kind") == "powerpoint",
                                                     line_groups={u["stable_key"]: paragraph_line_groups(u.get("source_payload") or {}) for u in units}), progress)
    fixed = dict(translations)
    repaired = 0
    by_key = {unit["stable_key"]: unit for unit in units}
    for stable_key, translated in retried.items():
        candidate = engine.restore(translated, mappings[stable_key])
        if not _has_required_term_blocker(_qa_issues(by_key[stable_key], engine, candidate)):
            fixed[stable_key] = candidate
            repaired += 1
    return fixed, repaired


@serialized
def auto_repair_missing_terms(file_id: str) -> dict[str, int]:
    """Repair term blockers and safely merge them into an existing PPT review."""
    file = db.row("SELECT * FROM files WHERE id=?", (file_id,))
    if not file or file["status"] not in {"review", "published"}:
        raise RuntimeError("文件尚未完成翻译")
    if _version(file_id, "review") and file["extension"].lower() != ".pptx":
        raise RuntimeError("已有校对版本；当前安全术语合并暂仅支持 PowerPoint，请在校对页处理术语")
    provider = db.row("SELECT * FROM providers WHERE enabled=1 AND is_default=1 LIMIT 1")
    if not provider:
        raise RuntimeError("没有启用的默认 Provider")
    units = db.rows("SELECT * FROM translation_units WHERE file_id=? ORDER BY CAST(stable_key AS INTEGER),stable_key", (file_id,))
    candidates = [unit for unit in units if unit.get("machine_text") and not unit.get("human_text")]
    term_records, engine = _terms(file.get("project_id"), file["source_language"], file["target_language"])
    memory = db.rows("SELECT source,target FROM translation_memory WHERE project_id IS ? LIMIT 100", (file.get("project_id"),))
    translations = {unit["stable_key"]: unit["machine_text"] for unit in candidates}
    translations, repaired = _retry_missing_terms(provider, file, engine, term_records, memory, candidates, translations)
    repaired_units = []
    for unit in candidates:
        before = unit.get("machine_text") or ""
        after = translations[unit["stable_key"]]
        if (_has_required_term_blocker(_qa_issues(unit, engine, before)) and
                not _has_required_term_blocker(_qa_issues(unit, engine, after))):
            repaired_units.append({**unit, "machine_text": after})
    now = db.utcnow()
    with db.connect() as conn:
        for unit in candidates:
            translated = translations[unit["stable_key"]]
            issues = _qa_issues(unit, engine, translated)
            status = "terminology_issue" if _has_required_term_blocker(issues) else "pending_review"
            conn.execute("UPDATE translation_units SET machine_text=?,status=?,qa_json=?,revision=revision+1,updated_at=? WHERE id=?",
                         (translated, status, json.dumps(issues, ensure_ascii=False), now, unit["id"]))
    applied = skipped = 0
    if repaired:
        review = _version(file_id, "review")
        if review and file["extension"].lower() == ".pptx":
            # The reviewed PPT is the active teacher copy.  Merge only into
            # structurally proven, untouched canvas objects; never rebuild it.
            from .review_versions import apply_safe_machine_term_repairs
            merge = apply_safe_machine_term_repairs(file_id, repaired_units)
            applied, skipped = merge["applied"], merge["skipped"]
        elif not review:
            refresh_machine_version(file_id)
    remaining = sum(1 for unit in db.rows("SELECT * FROM translation_units WHERE file_id=?", (file_id,))
                    if _has_required_term_blocker(unit.get("qa") or []))
    return {"repaired": repaired, "applied": applied, "skipped": skipped, "remaining": remaining}


@serialized
def confirm_manual_term_exceptions(file_id: str) -> dict[str, int]:
    """Let the teacher explicitly retain a terminology choice they edited.

    This is deliberately a one-click *exception record*, not another automatic
    rewrite.  It is available only after all machine-owned term blockers have
    been resolved, so a publish can never bypass an unresolved machine error.
    """
    file = db.row("SELECT * FROM files WHERE id=?", (file_id,))
    if not file or file["status"] not in {"review", "published"}:
        raise RuntimeError("文件尚未完成翻译")
    units = db.rows("SELECT * FROM translation_units WHERE file_id=?", (file_id,))
    blockers = [unit for unit in units if _has_required_term_blocker(unit.get("qa") or [])]
    if not blockers:
        return {"confirmed": 0}
    manual = [unit for unit in blockers if unit.get("human_text")]
    if len(manual) != len(blockers):
        raise RuntimeError("仍有未人工校对的术语问题，请先使用“自动安全修正术语”")
    now = db.utcnow()
    with db.connect() as conn:
        for unit in manual:
            issues = []
            for issue in unit.get("qa") or []:
                if issue.get("code") == "TERM_MISSING" and issue.get("severity") == "blocker":
                    issues.append({**issue, "severity": "warning", "code": "TERM_EXCEPTION_CONFIRMED",
                                   "message": "老师已确认保留此处人工校对用词：" + issue.get("message", "术语例外")})
                else:
                    issues.append(issue)
            conn.execute("UPDATE translation_units SET qa_json=?,status='confirmed',revision=revision+1,updated_at=? WHERE id=?",
                         (json.dumps(issues, ensure_ascii=False), now, unit["id"]))
    return {"confirmed": len(manual)}


def run_translation(task_id: str, provider_id: str | None = None) -> None:
    task = db.row("SELECT * FROM tasks WHERE id=?",(task_id,))
    if not task or task["status"] != "queued":
        return
    file = db.row("SELECT * FROM files WHERE id=?",(task["file_id"],))
    db.execute("UPDATE tasks SET status='running',progress=0.01,updated_at=? WHERE id=?",(db.utcnow(),task_id))
    try:
        provider = db.row("SELECT * FROM providers WHERE id=?",(provider_id,)) if provider_id else db.row("SELECT * FROM providers WHERE enabled=1 AND is_default=1 LIMIT 1")
        if not provider:
            raise RuntimeError("没有启用的默认 Provider")
        units = db.rows("SELECT * FROM translation_units WHERE file_id=? ORDER BY CAST(stable_key AS INTEGER),stable_key",(file["id"],))
        if not TRANSLATE_OFFICE_IMAGE_TEXT:
            image_unit_ids = [unit["id"] for unit in units if unit["unit_type"] == "office_image_ocr"]
            if image_unit_ids:
                placeholders = ",".join("?" for _ in image_unit_ids)
                db.execute(f"UPDATE translation_units SET machine_text=NULL,human_text=NULL,status='skipped_image',qa_json='[]',updated_at=? WHERE id IN ({placeholders})",
                           (db.utcnow(), *image_unit_ids))
                units = [unit for unit in units if unit["unit_type"] != "office_image_ocr"]
        term_records,engine = _terms(file.get("project_id"),file["source_language"],file["target_language"])
        memory = db.rows("SELECT source,target FROM translation_memory WHERE project_id IS ? LIMIT 100",(file.get("project_id"),))
        pending, protected_map, direct = [],{},{}
        tm = {m["source_norm"]:m["target"] for m in db.rows("SELECT source_norm,target FROM translation_memory WHERE project_id IS ?",(file.get("project_id"),))}
        for unit in units:
            remembered = tm.get(normalize(unit["source_text"]))
            if remembered is not None and (file.get("kind") != "powerpoint" or all(remembered.count(c) == unit["source_text"].count(c) for c in ('\n', '\t'))):
                direct[unit["stable_key"]] = tm[normalize(unit["source_text"])]
                continue
            text,mapping = engine.protect(unit["source_text"])
            pending.append((unit["stable_key"],text))
            protected_map[unit["stable_key"]] = mapping
        def progress(done,total):
            db.execute("UPDATE tasks SET progress=?,detail=?,updated_at=? WHERE id=?",(0.05+0.8*(done/max(1,total)),f"已翻译 {done}/{total}",db.utcnow(),task_id))
        result = direct | (translate(provider,TranslationRequest(pending,file["source_language"],file["target_language"],term_records,memory,
                            preserve_line_breaks=file.get("kind") == "powerpoint",
                            line_groups={u["stable_key"]: paragraph_line_groups(u.get("source_payload") or {}) for u in units}),progress) if pending else {})
        def repair_progress(done, total):
            db.execute("UPDATE tasks SET progress=?,detail=?,updated_at=? WHERE id=?",
                       (0.88 + 0.07 * (done / max(1, total)), f"正在自动修正术语 {done}/{total}", db.utcnow(), task_id))
        result, repaired = _retry_missing_terms(provider, file, engine, term_records, memory, units, result, repair_progress)
        with db.connect() as conn:
            for unit in units:
                translated = engine.restore(result[unit["stable_key"]],protected_map.get(unit["stable_key"],{}))
                issues = _qa_issues(unit, engine, translated)
                status = "terminology_issue" if any(x["code"]=="TERM_MISSING" for x in issues) else "pending_review"
                conn.execute("UPDATE translation_units SET machine_text=?,status=?,qa_json=?,revision=revision+1,updated_at=? WHERE id=?",
                             (translated,status,json.dumps(issues,ensure_ascii=False),db.utcnow(),unit["id"]))
            conn.execute("UPDATE files SET status='review',error=NULL,updated_at=? WHERE id=?",(db.utcnow(),file["id"]))
            detail = "翻译完成，等待校对" + (f"；已自动修正 {repaired} 条术语" if repaired else "")
            conn.execute("UPDATE tasks SET status='completed',progress=1,detail=?,updated_at=? WHERE id=?",(detail,db.utcnow(),task_id))
        try:
            refresh_machine_version(file["id"])
        except Exception as exc:
            db.execute("UPDATE files SET error=?,updated_at=? WHERE id=?", (f"校对预览生成失败：{exc}", db.utcnow(), file["id"]))
    except Exception as exc:
        db.execute("UPDATE tasks SET status='failed',detail=?,updated_at=? WHERE id=?",(str(exc),db.utcnow(),task_id))
        db.execute("UPDATE files SET status='failed',error=?,updated_at=? WHERE id=?",(str(exc),db.utcnow(),file["id"]))
    finally:
        # The final file in a chapter releases the next teacher-confirmed
        # chapter.  A call from an earlier sibling sees the still-active task
        # and therefore cannot interleave chapters.
        if file and file.get("project_id"):
            schedule_project_translation(file["project_id"])


@serialized
def publish(file_id: str) -> str:
    file = db.row("SELECT * FROM files WHERE id=?",(file_id,))
    source_v = _version(file_id,"working") or _version(file_id,"original")
    source = Path(source_v["path"])
    units = db.rows("SELECT * FROM translation_units WHERE file_id=? ORDER BY CAST(stable_key AS INTEGER),stable_key",(file_id,))
    blockers = [q for u in units for q in (u.get("qa") or []) if q.get("severity")=="blocker" and u["status"]!="confirmed"]
    if blockers:
        raise RuntimeError(f"仍有 {len(blockers)} 个阻止发布的 QA 问题")
    reviewed = _version(file_id, "review")
    destination = file_dir(file_id)/"published"/db.new_id()
    if reviewed:
        destination.mkdir(parents=True, exist_ok=True)
        result = destination/Path(reviewed["path"]).name
        shutil.copyfile(reviewed["path"], result)
    else:
        source_json_path = (file.get("preflight") or {}).get("source_json")
        result = rebuild(source,Path(source_json_path) if source_json_path else None,units,destination,file["source_language"],file["target_language"])
    vid = add_version(file_id,"published",result,{"qa_blockers":0, **({"base_version":reviewed["id"]} if reviewed else {})})
    db.execute("UPDATE files SET status='published',updated_at=? WHERE id=?",(db.utcnow(),file_id))
    return vid
