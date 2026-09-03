from __future__ import annotations

import json
import shutil
import traceback
from pathlib import Path

from . import db
from .documents import TRANSLATE_OFFICE_IMAGE_TEXT, convert_legacy, extract_units, preflight, prepare_preview_cache, rebuild
from .providers import TranslationRequest, translate
from .storage import checksum, file_dir
from .terminology import Term, TerminologyEngine, normalize, residual_source_issues


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
        start_automatic_translation(file_id)
    except Exception as exc:
        db.execute("UPDATE files SET status='failed',error=?,updated_at=? WHERE id=?",(str(exc),db.utcnow(),file_id))
        raise


def start_automatic_translation(file_id: str) -> None:
    revision = db.scalar("SELECT COALESCE(MAX(revision),0) FROM translation_units WHERE file_id=?", (file_id,))
    key = f"translate:{file_id}:{revision}:default"
    if db.row("SELECT id FROM tasks WHERE idempotency_key=?", (key,)):
        return
    tid, now = db.new_id(), db.utcnow()
    db.execute("INSERT INTO tasks(id,file_id,task_type,status,idempotency_key,created_at,updated_at) VALUES(?,?,?,?,?,?,?)",
               (tid, file_id, "translate", "queued", key, now, now))
    run_translation(tid)


def refresh_machine_version(file_id: str) -> str | None:
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
        prepare_preview_cache(output, file_dir(file_id) / "previews", int((file.get("preflight") or {}).get("pages") or 0))
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


def run_translation(task_id: str, provider_id: str | None = None) -> None:
    task = db.row("SELECT * FROM tasks WHERE id=?",(task_id,))
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
            if normalize(unit["source_text"]) in tm:
                direct[unit["stable_key"]] = tm[normalize(unit["source_text"])]
                continue
            text,mapping = engine.protect(unit["source_text"])
            pending.append((unit["stable_key"],text))
            protected_map[unit["stable_key"]] = mapping
        def progress(done,total):
            db.execute("UPDATE tasks SET progress=?,detail=?,updated_at=? WHERE id=?",(0.05+0.8*(done/max(1,total)),f"已翻译 {done}/{total}",db.utcnow(),task_id))
        result = direct | (translate(provider,TranslationRequest(pending,file["source_language"],file["target_language"],term_records,memory),progress) if pending else {})
        with db.connect() as conn:
            for unit in units:
                translated = engine.restore(result[unit["stable_key"]],protected_map.get(unit["stable_key"],{}))
                issues = engine.validate(unit["source_text"],translated)+residual_source_issues(unit["source_text"],translated)
                status = "terminology_issue" if any(x["code"]=="TERM_MISSING" for x in issues) else "pending_review"
                conn.execute("UPDATE translation_units SET machine_text=?,status=?,qa_json=?,revision=revision+1,updated_at=? WHERE id=?",
                             (translated,status,json.dumps(issues,ensure_ascii=False),db.utcnow(),unit["id"]))
            conn.execute("UPDATE files SET status='review',error=NULL,updated_at=? WHERE id=?",(db.utcnow(),file["id"]))
            conn.execute("UPDATE tasks SET status='completed',progress=1,detail='翻译完成，等待校对',updated_at=? WHERE id=?",(db.utcnow(),task_id))
        try:
            refresh_machine_version(file["id"])
        except Exception as exc:
            db.execute("UPDATE files SET error=?,updated_at=? WHERE id=?", (f"校对预览生成失败：{exc}", db.utcnow(), file["id"]))
    except Exception as exc:
        db.execute("UPDATE tasks SET status='failed',detail=?,updated_at=? WHERE id=?",(str(exc),db.utcnow(),task_id))
        db.execute("UPDATE files SET status='failed',error=?,updated_at=? WHERE id=?",(str(exc),db.utcnow(),file["id"]))


def publish(file_id: str) -> str:
    file = db.row("SELECT * FROM files WHERE id=?",(file_id,))
    source_v = _version(file_id,"working") or _version(file_id,"original")
    source = Path(source_v["path"])
    units = db.rows("SELECT * FROM translation_units WHERE file_id=? ORDER BY CAST(stable_key AS INTEGER),stable_key",(file_id,))
    blockers = [q for u in units for q in (u.get("qa") or []) if q.get("severity")=="blocker" and u["status"]!="confirmed"]
    if blockers:
        raise RuntimeError(f"仍有 {len(blockers)} 个阻止发布的 QA 问题")
    source_json_path = (file.get("preflight") or {}).get("source_json")
    result = rebuild(source,Path(source_json_path) if source_json_path else None,units,file_dir(file_id)/"published",file["source_language"],file["target_language"])
    vid = add_version(file_id,"published",result,{"qa_blockers":0})
    db.execute("UPDATE files SET status='published',updated_at=? WHERE id=?",(db.utcnow(),file_id))
    return vid
