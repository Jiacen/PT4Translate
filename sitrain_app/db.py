from __future__ import annotations

import json
import sqlite3
import threading
import uuid
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

from .config import DB_PATH, ensure_data_dirs

_lock = threading.RLock()


def utcnow() -> str:
    return datetime.now(timezone.utc).isoformat()


def new_id() -> str:
    return uuid.uuid4().hex


@contextmanager
def connect():
    ensure_data_dirs()
    conn = sqlite3.connect(DB_PATH, timeout=30, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA journal_mode=WAL")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


SCHEMA = """
CREATE TABLE IF NOT EXISTS projects (
  id TEXT PRIMARY KEY, name TEXT NOT NULL, description TEXT NOT NULL DEFAULT '',
  owner_id TEXT NOT NULL DEFAULT 'local-admin', created_at TEXT NOT NULL, updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS project_members (
  project_id TEXT NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
  user_id TEXT NOT NULL, role TEXT NOT NULL DEFAULT 'editor', PRIMARY KEY(project_id,user_id)
);
CREATE TABLE IF NOT EXISTS chapters (
  id TEXT PRIMARY KEY, project_id TEXT NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
  name TEXT NOT NULL, position INTEGER NOT NULL DEFAULT 0, created_at TEXT NOT NULL, updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS files (
  id TEXT PRIMARY KEY, project_id TEXT REFERENCES projects(id) ON DELETE SET NULL,
  chapter_id TEXT REFERENCES chapters(id) ON DELETE SET NULL, owner_id TEXT NOT NULL DEFAULT 'local-admin',
  original_name TEXT NOT NULL, extension TEXT NOT NULL, kind TEXT NOT NULL,
  mode TEXT NOT NULL CHECK(mode IN ('standalone','chapter')), status TEXT NOT NULL,
  source_language TEXT NOT NULL DEFAULT 'auto', target_language TEXT NOT NULL DEFAULT 'Chinese',
  preflight_json TEXT NOT NULL DEFAULT '{}', error TEXT, created_at TEXT NOT NULL, updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS file_versions (
  id TEXT PRIMARY KEY, file_id TEXT NOT NULL REFERENCES files(id) ON DELETE CASCADE,
  version_type TEXT NOT NULL, sequence INTEGER NOT NULL, path TEXT NOT NULL,
  checksum TEXT NOT NULL, size INTEGER NOT NULL, metadata_json TEXT NOT NULL DEFAULT '{}', created_at TEXT NOT NULL,
  UNIQUE(file_id, version_type, sequence)
);
CREATE TABLE IF NOT EXISTS translation_units (
  id TEXT PRIMARY KEY, file_id TEXT NOT NULL REFERENCES files(id) ON DELETE CASCADE,
  stable_key TEXT NOT NULL, page INTEGER, object_ref TEXT, unit_type TEXT NOT NULL DEFAULT 'text',
  source_text TEXT NOT NULL, machine_text TEXT, human_text TEXT, status TEXT NOT NULL DEFAULT 'pending_translation',
  qa_json TEXT NOT NULL DEFAULT '[]', confidence REAL, source_payload_json TEXT NOT NULL DEFAULT '{}',
  revision INTEGER NOT NULL DEFAULT 1, updated_at TEXT NOT NULL, UNIQUE(file_id, stable_key)
);
CREATE TABLE IF NOT EXISTS tasks (
  id TEXT PRIMARY KEY, file_id TEXT REFERENCES files(id) ON DELETE CASCADE,
  task_type TEXT NOT NULL, status TEXT NOT NULL, progress REAL NOT NULL DEFAULT 0,
  idempotency_key TEXT NOT NULL UNIQUE, detail TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL, updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS providers (
  id TEXT PRIMARY KEY, name TEXT NOT NULL, protocol TEXT NOT NULL, base_url TEXT NOT NULL DEFAULT '',
  api_key TEXT NOT NULL DEFAULT '', model TEXT NOT NULL DEFAULT '', enabled INTEGER NOT NULL DEFAULT 1,
  is_default INTEGER NOT NULL DEFAULT 0, is_fallback INTEGER NOT NULL DEFAULT 0,
  config_json TEXT NOT NULL DEFAULT '{}', last_test_status TEXT NOT NULL DEFAULT 'untested',
  last_tested_at TEXT, last_test_message TEXT NOT NULL DEFAULT '', updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS glossary_terms (
  id TEXT PRIMARY KEY, project_id TEXT REFERENCES projects(id) ON DELETE CASCADE,
  source TEXT NOT NULL, target TEXT NOT NULL, aliases_json TEXT NOT NULL DEFAULT '[]',
  chinese_aliases_json TEXT NOT NULL DEFAULT '[]',
  mode TEXT NOT NULL DEFAULT 'required', case_sensitive INTEGER NOT NULL DEFAULT 0, updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS translation_memory (
  id TEXT PRIMARY KEY, project_id TEXT REFERENCES projects(id) ON DELETE CASCADE,
  source_norm TEXT NOT NULL, source TEXT NOT NULL, target TEXT NOT NULL,
  origin_unit_id TEXT, approved_at TEXT NOT NULL, UNIQUE(project_id, source_norm)
);
CREATE TABLE IF NOT EXISTS review_changes (
  id TEXT PRIMARY KEY, file_id TEXT NOT NULL REFERENCES files(id) ON DELETE CASCADE,
  page INTEGER NOT NULL, object_id TEXT NOT NULL, unit_id TEXT REFERENCES translation_units(id) ON DELETE SET NULL,
  source_text TEXT NOT NULL, before_text TEXT NOT NULL, after_text TEXT NOT NULL,
  change_type TEXT NOT NULL, created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS ppt_word_mappings (
  id TEXT PRIMARY KEY, chapter_id TEXT NOT NULL REFERENCES chapters(id) ON DELETE CASCADE,
  ppt_file_id TEXT NOT NULL REFERENCES files(id), word_file_id TEXT NOT NULL REFERENCES files(id),
  slide_number INTEGER NOT NULL, media_part TEXT NOT NULL, confidence REAL NOT NULL,
  confirmed INTEGER NOT NULL DEFAULT 0, candidates_json TEXT NOT NULL DEFAULT '[]',
  decision TEXT NOT NULL DEFAULT 'pending', updated_at TEXT NOT NULL,
  UNIQUE(word_file_id, media_part)
);
CREATE INDEX IF NOT EXISTS idx_files_chapter ON files(chapter_id);
CREATE INDEX IF NOT EXISTS idx_units_file_status ON translation_units(file_id,status);
CREATE INDEX IF NOT EXISTS idx_versions_file ON file_versions(file_id,version_type,sequence);
CREATE INDEX IF NOT EXISTS idx_review_changes_file ON review_changes(file_id,created_at);
"""


def init_db() -> None:
    with _lock, connect() as conn:
        conn.executescript(SCHEMA)
        provider_columns = {r[1] for r in conn.execute("PRAGMA table_info(providers)").fetchall()}
        if "last_test_status" not in provider_columns:
            conn.execute("ALTER TABLE providers ADD COLUMN last_test_status TEXT NOT NULL DEFAULT 'untested'")
        if "last_tested_at" not in provider_columns:
            conn.execute("ALTER TABLE providers ADD COLUMN last_tested_at TEXT")
        if "last_test_message" not in provider_columns:
            conn.execute("ALTER TABLE providers ADD COLUMN last_test_message TEXT NOT NULL DEFAULT ''")
        glossary_columns = {r[1] for r in conn.execute("PRAGMA table_info(glossary_terms)").fetchall()}
        if "chinese_aliases_json" not in glossary_columns:
            conn.execute("ALTER TABLE glossary_terms ADD COLUMN chinese_aliases_json TEXT NOT NULL DEFAULT '[]'")
        mapping_columns = {r[1] for r in conn.execute("PRAGMA table_info(ppt_word_mappings)").fetchall()}
        if "candidates_json" not in mapping_columns:
            conn.execute("ALTER TABLE ppt_word_mappings ADD COLUMN candidates_json TEXT NOT NULL DEFAULT '[]'")
        if "decision" not in mapping_columns:
            conn.execute("ALTER TABLE ppt_word_mappings ADD COLUMN decision TEXT NOT NULL DEFAULT 'pending'")
        if conn.execute("SELECT COUNT(*) FROM providers").fetchone()[0] == 0:
            now = utcnow()
            presets = [
                ("qwen-mt", "Qwen-MT", "qwen_mt", "https://dashscope.aliyuncs.com/compatible-mode/v1", "qwen-mt-plus", 1, 0,
                 {"source_lang":"English","target_lang":"Chinese","temperature":0.1,"batch_size":1,"concurrency":4,"rpm":60,"timeout":120,"retries":3,"backoff":2,"structured_output":False,"headers":{},"prompt":""}),
                ("deepseek", "DeepSeek", "openai_chat", "https://api.deepseek.com", "deepseek-v4-flash", 0, 1,
                 {"temperature":0.1,"batch_size":20,"concurrency":4,"rpm":0,"timeout":120,"retries":3,"backoff":2,"structured_output":True,"headers":{},"prompt":"Translate each item faithfully into {target_lang}. Return only a JSON object using the same keys."}),
                ("kimi", "Kimi", "openai_chat", "https://api.moonshot.cn/v1", "kimi-k2.6", 0, 0,
                 {"temperature":0.1,"batch_size":20,"concurrency":4,"rpm":0,"timeout":120,"retries":3,"backoff":2,"structured_output":True,"headers":{},"prompt":"Translate each item faithfully into {target_lang}. Return only a JSON object using the same keys."}),
            ]
            conn.executemany("INSERT INTO providers(id,name,protocol,base_url,model,is_default,is_fallback,config_json,updated_at) VALUES(?,?,?,?,?,?,?,?,?)",
                             [(a,b,c,d,e,f,g,json.dumps(h,ensure_ascii=False),now) for a,b,c,d,e,f,g,h in presets])
        kimi = conn.execute("SELECT id FROM providers WHERE id='kimi'").fetchone()
        ernie = conn.execute("SELECT id FROM providers WHERE id='ernie-qianfan'").fetchone()
        legacy = conn.execute("SELECT api_key FROM providers WHERE id='openai-compatible'").fetchone()
        if not kimi and ernie:
            conn.execute("UPDATE providers SET id='kimi',name='Kimi',base_url='https://api.moonshot.cn/v1',model='kimi-k2.6',updated_at=? WHERE id='ernie-qianfan'", (utcnow(),))
        elif not kimi and legacy and not legacy[0]:
            conn.execute("UPDATE providers SET id='kimi',name='Kimi',base_url='https://api.moonshot.cn/v1',model='kimi-k2.6',updated_at=? WHERE id='openai-compatible'", (utcnow(),))
        conn.execute("UPDATE providers SET model='deepseek-v4-flash', updated_at=? WHERE id='deepseek' AND model=''", (utcnow(),))
        qwen = conn.execute("SELECT config_json FROM providers WHERE id='qwen-mt'").fetchone()
        if qwen:
            try:
                qwen_config = json.loads(qwen[0] or "{}")
            except json.JSONDecodeError:
                qwen_config = {}
            if not qwen_config.get("rpm"):
                qwen_config["rpm"] = 60
                conn.execute("UPDATE providers SET config_json=?, updated_at=? WHERE id='qwen-mt'",
                             (json.dumps(qwen_config, ensure_ascii=False), utcnow()))
        conn.execute("UPDATE tasks SET status='queued', detail='服务重启后恢复' WHERE status='running'")


def rows(sql: str, params: Iterable[Any] = ()) -> list[dict[str, Any]]:
    with connect() as conn:
        result = [dict(r) for r in conn.execute(sql, tuple(params)).fetchall()]
    for item in result:
        for key in tuple(item):
            if key.endswith("_json"):
                try:
                    item[key[:-5]] = json.loads(item[key] or "null")
                except json.JSONDecodeError:
                    item[key[:-5]] = None
    return result


def row(sql: str, params: Iterable[Any] = ()) -> dict[str, Any] | None:
    result = rows(sql, params)
    return result[0] if result else None


def execute(sql: str, params: Iterable[Any] = ()) -> None:
    with _lock, connect() as conn:
        conn.execute(sql, tuple(params))


def scalar(sql: str, params: Iterable[Any] = ()) -> Any:
    with connect() as conn:
        return conn.execute(sql, tuple(params)).fetchone()[0]
