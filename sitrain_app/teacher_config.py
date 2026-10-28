"""Optional per-teacher relay configuration for the portable Windows build."""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from urllib.parse import urlparse

from . import db


def config_path() -> Path:
    override = os.getenv("SITRAIN_TEACHER_CONFIG")
    if override:
        return Path(override).resolve()
    root = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[1]
    return root / "teacher-config.json"


def load_teacher_config() -> dict | None:
    path = config_path()
    if not path.exists():
        return None
    data = json.loads(path.read_text(encoding="utf-8"))
    url = str(data.get("relay_url") or "").rstrip("/")
    token = str(data.get("access_token") or "")
    parsed = urlparse(url)
    if parsed.scheme != "https" or not parsed.netloc or parsed.username or parsed.password or parsed.query or parsed.fragment:
        raise RuntimeError("teacher-config.json 中的 relay_url 必须是 HTTPS 地址")
    if not url.endswith("/v1") or len(token) < 32:
        raise RuntimeError("teacher-config.json 的翻译入口或访问凭证无效")
    return {"relay_url": url, "access_token": token}


def configure_teacher_provider() -> bool:
    data = load_teacher_config()
    if not data:
        return False
    with db.connect() as conn:
        conn.execute("UPDATE providers SET enabled=0,is_default=0")
        conn.execute("""UPDATE providers SET name='统一翻译服务',protocol='qwen_mt',base_url=?,api_key=?,
                       enabled=1,is_default=1,last_test_status='untested',updated_at=? WHERE id='qwen-mt'""",
                     (data["relay_url"], data["access_token"], db.utcnow()))
    return True
