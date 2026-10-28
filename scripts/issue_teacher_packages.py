"""Issue separate clean portable ZIPs with relay or direct Qwen credentials."""
from __future__ import annotations

import argparse
import json
import os
import re
import secrets
import sqlite3
import sys
import tempfile
import zipfile
from pathlib import Path
from urllib.parse import urlparse

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


def _read_active_provider(source_db: Path) -> dict:
    if not source_db.is_file():
        raise ValueError("现有本地数据库不存在")
    with sqlite3.connect(source_db) as conn:
        conn.row_factory = sqlite3.Row
        row = conn.execute("SELECT name,protocol,base_url,api_key,model,config_json FROM providers WHERE is_default=1 AND enabled=1 LIMIT 1").fetchone()
    if not row or not row["api_key"]:
        raise ValueError("当前翻译服务没有已配置的 API Key")
    if row["protocol"] != "qwen_mt":
        raise ValueError("本次共享 Key 试用包仅支持当前 Qwen-MT 翻译服务")
    return dict(row)


def _fresh_database(provider: dict, path: Path) -> bytes:
    from sitrain_app import db
    previous = db.DB_PATH
    try:
        db.DB_PATH = path
        db.init_db()
        settings = json.loads(provider["config_json"] or "{}")
        # Copy translation limits only; never include arbitrary request headers.
        safe_settings = {key: settings[key] for key in ("source_lang", "target_lang", "temperature", "batch_size", "concurrency", "rpm", "timeout", "retries", "backoff") if key in settings}
        with db.connect() as conn:
            conn.execute("UPDATE providers SET enabled=0,is_default=0")
            conn.execute("""UPDATE providers SET name=?,protocol=?,base_url=?,api_key=?,model=?,config_json=?,
                            enabled=1,is_default=1,last_test_status='untested' WHERE id='qwen-mt'""",
                         (provider["name"], provider["protocol"], provider["base_url"], provider["api_key"],
                          provider["model"], json.dumps(safe_settings),))
        conn = sqlite3.connect(path)
        try:
            conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")
            conn.execute("PRAGMA journal_mode=DELETE")
        finally:
            conn.close()
        return path.read_bytes()
    finally:
        db.DB_PATH = previous


def issue(runtime: Path, destination: Path, teachers: list[str], relay_url: str | None = None,
          source_db: Path | None = None) -> None:
    if bool(relay_url) == bool(source_db):
        raise ValueError("请选择统一入口模式或共享本机 Key 试用模式")
    url = None
    if relay_url:
        parsed = urlparse(relay_url)
        if parsed.scheme != "https" or not parsed.netloc or parsed.query or parsed.fragment:
            raise ValueError("relay_url 必须是老师可访问的 HTTPS 地址")
        url = relay_url.rstrip("/")
        if not url.endswith("/v1"):
            raise ValueError("relay_url 必须以 /v1 结尾")
    if not runtime.is_dir() or not (runtime / "SITRAIN-Teacher.exe").is_file():
        raise ValueError("请先构建 SITRAIN-Teacher.exe")
    if destination.exists() and any(destination.iterdir()):
        raise ValueError("输出目录必须为空，避免覆盖已经发出的凭证")
    names = [name.strip() for name in teachers]
    if not names or len(names) > 10 or len(names) != len(set(names)) or any(not re.fullmatch(r"[A-Za-z0-9_-]{1,40}", name) for name in names):
        raise ValueError("请提供 1–10 个不重复的英文或数字老师代号")
    if source_db and len(names) != 1:
        raise ValueError("共用 API Key 模式只需生成一份通用绿色包")
    if not (runtime / "_internal").is_dir():
        raise ValueError("运行包缺少依赖目录")
    runtime_files = [runtime / "SITRAIN-Teacher.exe"] + [path for path in (runtime / "_internal").rglob("*") if path.is_file()]
    if any(path.is_symlink() or "sitrain_data" in path.parts or path.name == "teacher-config.json" for path in runtime_files):
        raise ValueError("运行包依赖目录混入用户数据或本机配置，已停止发包")
    destination.mkdir(parents=True, exist_ok=True)
    tokens = {name: secrets.token_urlsafe(36) for name in names} if url else {}
    if tokens:
        # This file is for the relay host only. Never place it inside a ZIP.
        (destination / "relay-tokens.SERVER-ONLY.json").write_text(json.dumps(tokens, indent=2), encoding="utf-8")
    provider = _read_active_provider(source_db) if source_db else None
    seed_path = None
    try:
        if provider:
            fd, filename = tempfile.mkstemp(prefix="sitrain-clean-db-", suffix=".db")
            os.close(fd)
            seed_path = Path(filename)
            seed_db = _fresh_database(provider, seed_path)
        else:
            seed_db = None
        for name in names:
            archive_path = destination / f"SITRAIN-Teacher-{name}.zip"
            with zipfile.ZipFile(archive_path, "x", zipfile.ZIP_DEFLATED) as archive:
                for path in runtime_files:
                    archive.write(path, "SITRAIN-Teacher/" + path.relative_to(runtime).as_posix())
                if url:
                    archive.writestr("SITRAIN-Teacher/teacher-config.json", json.dumps({"relay_url": url, "access_token": tokens[name]}, ensure_ascii=False))
                else:
                    archive.writestr("SITRAIN-Teacher/sitrain_data/sitrain.db", seed_db)
                archive.writestr("SITRAIN-Teacher/使用说明.txt",
                                 "解压整个 SITRAIN-Teacher 文件夹，双击 SITRAIN-Teacher.exe。\n"
                                 "浏览器地址：http://127.0.0.1:9000/\n"
                                 "请保留整个文件夹；教材与校对记录保存在其中的 sitrain_data 子目录。\n"
                                 "更新程序前，请先备份整个 sitrain_data 目录。\n")
    finally:
        if seed_path is not None:
            seed_path.unlink(missing_ok=True)
    print(f"已生成 {len(names)} 份绿色包；不含当前测试教材和校对数据。")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--runtime", type=Path, default=Path("dist/SITRAIN-Teacher"))
    parser.add_argument("--output", type=Path, required=True)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--relay-url")
    mode.add_argument("--source-db", type=Path)
    parser.add_argument("teachers", nargs="+")
    args = parser.parse_args()
    issue(args.runtime, args.output, args.teachers, relay_url=args.relay_url, source_db=args.source_db)
