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
from html import escape
from pathlib import Path
from urllib.parse import urlparse

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
ROOT = Path(__file__).resolve().parents[1]


def _teacher_manual() -> tuple[str, list[tuple[Path, str]]]:
    """Render the checked-in teacher guide as an offline, pictured HTML page."""
    from markdown_it import MarkdownIt

    source = ROOT / "docs" / "teacher_user_manual.md"
    markdown = source.read_text(encoding="utf-8")
    pictures = []
    for relative in set(re.findall(r"!\[[^\]]*\]\((images/[^)]+)\)", markdown)):
        path = (source.parent / relative).resolve()
        if not path.is_file() or source.parent.resolve() not in path.parents:
            raise ValueError(f"手册图片不存在或路径无效：{relative}")
        pictures.append((path, relative))
    body = MarkdownIt("commonmark", {"html": False}).render(markdown)
    style = """body{font:16px/1.75 'Microsoft YaHei','Segoe UI',sans-serif;color:#173033;background:#f2f6f6;margin:0}
main{max-width:940px;margin:32px auto;padding:44px 58px;background:#fff;box-shadow:0 8px 35px #1234}
h1,h2,h3{line-height:1.35;color:#073e43}h1{font-size:30px}h2{margin-top:42px;border-bottom:2px solid #007c83;padding-bottom:8px}
img{max-width:100%;height:auto;display:block;margin:18px auto 6px;border:1px solid #d9e3e4}
code{background:#edf5f5;padding:2px 5px;overflow-wrap:anywhere}li{margin:8px 0}a{color:#007c83}
blockquote{border-left:4px solid #007c83;padding:4px 18px;background:#edf7f7;margin:18px 0}
@media(max-width:700px){main{margin:0;padding:18px}body{font-size:15px}}
@page{size:A4;margin:15mm 13mm 17mm;@bottom-center{content:counter(page);font:9pt 'Segoe UI';color:#66777a}}
@media print{body{background:#fff}main{margin:0;padding:0;box-shadow:none}img{break-inside:avoid}h2{break-after:avoid}}"""
    html = f'<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{escape("SITRAIN 教师绿色版使用手册")}</title><style>{style}</style></head><body><main>{body}</main></body></html>'
    return html, sorted(pictures, key=lambda item: item[1])


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
    from sitrain_app import __build__, __release__, __version__

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
    manual_html, manual_pictures = _teacher_manual()
    manual_pdf = ROOT / "docs" / "teacher_user_manual.pdf"
    if not manual_pdf.is_file() or not manual_pdf.read_bytes().startswith(b"%PDF-"):
        raise ValueError("请先生成 docs/teacher_user_manual.pdf")
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
                archive.writestr("SITRAIN-Teacher/使用手册.html", manual_html)
                archive.write(manual_pdf, "SITRAIN-Teacher/使用手册.pdf")
                archive.writestr("SITRAIN-Teacher/version.json", json.dumps({"product": "SITRAIN-Teacher", "version": __version__, "build": __build__, "release": __release__}, ensure_ascii=False, indent=2))
                release_notes = ROOT / "docs" / "releases" / f"v{__release__}.txt"
                if not release_notes.is_file():
                    release_notes = ROOT / "docs" / "releases" / f"v{__version__}.txt"
                if release_notes.is_file():
                    archive.write(release_notes, "SITRAIN-Teacher/版本说明.txt")
                for picture, relative in manual_pictures:
                    archive.write(picture, "SITRAIN-Teacher/" + relative)
                archive.writestr("SITRAIN-Teacher/使用说明.txt",
                                 f"SITRAIN 教材翻译平台 v{__release__}\n"
                                 "解压整个 SITRAIN-Teacher 文件夹，双击 SITRAIN-Teacher.exe。\n"
                                 "首次使用请打开同目录的 使用手册.pdf，按图文步骤从启动走到交付。\n"
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
