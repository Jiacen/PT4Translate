import json
import sqlite3
import zipfile

from fastapi.testclient import TestClient

from sitrain_app import db
from sitrain_app.relay_server import app as relay_app
from scripts.issue_teacher_packages import issue


def test_teacher_package_contains_clean_seed_only(tmp_path):
    runtime = tmp_path / "runtime"
    runtime.mkdir()
    (runtime / "SITRAIN-Teacher.exe").write_bytes(b"test executable")
    (runtime / "_internal").mkdir()
    (runtime / "_internal" / "app.dat").write_bytes(b"runtime")
    (runtime / "local-testing-material.docx").write_bytes(b"private")
    (runtime / "config").mkdir()
    (runtime / "config" / "system_config.json").write_text("local", encoding="utf-8")
    source = tmp_path / "source.db"
    with sqlite3.connect(source) as conn:
        conn.executescript(db.SCHEMA)
        conn.execute("INSERT INTO providers(id,name,protocol,base_url,api_key,model,enabled,is_default,updated_at) VALUES(?,?,?,?,?,?,?,?,?)",
                     ("qwen-mt", "Qwen-MT", "qwen_mt", "https://example.test/v1", "secret-test-key", "qwen-mt-plus", 1, 1, db.utcnow()))
        conn.execute("INSERT INTO projects(id,name,description,created_at,updated_at) VALUES(?,?,?,?,?)",
                     ("private", "用户测试教材", "", db.utcnow(), db.utcnow()))
    output = tmp_path / "issued"
    issue(runtime, output, ["Shared"], source_db=source)
    packages = sorted(output.glob("*.zip"))
    assert len(packages) == 1
    assert not list(output.glob("*SERVER-ONLY*"))
    for package in packages:
        with zipfile.ZipFile(package) as archive:
            names = archive.namelist()
            assert "SITRAIN-Teacher/sitrain_data/sitrain.db" in names
            assert "SITRAIN-Teacher/使用手册.html" in names
            assert "SITRAIN-Teacher/images/01-welcome.png" in names
            assert "教师绿色版 Pilot 使用手册" in archive.read("SITRAIN-Teacher/使用手册.html").decode("utf-8")
            assert not any("teacher-config.json" in name for name in names)
            assert not any("local-testing-material.docx" in name or "system_config.json" in name for name in names)
            seeded = tmp_path / (package.stem + ".db")
            seeded.write_bytes(archive.read("SITRAIN-Teacher/sitrain_data/sitrain.db"))
        with sqlite3.connect(seeded) as conn:
            assert conn.execute("SELECT COUNT(*) FROM projects").fetchone()[0] == 0
            assert conn.execute("SELECT api_key FROM providers WHERE is_default=1").fetchone()[0] == "secret-test-key"


def test_relay_requires_teacher_token_and_forwards_only_translation(tmp_path, monkeypatch):
    tokens = tmp_path / "tokens.json"
    tokens.write_text(json.dumps({"teacher01": "t" * 40}), encoding="utf-8")
    monkeypatch.setenv("SITRAIN_RELAY_TOKENS_FILE", str(tokens))
    monkeypatch.setenv("SITRAIN_RELAY_UPSTREAM_KEY", "upstream-test-key")
    monkeypatch.setenv("SITRAIN_RELAY_RPM", "600")
    monkeypatch.setattr("sitrain_app.relay_server._wait_for_slot", lambda: None)
    calls = []

    class Response:
        status_code = 200

        def json(self):
            return {"choices": [{"message": {"content": "译文"}}]}

    def fake_post(url, headers, json, timeout):
        calls.append((url, headers, json))
        return Response()

    monkeypatch.setattr("sitrain_app.relay_server.requests.post", fake_post)
    payload = {"model": "untrusted-model", "messages": [{"role": "user", "content": "Hello"}],
               "translation_options": {"source_lang": "English", "target_lang": "Chinese"}}
    with TestClient(relay_app) as client:
        assert client.post("/v1/chat/completions", json=payload).status_code == 401
        answer = client.post("/v1/chat/completions", json=payload,
                             headers={"Authorization": "Bearer " + "t" * 40})
        assert answer.status_code == 200
        assert answer.json()["choices"][0]["message"]["content"] == "译文"
    assert calls[0][1]["Authorization"] == "Bearer upstream-test-key"
    assert calls[0][2]["model"] == "qwen-mt-plus"
