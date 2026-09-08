import io
import json
import threading
import time
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from sitrain_app import db, pipeline, storage
from sitrain_app import main
from sitrain_app.main import app
from sitrain_app.terminology import residual_source_issues


@pytest.fixture()
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(db, "DB_PATH", tmp_path / "test.db")
    monkeypatch.setattr(storage, "STORAGE_ROOT", tmp_path / "storage")
    inline_translation = lambda task_id, provider_id=None: pipeline.run_translation(task_id, provider_id)
    monkeypatch.setattr(pipeline, "schedule_translation", inline_translation)
    monkeypatch.setattr(main, "schedule_translation", inline_translation)
    with TestClient(app) as c:
        db.execute("UPDATE providers SET protocol='mock',is_default=1,enabled=1 WHERE id='qwen-mt'")
        db.execute("UPDATE providers SET is_default=0 WHERE id<>'qwen-mt'")
        yield c


def test_health_and_dual_entry_data_model(client):
    health = client.get("/api/health").json()
    assert health["ok"] is True
    assert health["version"] == "0.1.0"
    assert health["llm"]["status"] == "neutral"
    project = client.post("/api/projects", json={"name":"TIA Portal 教材","description":"回归"}).json()
    chapter = client.post(f"/api/projects/{project['id']}/chapters", json={"name":"用户界面"}).json()
    standalone = client.post("/api/files/upload", files={"upload":("standalone.txt",b"Hello operator\n","text/plain")})
    bundled = client.post("/api/files/upload", data={"project_id":project["id"],"chapter_id":chapter["id"]},
                          files={"upload":("chapter.txt",b"Commissioning device\n","text/plain")})
    assert standalone.status_code == bundled.status_code == 200
    dash = client.get("/api/dashboard").json()
    modes = {f["mode"] for f in dash["files"]}
    assert modes == {"standalone","chapter"}


def test_package_import_creates_chapters_and_pairs_by_filename(client, tmp_path):
    from docx import Document
    from pptx import Presentation

    word = tmp_path / "SCL1_01E_Introduction.docx"
    document = Document(); document.add_paragraph("Introduction"); document.save(word)
    cover = tmp_path / "SCL1_00_Cover.docx"
    document = Document(); document.add_paragraph("Cover"); document.save(cover)
    slides = tmp_path / "SCL1_01E_Introduction.pptx"
    Presentation().save(slides)
    project = client.post("/api/projects", json={"name": "整套教材", "description": ""}).json()
    with word.open("rb") as word_stream, cover.open("rb") as cover_stream, slides.open("rb") as slide_stream:
        response = client.post(f"/api/projects/{project['id']}/package-import",
                               files=[("word_files", (word.name, word_stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")),
                                      ("word_files", (cover.name, cover_stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")),
                                      ("ppt_files", (slides.name, slide_stream, "application/vnd.openxmlformats-officedocument.presentationml.presentation"))])
    assert response.status_code == 200, response.text
    summary = response.json()
    assert summary["paired_chapters"] == 1
    assert summary["word_only"] == 1
    detail = client.get(f"/api/projects/{project['id']}").json()
    chapters = {chapter["name"]: chapter["files"] for chapter in detail["chapters"]}
    assert len(chapters) == 2
    assert sorted(file["kind"] for files in chapters.values() for file in files) == ["powerpoint", "word", "word"]


def test_txt_upload_translate_review_publish(client):
    uploaded = client.post("/api/files/upload", files={"upload":("manual.txt",b"Start the motor 24 V\nStop the motor\n","text/plain")}).json()
    fid = uploaded["id"]
    detail = client.get(f"/api/files/{fid}").json()
    assert detail["status"] == "review"
    assert detail["preflight"]["text_units"] == 2
    task = client.post(f"/api/files/{fid}/translate").json()
    task_state = db.row("SELECT status FROM tasks WHERE id=?", (task["id"],))
    assert task_state["status"] == "completed"
    assert task["id"] not in {t["id"] for t in client.get("/api/dashboard").json()["tasks"]}
    units = client.get(f"/api/files/{fid}/units").json()["units"]
    assert units[0]["machine_text"].startswith("[机翻]")
    saved = client.patch(f"/api/units/{units[0]['id']}", json={"human_text":"启动电机 24 V","status":"confirmed","revision":units[0]["revision"]})
    assert saved.status_code == 200
    # Confirm remaining unit so source-residual warnings do not block; only blockers block publishing.
    for unit in client.get(f"/api/files/{fid}/units").json()["units"][1:]:
        client.patch(f"/api/units/{unit['id']}", json={"human_text":"停止电机","status":"confirmed","revision":unit["revision"]})
    published = client.post(f"/api/files/{fid}/publish")
    assert published.status_code == 200
    download = client.get(f"/api/versions/{published.json()['version_id']}/download")
    assert "启动电机 24 V" in download.content.decode("utf-8")


def test_pptx_roundtrip_uses_linguaharu_kernel(client, tmp_path):
    from pptx import Presentation
    from pptx.util import Inches, Pt
    path = tmp_path / "sample.pptx"
    prs = Presentation(); slide = prs.slides.add_slide(prs.slide_layouts[6])
    textbox = slide.shapes.add_textbox(Inches(1),Inches(1),Inches(5),Inches(1))
    textbox.text = "Operator control panel"
    textbox.text_frame.paragraphs[0].runs[0].font.size = Pt(24)
    prs.save(path)
    with path.open("rb") as stream:
        fid = client.post("/api/files/upload", files={"upload":(path.name,stream,"application/vnd.openxmlformats-officedocument.presentationml.presentation")}).json()["id"]
    assert client.get(f"/api/files/{fid}").json()["status"] == "review"
    client.post(f"/api/files/{fid}/translate")
    units = client.get(f"/api/files/{fid}/units").json()["units"]
    assert any("Operator control panel" in u["source_text"] for u in units)
    for unit in units:
        client.patch(f"/api/units/{unit['id']}",json={"human_text":"操作员控制面板","status":"confirmed","revision":unit["revision"]})
    response = client.post(f"/api/files/{fid}/publish")
    assert response.status_code == 200, response.text
    out = tmp_path / "out.pptx"
    out.write_bytes(client.get(f"/api/versions/{response.json()['version_id']}/download").content)
    output_prs = Presentation(out)
    output_shape = output_prs.slides[0].shapes[0]
    assert "操作员控制面板" in output_shape.text
    assert (output_shape.left, output_shape.top, output_shape.width, output_shape.height) == (
        textbox.left, textbox.top, textbox.width, textbox.height,
    )
    assert output_shape.text_frame.paragraphs[0].runs[0].font.size == Pt(24)
    preview = client.get(f"/api/files/{fid}/preview/1")
    assert preview.status_code == 200 and preview.headers["content-type"].startswith("image/png")


def test_server_rejects_disallowed_extension(client):
    response = client.post("/api/files/upload", files={"upload":("picture.png",b"not an image","image/png")})
    assert response.status_code == 415


def test_upload_language_options_default_to_auto_and_limit_targets(client):
    uploaded = client.post("/api/files/upload", files={"upload":("manual.txt",b"Hallo operator\n","text/plain")}).json()
    detail = client.get(f"/api/files/{uploaded['id']}").json()
    assert detail["source_language"] == "auto"
    assert detail["target_language"] == "Chinese"
    invalid = client.post("/api/files/upload", data={"source_language":"German","target_language":"French"},
                          files={"upload":("manual.txt",b"Hallo operator\n","text/plain")})
    assert invalid.status_code == 422


def test_office_embedded_image_is_preserved_by_default(client, tmp_path):
    from PIL import Image, ImageDraw, ImageFont
    from pptx import Presentation
    from pptx.util import Inches
    image_path = tmp_path / "screen.png"
    image = Image.new("RGB", (640, 130), "white")
    ImageDraw.Draw(image).text((20, 30), "START MOTOR 24 V", fill="black", font=ImageFont.truetype("arial.ttf", 42))
    image.save(image_path)
    ppt_path = tmp_path / "image_ocr.pptx"
    prs = Presentation(); slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_picture(str(image_path), Inches(1), Inches(1), width=Inches(7))
    prs.save(ppt_path)
    with ppt_path.open("rb") as stream:
        fid = client.post("/api/files/upload", files={"upload":(ppt_path.name,stream,"application/vnd.openxmlformats-officedocument.presentationml.presentation")}).json()["id"]
    detail = client.get(f"/api/files/{fid}").json()
    assert detail["preflight"]["image_text_translation"] == "disabled"
    units = client.get(f"/api/files/{fid}/units").json()["units"]
    assert not [u for u in units if u["unit_type"] == "office_image_ocr"]
    published = client.post(f"/api/files/{fid}/publish")
    assert published.status_code == 200, published.text
    import zipfile
    output = tmp_path / "image_preserved_output.pptx"
    output.write_bytes(client.get(f"/api/versions/{published.json()['version_id']}/download").content)
    with zipfile.ZipFile(ppt_path) as before, zipfile.ZipFile(output) as after:
        assert before.read("ppt/media/image1.png") == after.read("ppt/media/image1.png")


def test_api_never_returns_provider_secret(client):
    db.execute("UPDATE providers SET api_key='super-secret' WHERE id='qwen-mt'")
    body = client.get("/api/providers").text
    assert "super-secret" not in body
    assert "api_key_configured" in body


def test_single_active_provider_and_custom_provider_creation(client):
    before = client.get("/api/providers").json()["providers"]
    assert [p["id"] for p in before[:3]] == ["qwen-mt", "deepseek", "kimi"]
    assert next(p for p in before if p["id"] == "qwen-mt")["config"]["rpm"] == 60
    deepseek = next(p for p in before if p["id"] == "deepseek")
    response = client.put("/api/providers/deepseek", json={
        "name": deepseek["name"], "protocol": deepseek["protocol"], "base_url": deepseek["base_url"],
        "model": "deepseek-chat", "api_key": "test-key", "enabled": True, "is_default": True,
        "is_fallback": False, "config": deepseek["config"],
    })
    assert response.status_code == 200
    after = client.get("/api/providers").json()["providers"]
    assert [p["id"] for p in after if p["is_active"]] == ["deepseek"]
    created = client.post("/api/providers", json={"name":"Internal API","base_url":"https://api.example.com/v1","model":"translator-x","api_key":"custom-secret"})
    assert created.status_code == 200
    providers = client.get("/api/providers").json()["providers"]
    assert any(p["id"] == created.json()["id"] and p["api_key_configured"] for p in providers)
    assert "custom-secret" not in client.get("/api/providers").text
    removed = client.delete(f"/api/providers/{created.json()['id']}")
    assert removed.status_code == 200
    assert all(p["id"] != created.json()["id"] for p in client.get("/api/providers").json()["providers"])
    assert client.delete("/api/providers/qwen-mt").status_code == 404


def test_provider_connection_status_is_persisted(client, monkeypatch):
    db.execute("UPDATE providers SET api_key='test-key' WHERE id='qwen-mt'")
    monkeypatch.setattr(main, "translate", lambda *_args, **_kwargs: {"connection-check":"连接测试"})
    assert client.post("/api/providers/qwen-mt/test").status_code == 200
    qwen = next(p for p in client.get("/api/providers").json()["providers"] if p["id"] == "qwen-mt")
    assert qwen["last_test_status"] == "success"
    assert qwen["last_tested_at"]


def test_single_digit_chinese_equivalents_do_not_block_release():
    assert not [issue for issue in residual_source_issues("Level 3", "第三级") if issue["code"] == "NUMBER_CHANGED"]
    assert not [issue for issue in residual_source_issues("2-phasig", "两相") if issue["code"] == "NUMBER_CHANGED"]
    assert [issue for issue in residual_source_issues("24 V", "二十四伏") if issue["code"] == "NUMBER_CHANGED"]


def test_qwen_translation_uses_bounded_parallel_requests(monkeypatch):
    from sitrain_app.providers import TranslationRequest, translate_qwen

    active = peak = 0
    lock = threading.Lock()

    def fake_post(_config, payload):
        nonlocal active, peak
        with lock:
            active += 1
            peak = max(peak, active)
        time.sleep(0.03)
        with lock:
            active -= 1
        return {"choices": [{"message": {"content": payload["messages"][0]["content"] + "-zh"}}]}

    monkeypatch.setattr("sitrain_app.providers._post", fake_post)
    result = translate_qwen(
        {"name": "Qwen-MT", "model": "qwen-mt-plus", "config": {"concurrency": 4}},
        TranslationRequest([(str(i), f"unit-{i}") for i in range(12)], "English", "Chinese", [], []),
        lambda _done, _total: None,
    )
    assert result == {str(i): f"unit-{i}-zh" for i in range(12)}
    assert 2 <= peak <= 4


def test_bilingual_glossary_excel_and_reverse_lookup(client):
    created = client.post("/api/glossary", json={
        "source": "TAG", "target": "变量", "aliases": ["tag"], "chinese_aliases": ["变量标签"],
    })
    assert created.status_code == 200
    assert created.json()["chinese_aliases"] == ["变量标签"]

    from openpyxl import load_workbook
    exported = client.get("/api/glossary/export")
    assert exported.headers["cache-control"] == "no-store, no-cache, must-revalidate, max-age=0"
    assert "SITRAIN-global-glossary-v2.xlsx" in exported.headers["content-disposition"]
    workbook = load_workbook(io.BytesIO(exported.content), read_only=True)
    assert tuple(next(workbook["术语表"].iter_rows(min_row=1, max_row=1, values_only=True))) == (
        "英文术语", "中文标准译法", "英文别名（逗号分隔）", "中文别名（逗号分隔，可选）",
    )
    workbook.close()

    from sitrain_app.pipeline import _terms
    _records, engine = _terms(None, "Chinese", "English")
    protected, mapping = engine.protect("变量标签")
    assert engine.restore(protected, mapping) == "TAG"


def test_global_glossary_can_be_edited_and_deleted(client):
    created = client.post("/api/glossary", json={
        "source": "HMI", "target": "人机界面", "aliases": ["hmi"], "chinese_aliases": ["操作屏"],
    })
    assert created.status_code == 200
    term_id = created.json()["id"]
    updated = client.put(f"/api/glossary/{term_id}", json={
        "source": "HMI", "target": "人机接口", "aliases": ["hmi", "Human Machine Interface"],
        "chinese_aliases": ["操作屏"], "mode": "required",
    })
    assert updated.status_code == 200
    assert updated.json()["target"] == "人机接口"
    assert updated.json()["aliases"] == ["hmi", "Human Machine Interface"]
    assert client.delete(f"/api/glossary/{term_id}").status_code == 200
    assert all(item["id"] != term_id for item in client.get("/api/glossary").json()["terms"])
