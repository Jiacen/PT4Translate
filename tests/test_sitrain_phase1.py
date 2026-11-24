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
from sitrain_app.terminology import Term, TerminologyEngine, residual_source_issues


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


@pytest.fixture()
def import_pair(monkeypatch):
    from docx import Document
    from pptx import Presentation

    word, ppt = io.BytesIO(), io.BytesIO()
    document = Document()
    document.add_paragraph("Configuration and commissioning")
    document.save(word)
    Presentation().save(ppt)
    prepared = []
    monkeypatch.setattr(main, "prepare_file", prepared.append)
    return [
        ("word_files", ("01_DR-G220_Configuration.docx", word.getvalue())),
        ("ppt_files", ("01_DR-G220_Configuration.pptx", ppt.getvalue())),
    ], prepared


@pytest.mark.parametrize("concurrent", [False, True])
def test_package_import_retry_keeps_one_pair(client, import_pair, monkeypatch, concurrent):
    import asyncio
    from concurrent.futures import ThreadPoolExecutor

    files, prepared = import_pair
    project = client.post("/api/projects", json={"name": "重复导入回归"}).json()
    url = f"/api/projects/{project['id']}/package-import"
    save = main.save_upload

    async def delayed_save(*args):
        # Force two concurrent requests to overlap while writing their files.
        await asyncio.sleep(0.02)
        return await save(*args)

    monkeypatch.setattr(main, "save_upload", delayed_save)
    if concurrent:
        with ThreadPoolExecutor(max_workers=2) as executor:
            futures = [executor.submit(client.post, url, files=files) for _ in range(2)]
            responses = [future.result() for future in futures]
    else:
        responses = [client.post(url, files=files) for _ in range(2)]
    assert [response.status_code for response in responses] == [200, 200]
    detail = client.get(f"/api/projects/{project['id']}").json()
    assert len(detail["chapters"]) == 1
    assert len(detail["chapters"][0]["files"]) == 2
    assert sorted(len(response.json()["files"]) for response in responses) == [0, 2]
    assert sorted(len(response.json()["skipped_files"]) for response in responses) == [0, 2]
    assert len(prepared) == 2
    assert db.scalar("SELECT COUNT(*) FROM file_versions") == 2
    assert len(list(storage.STORAGE_ROOT.glob("*/*/original/*"))) == 2


def test_package_import_retry_preserves_confirmed_structure_and_files(client, import_pair):
    files, prepared = import_pair
    project = client.post("/api/projects", json={"name": "保留已确认结构"}).json()
    url = f"/api/projects/{project['id']}/package-import"
    original = client.post(url, files=files).json()["files"]
    chapter_id = original[0]["chapter"]
    db.execute("UPDATE chapters SET name='教师调整后的章节' WHERE id=?", (chapter_id,))
    db.execute("UPDATE files SET status='published' WHERE project_id=?", (project["id"],))
    db.execute("UPDATE projects SET workflow_status='waiting_start' WHERE id=?", (project["id"],))
    response = client.post(url, files=files)
    assert response.status_code == 200, response.text
    assert {file["id"] for file in response.json()["skipped_files"]} == {file["id"] for file in original}
    detail = client.get(f"/api/projects/{project['id']}").json()
    assert detail["workflow_status"] == "waiting_start"
    assert len(detail["chapters"]) == 1
    assert detail["chapters"][0]["name"] == "教师调整后的章节"
    assert {file["status"] for file in detail["chapters"][0]["files"]} == {"published"}
    assert len(prepared) == 2


@pytest.mark.parametrize("conflict", ["content", "language"])
def test_package_import_same_name_conflict_does_not_append_or_overwrite(client, import_pair, conflict):
    from docx import Document

    files, prepared = import_pair
    project = client.post("/api/projects", json={"name": "同名冲突回归"}).json()
    url = f"/api/projects/{project['id']}/package-import"
    client.post(url, files=files)
    before = db.rows("SELECT * FROM file_versions ORDER BY id")
    db.execute("UPDATE projects SET workflow_status='waiting_start' WHERE id=?", (project["id"],))
    new_files = list(files)
    data = {}
    if conflict == "content":
        changed = io.BytesIO()
        document = Document()
        document.add_paragraph("Changed source material")
        document.save(changed)
        new_files[0] = ("word_files", (files[0][1][0], changed.getvalue()))
    else:
        data = {"target_language": "English"}
    # A conflict must be detected before even unrelated new files are added.
    new_files.insert(0, ("word_files", ("00_Cover.docx", files[0][1][1])))
    response = client.post(url, files=new_files, data=data)
    assert response.status_code == 409, response.text
    assert "同名" in response.json()["detail"]
    assert db.rows("SELECT * FROM file_versions ORDER BY id") == before
    assert db.scalar("SELECT COUNT(*) FROM files") == 2
    assert db.scalar("SELECT COUNT(*) FROM chapters") == 1
    assert db.row("SELECT workflow_status FROM projects WHERE id=?", (project["id"],))["workflow_status"] == "waiting_start"
    assert len(prepared) == 2


def test_package_import_adds_missing_companion_and_isolates_projects(client, import_pair):
    files, prepared = import_pair
    project = client.post("/api/projects", json={"name": "补充配套文件"}).json()
    url = f"/api/projects/{project['id']}/package-import"
    first = client.post(url, files=files[:1]).json()
    second = client.post(url, files=files).json()
    assert len(second["files"]) == len(second["skipped_files"]) == 1
    assert second["files"][0]["chapter"] == first["files"][0]["chapter"]
    other = client.post("/api/projects", json={"name": "另一教材项目"}).json()
    independent = client.post(f"/api/projects/{other['id']}/package-import", files=files).json()
    assert len(independent["files"]) == 2
    assert independent["skipped_files"] == []
    assert len(prepared) == 4


def test_paired_chapter_without_word_images_needs_no_sync(monkeypatch):
    monkeypatch.setattr(db, "row", lambda *_args, **_kwargs: None)
    files = [{"kind": "word", "status": "published", "preflight": {"images": 0}, "id": "word"},
             {"kind": "powerpoint", "status": "published", "id": "ppt"}]
    assert main._chapter_sync_gate("chapter", files) is None
    files[0]["preflight"]["images"] = 1
    assert main._chapter_sync_gate("chapter", files)["code"] == "mapping"


def test_project_requires_structure_confirmation_and_releases_one_chapter(client, monkeypatch):
    project = client.post("/api/projects", json={"name": "章节队列", "description": ""}).json()
    first = client.post(f"/api/projects/{project['id']}/chapters", json={"name": "第一章"}).json()
    second = client.post(f"/api/projects/{project['id']}/chapters", json={"name": "第二章"}).json()
    first_file = client.post("/api/files/upload", data={"project_id": project["id"], "chapter_id": first["id"]},
                             files={"upload": ("one.txt", b"First chapter", "text/plain")}).json()["id"]
    second_file = client.post("/api/files/upload", data={"project_id": project["id"], "chapter_id": second["id"]},
                              files={"upload": ("two.txt", b"Second chapter", "text/plain")}).json()["id"]
    second_companion = client.post("/api/files/upload", data={"project_id": project["id"], "chapter_id": second["id"]},
                                   files={"upload": ("two-notes.txt", b"Second chapter notes", "text/plain")}).json()["id"]
    detail = client.get(f"/api/projects/{project['id']}").json()
    assert detail["teacher_status"]["code"] == "confirm_structure"
    assert client.get(f"/api/files/{first_file}").json()["status"] == "ready"
    assert client.post(f"/api/files/{first_file}/translate").status_code == 409

    reordered = client.put(f"/api/projects/{project['id']}/structure", json={"chapters": [
        {"id": second["id"], "name": "第二章（优先）", "file_ids": [second_file, second_companion]},
        {"id": first["id"], "name": "第一章", "file_ids": [first_file]},
    ]})
    assert reordered.status_code == 200
    assert reordered.json()["teacher_status"]["code"] == "waiting_start"

    submitted = []
    monkeypatch.setattr(pipeline, "schedule_translation", lambda task_id, provider_id=None: submitted.append(task_id))
    started = client.post(f"/api/projects/{project['id']}/translate")
    assert started.status_code == 200
    assert started.json()["teacher_status"]["code"] == "translating"
    assert len(submitted) == 2
    task_files = {db.row("SELECT file_id FROM tasks WHERE id=?", (task_id,))["file_id"] for task_id in submitted}
    assert task_files == {second_file, second_companion}
    assert not db.row("SELECT id FROM tasks WHERE file_id=?", (first_file,))


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


def test_ppt_translation_pipeline_preserves_title_breaks_and_rejects_flattened_memory(client, tmp_path, monkeypatch):
    from pptx import Presentation
    from pptx.oxml.xmlchemy import OxmlElement
    from pptx.util import Inches, Pt
    from sitrain_app.terminology import normalize

    monkeypatch.setattr(pipeline, "prepare_preview_cache", lambda *args: None)
    source = tmp_path / 'title.pptx'
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(2))
    paragraph = box.text_frame.paragraphs[0]
    first = paragraph.add_run(); first.text = 'Step 1'; first.font.size = Pt(28); first.font.bold = True
    paragraph._p.append(OxmlElement('a:br'))
    second = paragraph.add_run(); second.text = 'Control panel'; second.font.size = Pt(18); second.font.bold = False
    deck.save(source)
    db.execute("INSERT INTO translation_memory(id,source_norm,source,target,approved_at) VALUES(?,?,?,?,?)",
               (db.new_id(), normalize('Step 1\nControl panel'), 'Step 1\nControl panel', '步骤1控制面板', db.utcnow()))
    with source.open('rb') as stream:
        response = client.post('/api/files/upload', files={'upload': (source.name, stream)})
    assert response.status_code == 200, response.text
    fid = response.json()['id']
    units = client.get(f'/api/files/{fid}/units').json()['units']
    unit = next(unit for unit in units if unit['source_text'] == 'Step 1\nControl panel')
    assert unit['machine_text'] == '[机翻]Step 1\n[机翻]Control panel'
    version = db.row("SELECT path FROM file_versions WHERE file_id=? AND version_type='machine'", (fid,))
    translated = Presentation(version['path']).slides[0].shapes[0].text_frame.paragraphs[0]
    assert translated.text == '[机翻]Step 1\v[机翻]Control panel'
    assert [run.font.size for run in translated.runs] == [Pt(28), Pt(18)]
    assert [run.font.bold for run in translated.runs] == [True, False]


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


def test_short_uppercase_glossary_term_does_not_match_lowercase_word():
    engine = TerminologyEngine([Term("to", "TO", "工艺对象")])
    protected, _mapping = engine.protect("Click to continue")
    assert protected == "Click to continue"
    assert not engine.validate("Click to continue", "单击以继续")


def test_missing_required_term_is_safely_retranslated_not_blindly_appended(monkeypatch):
    term = Term("unified", "Unified", "Unified")
    engine = TerminologyEngine([term])
    unit = {"stable_key": "1", "source_text": "Unified Basic Panel", "unit_type": "text"}
    file = {"source_language": "English", "target_language": "Chinese"}

    def protected_echo(_provider, request, _progress):
        # A provider returning the protected placeholder lets the engine restore
        # the required product term in its original sentence position.
        return {key: text for key, text in request.units}

    monkeypatch.setattr(pipeline, "translate", protected_echo)
    fixed, repaired = pipeline._retry_missing_terms(
        {"id": "mock"}, file, engine, [{"source": "Unified", "target": "Unified"}], [], [unit], {"1": "统一基础面板"}
    )
    assert repaired == 1
    assert fixed["1"] == "Unified Basic Panel"
    assert not engine.validate(unit["source_text"], fixed["1"])


def test_specific_glossary_term_suppresses_nested_generic_term_blocker():
    engine = TerminologyEngine([
        Term("comfort", "Comfort", "Comfort"),
        Term("comfort-panel", "Comfort Panel", "精智面板"),
    ])
    assert not engine.validate("Choose a Comfort Panel.", "请选择精智面板。")
    assert [issue for issue in engine.validate("Choose Comfort.", "请选择精智。") if issue["code"] == "TERM_MISSING"]


def test_ppt_alttext_qa_warns_but_does_not_block_release():
    from sitrain_app.pipeline import _qa_issues
    engine = TerminologyEngine([])
    issues = _qa_issues({"unit_type": "ppt_alttext", "source_text": "click left 2"}, engine, "单击左侧")
    assert any(issue["code"] == "NUMBER_CHANGED" and issue["severity"] == "warning" for issue in issues)


def test_word_page_field_qa_warns_but_does_not_block_release():
    from sitrain_app.pipeline import _qa_issues
    engine = TerminologyEngine([])
    unit = {"unit_type": "header_footer", "source_text": "{{PAGE_NUMBER}}1", "source_payload": {"field_info": [{"type": "field_instruction"}]}}
    issues = _qa_issues(unit, engine, "{{PAGE_NUMBER}}")
    assert any(issue["code"] == "NUMBER_CHANGED" and issue["severity"] == "warning" for issue in issues)


def test_glossary_change_rechecks_unconfirmed_qa_without_overwriting_text(client):
    uploaded = client.post("/api/files/upload", files={"upload": ("manual.txt", b"Click to continue\n", "text/plain")}).json()
    unit = client.get(f"/api/files/{uploaded['id']}/units").json()["units"][0]
    db.execute("UPDATE translation_units SET machine_text=?,human_text=?,status='terminology_issue',qa_json=? WHERE id=?",
               ("单击以继续", "人工译文保持不变", json.dumps([{"code":"TERM_MISSING","severity":"blocker"}]), unit["id"]))
    result = client.post("/api/glossary", json={"source":"TO", "target":"工艺对象"})
    assert result.status_code == 200
    assert result.json()["revalidation"]["units"] >= 1
    refreshed = client.get(f"/api/files/{uploaded['id']}/units").json()["units"][0]
    assert refreshed["human_text"] == "人工译文保持不变"
    assert refreshed["status"] == "pending_review"
    assert not [issue for issue in refreshed["qa"] if issue["code"] == "TERM_MISSING"]


def test_teacher_can_record_a_manual_term_exception_only_after_machine_blockers_are_gone(client):
    uploaded = client.post("/api/files/upload", files={"upload": ("manual-term.txt", b"Comfort\n", "text/plain")}).json()
    unit = client.get(f"/api/files/{uploaded['id']}/units").json()["units"][0]
    db.execute("UPDATE translation_units SET human_text=?,status='modified',qa_json=? WHERE id=?",
               ("精智面板", json.dumps([{"code": "TERM_MISSING", "severity": "blocker", "message": "必须使用术语"}]), unit["id"]))
    accepted = client.post(f"/api/files/{uploaded['id']}/confirm-manual-term-exceptions")
    assert accepted.status_code == 200 and accepted.json()["confirmed"] == 1
    refreshed = client.get(f"/api/files/{uploaded['id']}/units").json()["units"][0]
    assert refreshed["status"] == "confirmed"
    assert refreshed["qa"][0]["code"] == "TERM_EXCEPTION_CONFIRMED"
    db.execute("UPDATE translation_units SET human_text=NULL,status='terminology_issue',qa_json=? WHERE id=?",
               (json.dumps([{"code": "TERM_MISSING", "severity": "blocker"}]), unit["id"]))
    rejected = client.post(f"/api/files/{uploaded['id']}/confirm-manual-term-exceptions")
    assert rejected.status_code == 409


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
