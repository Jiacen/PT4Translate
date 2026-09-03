import io
import json
import time
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from sitrain_app import db, storage
from sitrain_app.main import app


@pytest.fixture()
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(db, "DB_PATH", tmp_path / "test.db")
    monkeypatch.setattr(storage, "STORAGE_ROOT", tmp_path / "storage")
    with TestClient(app) as c:
        db.execute("UPDATE providers SET protocol='mock',is_default=1,enabled=1 WHERE id='openai-compatible'")
        db.execute("UPDATE providers SET is_default=0 WHERE id<>'openai-compatible'")
        yield c


def test_health_and_dual_entry_data_model(client):
    assert client.get("/api/health").json()["ok"] is True
    project = client.post("/api/projects", json={"name":"TIA Portal 教材","description":"回归"}).json()
    chapter = client.post(f"/api/projects/{project['id']}/chapters", json={"name":"用户界面"}).json()
    standalone = client.post("/api/files/upload", files={"upload":("standalone.txt",b"Hello operator\n","text/plain")})
    bundled = client.post("/api/files/upload", data={"project_id":project["id"],"chapter_id":chapter["id"]},
                          files={"upload":("chapter.txt",b"Commissioning device\n","text/plain")})
    assert standalone.status_code == bundled.status_code == 200
    dash = client.get("/api/dashboard").json()
    modes = {f["mode"] for f in dash["files"]}
    assert modes == {"standalone","chapter"}


def test_txt_upload_translate_review_publish(client):
    uploaded = client.post("/api/files/upload", files={"upload":("manual.txt",b"Start the motor 24 V\nStop the motor\n","text/plain")}).json()
    fid = uploaded["id"]
    detail = client.get(f"/api/files/{fid}").json()
    assert detail["status"] == "ready"
    assert detail["preflight"]["text_units"] == 2
    task = client.post(f"/api/files/{fid}/translate").json()
    task_state = next(t for t in client.get("/api/dashboard").json()["tasks"] if t["id"] == task["id"])
    assert task_state["status"] == "completed"
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
    from pptx.util import Inches
    path = tmp_path / "sample.pptx"
    prs = Presentation(); slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(Inches(1),Inches(1),Inches(5),Inches(1)).text = "Operator control panel"
    prs.save(path)
    with path.open("rb") as stream:
        fid = client.post("/api/files/upload", files={"upload":(path.name,stream,"application/vnd.openxmlformats-officedocument.presentationml.presentation")}).json()["id"]
    assert client.get(f"/api/files/{fid}").json()["status"] == "ready"
    client.post(f"/api/files/{fid}/translate")
    units = client.get(f"/api/files/{fid}/units").json()["units"]
    assert any("Operator control panel" in u["source_text"] for u in units)
    for unit in units:
        client.patch(f"/api/units/{unit['id']}",json={"human_text":"操作员控制面板","status":"confirmed","revision":unit["revision"]})
    response = client.post(f"/api/files/{fid}/publish")
    assert response.status_code == 200, response.text
    out = tmp_path / "out.pptx"
    out.write_bytes(client.get(f"/api/versions/{response.json()['version_id']}/download").content)
    assert "操作员控制面板" in Presentation(out).slides[0].shapes[0].text
    preview = client.get(f"/api/files/{fid}/preview/1")
    assert preview.status_code == 200 and preview.headers["content-type"].startswith("image/png")


def test_server_rejects_disallowed_extension(client):
    response = client.post("/api/files/upload", files={"upload":("picture.png",b"not an image","image/png")})
    assert response.status_code == 415


def test_office_embedded_image_enters_ocr_and_is_rebuilt(client, tmp_path):
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
    units = client.get(f"/api/files/{fid}/units").json()["units"]
    ocr_units = [u for u in units if u["unit_type"] == "office_image_ocr"]
    assert ocr_units and all(u["confidence"] > 0 for u in ocr_units)
    for unit in ocr_units:
        response = client.patch(f"/api/units/{unit['id']}", json={"human_text":"启动电机 24 V","status":"confirmed","revision":unit["revision"]})
        assert response.status_code == 200
    published = client.post(f"/api/files/{fid}/publish")
    assert published.status_code == 200, published.text
    import zipfile
    output = tmp_path / "ocr_output.pptx"
    output.write_bytes(client.get(f"/api/versions/{published.json()['version_id']}/download").content)
    with zipfile.ZipFile(ppt_path) as before, zipfile.ZipFile(output) as after:
        assert before.read("ppt/media/image1.png") != after.read("ppt/media/image1.png")


def test_api_never_returns_provider_secret(client):
    db.execute("UPDATE providers SET api_key='super-secret' WHERE id='openai-compatible'")
    body = client.get("/api/providers").text
    assert "super-secret" not in body
    assert "api_key_configured" in body
