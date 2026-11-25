import io
import json
import zipfile
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from pptx import Presentation
from docx import Document
from PIL import Image

from sitrain_app import db, main, storage, sync
from sitrain_app.pipeline import add_version
from sitrain_app.word_image_index import bookmarked_copy, W


@pytest.fixture
def pair(tmp_path, monkeypatch):
    monkeypatch.setattr(db, "DB_PATH", tmp_path / "test.db")
    monkeypatch.setattr(storage, "STORAGE_ROOT", tmp_path / "storage")
    db.init_db()
    now = db.utcnow()
    db.execute("INSERT INTO projects(id,name,created_at,updated_at) VALUES('p','课件',?,?)", (now, now))
    db.execute("INSERT INTO chapters(id,project_id,name,created_at,updated_at) VALUES('c','p','章节',?,?)", (now, now))
    ppt = tmp_path / "source.pptx"
    deck = Presentation()
    for _ in range(3): deck.slides.add_slide(deck.slide_layouts[6])
    deck.save(ppt)
    word = tmp_path / "source.docx"
    Document().save(word)
    for fid, path, kind in [("ppt", ppt, "powerpoint"), ("word", word, "word")]:
        db.execute("INSERT INTO files(id,project_id,chapter_id,original_name,extension,kind,mode,status,preflight_json,created_at,updated_at) VALUES(?, 'p','c',?,?,?,'chapter','published','{\"images\":3}',?,?)",
                   (fid, path.name, path.suffix, kind, now, now))
        add_version(fid, "original", path, {})
        add_version(fid, "published", path, {})
    monkeypatch.setattr(sync, "_visual_mappings", lambda *_: [
        ("word/media/image1.emf", 1, .9, True, [{"slide_number": 1}]),
        ("word/media/image2.emf", 2, .3, False, [{"slide_number": 2}, {"slide_number": 3}]),
        ("word/media/image10.emf", 3, .2, False, [{"slide_number": 3}])])
    monkeypatch.setattr(sync, "_changed_word_pages", lambda *_: [2])
    monkeypatch.setattr(sync, "_export_replacements", lambda *_: {})
    with TestClient(main.app) as client:
        yield client, tmp_path


def test_progress_resume_idempotence_and_explicit_finish(pair):
    client, _ = pair
    mappings = client.post('/api/chapters/c/sync/map').json()['mappings']
    pending = client.get('/api/chapters/c/sync/pending?limit=1').json()
    assert pending['progress'] == {'initial_pending': 2, 'completed': 0, 'remaining': 2, 'automatic': 1}
    assert len(pending['mappings']) == 1
    assert pending['mappings'][0]['word_pages'] == []  # No fabricated image-number page.
    assert client.post('/api/chapters/c/sync/apply').status_code == 409
    first = pending['mappings'][0]['id']
    url = f'/api/mappings/{first}/decision'
    assert client.post(url, json={'slide_number': 9}).status_code == 409
    assert client.post(url, json={'slide_number': 2, 'keep_original': True}).status_code == 409
    assert client.post(url, json={'slide_number': 2}).status_code == 200
    # A repeated click is harmless; a stale tab cannot overwrite a decision.
    assert client.post(url, json={'slide_number': 2}).status_code == 200
    assert client.post(url, json={'slide_number': 3}).status_code == 409
    assert db.scalar("SELECT COUNT(*) FROM file_versions WHERE version_type='sync_review'") == 0
    # Reopening and even rerunning recognition preserves IDs and decisions.
    db.init_db()
    rerun = client.post('/api/chapters/c/sync/map').json()['mappings']
    assert {m['id'] for m in mappings} == {m['id'] for m in rerun}
    next_item = client.get('/api/chapters/c/sync/pending?limit=1').json()
    assert next_item['progress']['completed'] == next_item['progress']['remaining'] == 1
    assert next_item['mappings'][0]['id'] != first
    assert client.post(f"/api/mappings/{next_item['mappings'][0]['id']}/decision", json={'keep_original': True}).status_code == 200
    finished = client.get('/api/chapters/c/sync/pending?limit=1').json()
    assert finished['mappings'] == []
    assert finished['progress']['initial_pending'] == finished['progress']['completed'] == 2
    version = client.post('/api/chapters/c/sync/apply').json()['version_id']
    assert client.post('/api/chapters/c/sync/apply').json()['version_id'] == version
    assert client.post('/api/files/word/sync-review/confirm').status_code == 200
    assert main._chapter_sync_gate('c', db.rows("SELECT * FROM files")) is None


def test_legacy_pending_blocks_delivery_until_regenerated(pair):
    client, tmp_path = pair
    sync.create_mappings('c')
    # Simulate fix1: a partially synchronized Word was already confirmed.
    legacy = add_version('word', 'sync_review', tmp_path / 'source.docx', {'sync_review_confirmed': True, 'ppt_sync': True})
    assert client.post('/api/files/word/sync-review/confirm').status_code == 409
    assert client.get('/api/chapters/c/package').status_code == 409
    pending = client.get('/api/chapters/c/sync/pending').json()
    for item in pending['mappings']:
        assert client.post(f"/api/mappings/{item['id']}/decision", json={'keep_original': True}).status_code == 200
    assert (db.row("SELECT * FROM file_versions WHERE id=?", (legacy,))['metadata'])['mapping_choices_changed']
    assert client.post('/api/files/word/sync-review/confirm').status_code == 409
    assert client.post('/api/chapters/c/sync/apply').status_code == 200
    assert client.post('/api/files/word/sync-review/confirm').status_code == 200


def test_real_page_cache_and_legacy_batch_count(pair):
    client, _ = pair
    sync.create_mappings('c')
    db.execute("DELETE FROM mapping_batches WHERE chapter_id='c'")
    db.execute("UPDATE ppt_word_mappings SET decision='manual',confirmed=1 WHERE media_part LIKE '%image2.emf'")
    _, target = sync._page_index_path('word')
    target.parent.mkdir(parents=True)
    target.write_text(json.dumps({'status': 'ready', 'pages': {'word/media/image10.emf': [7, 9]},
                                 'locations': {'word/media/image10.emf': [{'page': 7, 'printed_page': 5, 'position': 400}]}}), encoding='utf-8')
    pending = client.get('/api/chapters/c/sync/pending?limit=1').json()
    assert pending['progress']['initial_pending'] == 2
    assert pending['progress']['completed'] == 1
    assert pending['mappings'][0]['word_pages'] == [7, 9]
    # A changed published baseline must not reuse page numbers from the old one.
    add_version('word', 'published', target, {})
    assert sync.word_page_index('word')['status'] == 'not_loaded'


def test_all_keep_original_requires_no_ppt_export(pair, monkeypatch):
    client, _ = pair
    sync.create_mappings('c')
    db.execute("UPDATE ppt_word_mappings SET confirmed=0,decision='pending'")
    # Existing counts are a snapshot of the initial batch, not remaining rows.
    db.execute("UPDATE mapping_batches SET initial_pending=3")
    for item in client.get('/api/chapters/c/sync/pending').json()['mappings']:
        client.post(f"/api/mappings/{item['id']}/decision", json={'keep_original': True})
    def should_not_export(*_): raise AssertionError('No replaced image should invoke PowerPoint')
    monkeypatch.setattr(sync, '_export_replacements', should_not_export)
    response = client.post('/api/chapters/c/sync/apply')
    assert response.status_code == 200, response.text
    assert response.json()['mapping_count'] == 0
    assert response.json()['changed_pages'] == []


def test_bookmarks_target_actual_relationships_and_preserve_source(tmp_path):
    from lxml import etree
    picture = io.BytesIO()
    Image.new('RGB', (20, 20), 'teal').save(picture, 'PNG')
    source, target = tmp_path / 'word.docx', tmp_path / 'indexed.docx'
    document = Document()
    document.add_picture(io.BytesIO(picture.getvalue()))
    document.add_page_break()
    document.add_picture(io.BytesIO(picture.getvalue()))  # shared media, distinct occurrences
    document.save(source)
    original = source.read_bytes()
    bookmarks = bookmarked_copy(source, target)
    assert len(bookmarks) == 2
    assert len({p for parts in bookmarks.values() for p in parts}) == 1
    assert source.read_bytes() == original
    with zipfile.ZipFile(source) as before, zipfile.ZipFile(target) as after:
        for name in before.namelist():
            if name != 'word/document.xml': assert before.read(name) == after.read(name)
        root = etree.fromstring(after.read('word/document.xml'))
        assert len(list(root.iter(f'{{{W}}}bookmarkStart'))) == 2
        assert len(list(root.iter(f'{{{W}}}drawing'))) == 2
