import io
import zipfile
import os
import subprocess
import sys
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

from sitrain_app import db, storage, main
from sitrain_app.pipeline import add_version
from sitrain_app.review_versions import apply_safe_machine_term_repairs, current_version, preview_source


def png(color):
    b = io.BytesIO()
    Image.new('RGB', (160, 90), color).save(b, format='PNG')
    return b.getvalue()


@pytest.fixture
def deck_client(tmp_path, monkeypatch):
    monkeypatch.setattr(db, 'DB_PATH', tmp_path/'db.sqlite')
    monkeypatch.setattr(storage, 'STORAGE_ROOT', tmp_path/'storage')
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(10, 40, 80)
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
    p = box.text_frame.paragraphs[0]
    a = p.add_run(); a.text = 'Motor '; a.font.bold = True; a.font.size = Pt(24)
    b = p.add_run(); b.text = '24 V'; b.font.color.rgb = RGBColor(220, 0, 0)
    box.text_frame.add_paragraph().text = 'Second paragraph'
    image = png('blue')
    pic = slide.shapes.add_picture(io.BytesIO(image), Inches(1), Inches(3), Inches(2), Inches(1))
    pic.crop_left = .1
    slide.shapes.add_picture(io.BytesIO(image), Inches(4), Inches(3), Inches(2), Inches(1))
    table = slide.shapes.add_table(2, 2, Inches(1), Inches(5), Inches(4), Inches(1)).table
    table.cell(0, 0).merge(table.cell(0, 1)); table.cell(0, 0).text = 'Merged cell'
    group = slide.shapes.add_group_shape()
    group.shapes.add_textbox(Inches(6), Inches(2), Inches(2), Inches(1)).text = 'Grouped text'
    slide.shapes.add_textbox(Inches(6), Inches(4), Inches(2), Inches(1)).text = 'Shared source'
    second = deck.slides.add_slide(deck.slide_layouts[6])
    second.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1)).text = 'Untouched'
    second.shapes.add_textbox(Inches(4), Inches(1), Inches(2), Inches(1)).text = 'Shared source'
    fid = 'review-test'
    with TestClient(main.app) as client:
        source = storage.file_dir(fid)/'original.pptx'
        deck.save(source)
        now = db.utcnow()
        db.execute("""INSERT INTO files(id,original_name,extension,kind,mode,status,preflight_json,created_at,updated_at)
                      VALUES(?,?,?,?,?,?,?,?,?)""", (fid,'original.pptx','.pptx','powerpoint','standalone','review','{"pages":2}',now,now))
        add_version(fid, 'original', source)
        add_version(fid, 'machine', source)
        yield client, fid, source


def test_review_export_preserves_package_and_shared_picture(deck_client):
    client, fid, original = deck_client
    original_bytes = original.read_bytes()
    endpoint = f'/api/files/{fid}/pages/1'
    data = client.get(endpoint+'/objects').json()
    text = next(o for o in data['objects'] if o['paragraphs'] and o['paragraphs'][0] == 'Motor 24 V')
    result = client.patch(endpoint+'/text', json={'base_version':data['version_id'], 'object_id':text['id'],
                          'paragraphs':['电机 24 V','Second paragraph']})
    assert result.status_code == 200, result.text
    reviewed = current_version(fid)
    with zipfile.ZipFile(original) as a, zipfile.ZipFile(reviewed['path']) as b:
        assert a.namelist() == b.namelist()
        assert [name for name in a.namelist() if a.read(name) != b.read(name)] == ['ppt/slides/slide1.xml']
    deck = Presentation(reviewed['path'])
    paragraph = deck.slides[0].shapes[0].text_frame.paragraphs[0]
    assert paragraph.text == '电机 24 V'
    assert paragraph.runs[0].font.bold
    assert paragraph.runs[0].font.size == Pt(24)
    assert paragraph.runs[-1].font.color.rgb == RGBColor(220, 0, 0)
    # Optimistic concurrency prevents stale browser tabs from overwriting edits.
    stale = client.patch(endpoint+'/text', json={'base_version':data['version_id'], 'object_id':text['id'], 'paragraphs':['Wrong','Wrong']})
    assert stale.status_code == 409
    data = client.get(endpoint+'/objects').json()
    image = next(o for o in data['objects'] if o['kind'] == 'image')
    result = client.post(endpoint+'/image', data={'base_version':data['version_id'],'object_id':image['id']},
                         files={'upload':('new.png',png('red'),'image/png')})
    assert result.status_code == 200, result.text
    reviewed = current_version(fid)
    deck = Presentation(reviewed['path'])
    assert deck.slides[0].shapes[1].image.blob != deck.slides[0].shapes[2].image.blob
    assert deck.slides[0].shapes[2].image.blob == png('blue')
    assert deck.slides[0].shapes[1].crop_left == .1
    assert deck.slides[0].shapes[1].left == Inches(1)
    assert deck.slides[0].shapes[1].width == Inches(2)
    publish = client.post(f'/api/files/{fid}/publish')
    assert publish.status_code == 200, publish.text
    exported = client.get(f"/api/versions/{publish.json()['version_id']}/download")
    assert exported.content == Path(reviewed['path']).read_bytes()
    assert original.read_bytes() == original_bytes
    assert client.patch(endpoint+'/text', json={'base_version':reviewed['id'], 'object_id':text['id'], 'paragraphs':['Wrong','Wrong']}).status_code == 409


def test_cells_groups_validation_and_cache(deck_client):
    client, fid, original = deck_client
    endpoint = f'/api/files/{fid}/pages/1'
    data = client.get(endpoint+'/objects').json()
    cell = next(o for o in data['objects'] if o['paragraphs'] == ['Merged cell'])
    assert len(cell['id'].split(':')) == 3
    result = client.patch(endpoint+'/text',json={'base_version':data['version_id'],'object_id':cell['id'],'paragraphs':['合并单元格']})
    assert result.status_code == 200, result.text
    version = current_version(fid)
    # Page 2 reuses original cache; page 1 has an isolated cache and exact source.
    assert preview_source(version, 2)[0] == original
    assert preview_source(version, 1)[0] == Path(version['path'])
    assert preview_source(version, 1)[1].name == version['id']
    data = client.get(endpoint+'/objects').json()
    group = next(o for o in data['objects'] if o['paragraphs'] == ['Grouped text'])
    assert client.patch(endpoint+'/text',json={'base_version':data['version_id'],'object_id':group['id'],'paragraphs':['组合内文字\n换行']}).status_code == 200
    latest = client.get(endpoint+'/objects').json()
    assert next(o for o in latest['objects'] if o['id'] == group['id'])['paragraphs'] == ['组合内文字\n换行']
    assert client.post(f'/api/files/{fid}/translate').status_code == 409
    image = next(o for o in latest['objects'] if o['kind'] == 'image')
    assert client.post(endpoint+'/image', data={'base_version':latest['version_id'],'object_id':image['id']},files={'upload':('bad.png',b'not an image','image/png')}).status_code == 422
    assert current_version(fid)['id'] == latest['version_id']
    assert client.get(f'/api/files/{fid}/preview/0').status_code == 404
    assert client.get(f'/api/files/{fid}/preview/1?version_id=missing').status_code == 404


def test_reopen_published_copy_is_editable_and_preserves_old_delivery(deck_client):
    client, fid, original = deck_client
    old_id = add_version(fid, 'published', original)
    old_bytes = original.read_bytes()
    db.execute("UPDATE files SET status='published' WHERE id=?", (fid,))
    reopened = client.post(f'/api/files/{fid}/reopen-review')
    assert reopened.status_code == 200
    version = current_version(fid)
    assert version['version_type'] == 'review'
    assert version['metadata']['base_version'] == old_id
    assert Path(version['path']) != original
    assert Path(version['path']).read_bytes() == old_bytes
    assert preview_source(version, 2)[0] == original
    again = client.post(f'/api/files/{fid}/reopen-review').json()
    assert again['version_id'] == reopened.json()['version_id']
    data = client.get(f'/api/files/{fid}/pages/1/objects').json()
    obj = next(o for o in data['objects'] if o['paragraphs'] == ['Motor 24 V','Second paragraph'])
    result = client.patch(f'/api/files/{fid}/pages/1/text', json={'base_version':version['id'], 'object_id':obj['id'],
                         'paragraphs':['After reopening 24 V', 'Second paragraph']})
    assert result.status_code == 200
    published = client.post(f'/api/files/{fid}/publish').json()['version_id']
    assert published != old_id
    assert client.get(f'/api/versions/{old_id}/download').content == old_bytes
    assert client.get(f'/api/versions/{published}/download').content != old_bytes


def test_legacy_published_preview_reuses_completed_machine_cache(deck_client):
    _client, fid, original = deck_client
    published = add_version(fid, 'published', original)
    db.execute("UPDATE files SET status='published' WHERE id=?", (fid,))
    version = current_version(fid)
    source, cache = preview_source(version, 1)
    assert source == original
    assert cache == storage.file_dir(fid) / 'previews'
    assert version['id'] == published


def test_review_assist_uses_structural_source_mapping_for_memory_and_same_source(deck_client):
    client, fid, _original = deck_client
    first = client.get(f'/api/files/{fid}/pages/1/objects').json()
    second = client.get(f'/api/files/{fid}/pages/2/objects').json()
    first_obj = next(o for o in first['objects'] if o['paragraphs'] == ['Shared source'])
    second_obj = next(o for o in second['objects'] if o['paragraphs'] == ['Shared source'])
    for number, obj in ((1, first_obj), (2, second_obj)):
        ref = obj['unit_refs'][0]
        now = db.utcnow()
        db.execute("""INSERT INTO translation_units(id,file_id,stable_key,page,object_ref,unit_type,source_text,machine_text,status,source_payload_json,updated_at)
                      VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                   (f'unit-{number}', fid, str(number), number, str(number), ref['type'], 'Shared source',
                    '共享来源', 'translated', __import__('json').dumps({**ref, 'value': 'Shared source'}), now))
    context = client.get(f'/api/files/{fid}/pages/1/review-context', params={'object_id':first_obj['id']})
    assert context.status_code == 200 and context.json()['trusted']
    # First save is a normal PPTX edit; the assist operation then updates only
    # the structurally locatable second instance with the same original source.
    saved = client.patch(f'/api/files/{fid}/pages/1/text', json={'base_version':first['version_id'], 'object_id':first_obj['id'], 'paragraphs':['共享术语']})
    assert saved.status_code == 200, saved.text
    applied = client.post(f'/api/files/{fid}/pages/1/apply-same-source', json={'base_version':saved.json()['version_id'],'object_id':first_obj['id']})
    assert applied.status_code == 200, applied.text
    assert applied.json()['applied'] == 1
    refreshed = client.get(f'/api/files/{fid}/pages/2/objects').json()
    assert next(o for o in refreshed['objects'] if o['id'] == second_obj['id'])['paragraphs'] == ['共享术语']
    memory = client.post(f'/api/files/{fid}/pages/1/translation-memory', json={'base_version':applied.json()['version_id'],'object_id':first_obj['id']})
    assert memory.status_code == 200, memory.text
    assert db.scalar("SELECT target FROM translation_memory WHERE source='Shared source'") == '共享术语'
    assert db.scalar("SELECT COUNT(*) FROM review_changes WHERE file_id=?", (fid,)) == 2


def test_safe_term_repair_updates_only_untouched_ppt_objects(deck_client):
    client, fid, _original = deck_client
    first = client.get(f'/api/files/{fid}/pages/1/objects').json()
    second = client.get(f'/api/files/{fid}/pages/2/objects').json()
    edited = next(o for o in first['objects'] if o['paragraphs'] == ['Shared source'])
    untouched = next(o for o in second['objects'] if o['paragraphs'] == ['Shared source'])
    now = db.utcnow()
    for number, page, obj in ((1, 1, edited), (2, 2, untouched)):
        ref = obj['unit_refs'][0]
        db.execute("""INSERT INTO translation_units(id,file_id,stable_key,page,object_ref,unit_type,source_text,machine_text,status,source_payload_json,updated_at)
                      VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                   (f'term-{number}', fid, f'term-{number}', page, str(page), ref['type'], 'Shared source',
                    '旧译文', 'terminology_issue', __import__('json').dumps({**ref, 'value': 'Shared source'}), now))
    saved = client.patch(f'/api/files/{fid}/pages/1/text', json={'base_version': first['version_id'],
                         'object_id': edited['id'], 'paragraphs': ['老师修改']})
    assert saved.status_code == 200
    units = db.rows("SELECT * FROM translation_units WHERE file_id=? ORDER BY stable_key", (fid,))
    for unit in units:
        unit['machine_text'] = '已按术语修正'
    result = apply_safe_machine_term_repairs(fid, units)
    assert result['applied'] == 1 and result['skipped'] == 1
    assert client.get(f'/api/files/{fid}/pages/1/objects').json()['objects']
    page_one = client.get(f'/api/files/{fid}/pages/1/objects').json()
    page_two = client.get(f'/api/files/{fid}/pages/2/objects').json()
    assert next(o for o in page_one['objects'] if o['id'] == edited['id'])['paragraphs'] == ['老师修改']
    assert next(o for o in page_two['objects'] if o['id'] == untouched['id'])['paragraphs'] == ['已按术语修正']


@pytest.mark.skipif(os.name != 'nt', reason='Native PowerPoint smoke test is Windows-only')
def test_edited_export_opens_in_powerpoint(deck_client, tmp_path):
    client, fid, original = deck_client
    endpoint = f'/api/files/{fid}/pages/1'
    data = client.get(endpoint+'/objects').json()
    obj = next(o for o in data['objects'] if o['paragraphs'] == ['Motor 24 V','Second paragraph'])
    assert client.patch(endpoint+'/text', json={'base_version':data['version_id'],'object_id':obj['id'],
                        'paragraphs':['Reviewed motor 24 V','Second paragraph']}).status_code == 200
    data = client.get(endpoint+'/objects').json()
    pic = next(o for o in data['objects'] if o['kind'] == 'image')
    assert client.post(endpoint+'/image', data={'base_version':data['version_id'],'object_id':pic['id']},
                       files={'upload':('red.png',png('red'),'image/png')}).status_code == 200
    published = client.post(f'/api/files/{fid}/publish').json()['version_id']
    path = db.row('SELECT path FROM file_versions WHERE id=?', (published,))['path']
    helper = Path(__file__).resolve().parents[1]/'scripts'/'render_ppt_preview.py'
    for name, source in [('original', str(original)), ('edited', path)]:
        result = subprocess.run([sys.executable,str(helper),source,str(tmp_path/f'{name}.png'),'1'], timeout=60)
        assert result.returncode == 0
    assert (tmp_path/'original.png').read_bytes() != (tmp_path/'edited.png').read_bytes()
