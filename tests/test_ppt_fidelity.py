import importlib.util
import io
import sys
import zipfile
from pathlib import Path

import pytest
from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

from sitrain_app import documents, ppt_render, providers
from sitrain_app.ppt_fidelity import NS, inherited_part_pages, paragraph_line_groups, verify_pptx_preservation


@pytest.fixture
def rich_deck(tmp_path):
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(0, 0, 40)
    picture = io.BytesIO()
    Image.new('RGB', (160, 90), 'navy').save(picture, format='PNG')
    slide.shapes.add_picture(io.BytesIO(picture.getvalue()), Inches(5), Inches(1), Inches(3), Inches(2))
    title = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    title.text_frame.paragraphs[0].text = 'Step 1'
    title.text_frame.paragraphs[0].runs[0].font.bold = True
    title.text_frame.paragraphs[0].runs[0].font.name = 'Arial'
    title.text_frame.paragraphs[0].runs[0].font.size = Pt(28)
    title.text_frame.paragraphs[0]._p.append(OxmlElement('a:br'))
    run = title.text_frame.paragraphs[0].add_run()
    run.text = 'Control panel'
    run.font.size = Pt(18)
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(3), Inches(3), Inches(1))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0, 153, 153)
    paragraph = box.text_frame.paragraphs[0]
    paragraph.text = 'Start the drive'
    bullet = OxmlElement('a:buChar'); bullet.set('char', '•')
    paragraph._p.get_or_add_pPr().append(bullet)
    paragraph.runs[0].font.size = Pt(24)
    paragraph.runs[0].font.color.rgb = RGBColor(255, 255, 255)
    group = slide.shapes.add_group_shape()
    for number in range(1, 9):
        shape = group.shapes.add_shape(MSO_SHAPE.OVAL, Inches(number / 2), Inches(5), Inches(.4), Inches(.4))
        shape.text = str(number)
    table = slide.shapes.add_table(1, 1, Inches(5), Inches(4), Inches(3), Inches(1)).table
    table.cell(0, 0).text = 'Motor'
    table.cell(0, 0).text_frame.paragraphs[0].runs[0].font.size = Pt(32)
    # A video is never decoded or executed by translation; its bytes, poster
    # image and relationship must survive exactly, including unknown extensions.
    clip = tmp_path / 'clip.mp4'
    clip.write_bytes(b'opaque-test-video-do-not-execute')
    slide.shapes.add_movie(str(clip), Inches(8), Inches(1), Inches(1), Inches(1),
                           poster_frame_image=io.BytesIO(picture.getvalue()), mime_type='video/mp4')
    source = tmp_path / 'source.pptx'
    deck.save(source)
    return source


def test_translation_keeps_background_video_callouts_bullets_and_fonts(rich_deck, tmp_path):
    source_json, units = documents.extract_units(rich_deck, tmp_path / 'extraction')
    title = next(unit for unit in units if unit['source_text'] == 'Step 1\nControl panel')
    for unit in units:
        unit['machine_text'] = '步骤 1\n控制面板' if unit is title else '保持原字体字号的中文翻译' * 4
    output = documents.rebuild(rich_deck, source_json, units, tmp_path / 'output', 'English', 'Chinese')
    assert verify_pptx_preservation(rich_deck, output)['policy'] == 'preserve_original'
    with zipfile.ZipFile(rich_deck) as before, zipfile.ZipFile(output) as after:
        for name in before.namelist():
            if name.startswith(('ppt/media/', 'ppt/theme/')) or name.endswith('.rels'):
                assert before.read(name) == after.read(name)
        original = etree.fromstring(before.read('ppt/slides/slide1.xml'))
        translated = etree.fromstring(after.read('ppt/slides/slide1.xml'))
    for xpath in ('.//a:rPr/@sz', './/a:latin/@typeface', './/a:buChar/@char'):
        assert original.xpath(xpath, namespaces=NS) == translated.xpath(xpath, namespaces=NS)
    assert not translated.xpath('.//a:normAutofit', namespaces=NS)
    deck = Presentation(output)
    assert deck.slides[0].shapes[1].text == '步骤 1\v控制面板'
    assert [shape.text for shape in deck.slides[0].shapes[3].shapes] == [str(n) for n in range(1, 9)]
    assert deck.slides[0].shapes[4].table.cell(0, 0).text_frame.paragraphs[0].runs[0].font.size == Pt(32)


def test_ppt_layout_usage_distinguishes_unused_template_parts(rich_deck):
    deck = Presentation(rich_deck)
    used_layout = str(deck.slides[0].slide_layout.part.partname).lstrip('/')
    unused_layout = str(deck.slide_layouts[0].part.partname).lstrip('/')
    master = str(deck.slide_masters[0].part.partname).lstrip('/')
    usage = inherited_part_pages(rich_deck)
    assert usage[used_layout] == [1]
    assert usage[unused_layout] == []
    assert usage[master] == [1]


def test_unverifiable_ppt_layout_usage_is_not_assumed_unused(tmp_path):
    source = tmp_path / 'broken.pptx'
    source.write_bytes(b'not a zip')
    with pytest.raises(ValueError, match='could not be verified'):
        inherited_part_pages(source)


@pytest.mark.parametrize('damage', ['media', 'font', 'bullet', 'geometry', 'background'])
def test_fidelity_guard_rejects_non_text_damage(rich_deck, tmp_path, damage):
    output = tmp_path / f'damaged-{damage}.pptx'
    with zipfile.ZipFile(rich_deck) as before, zipfile.ZipFile(output, 'x') as after:
        for item in before.infolist():
            content = before.read(item.filename)
            if damage == 'media' and item.filename.endswith('.mp4'):
                content = b'damaged-video'
            elif item.filename == 'ppt/slides/slide1.xml' and damage != 'media':
                root = etree.fromstring(content)
                if damage == 'font': root.xpath('.//a:rPr[@sz]', namespaces=NS)[0].set('sz', '700')
                if damage == 'bullet': root.xpath('.//a:buChar', namespaces=NS)[0].set('char', '-')
                if damage == 'geometry': root.xpath('.//a:xfrm/a:off', namespaces=NS)[0].set('x', '1')
                if damage == 'background': root.xpath('.//p:bg//a:srgbClr', namespaces=NS)[0].set('val', 'FFFFFF')
                content = etree.tostring(root)
            after.writestr(item, content)
    with pytest.raises(RuntimeError, match='格式保真检查失败'):
        verify_pptx_preservation(rich_deck, output)


def test_provider_preserves_authored_lines_tabs_and_blank_lines(monkeypatch):
    calls = []
    expected = {'Step 1': '步骤1', 'Control panel': '控制面板', 'Drive:': '驱动器：', 'DI/DQ': 'DI/DQ'}

    def fake_qwen(config, request, progress):
        calls.extend(text for _, text in request.units)
        return {key: expected[text] for key, text in request.units}

    monkeypatch.setattr(providers, 'translate_qwen', fake_qwen)
    request = providers.TranslationRequest([('title', 'Step 1\nControl panel'), ('detail', 'Drive:\tDI/DQ\n\n')],
                                           'English', 'Chinese', [], [], preserve_line_breaks=True)
    result = providers.translate({'protocol': 'qwen_mt'}, request)
    assert result == {'title': '步骤1\n控制面板', 'detail': '驱动器：\tDI/DQ\n\n'}
    assert calls == ['Step 1', 'Control panel', 'Drive:', 'DI/DQ']


def test_provider_rejects_missing_line_translations(monkeypatch):
    monkeypatch.setattr(providers, 'translate_qwen', lambda *args: {})
    request = providers.TranslationRequest([('title', 'Step 1\nControl panel')], 'English', 'Chinese', [], [], True)
    with pytest.raises(providers.ProviderError, match='分行译文缺项'):
        providers.translate({'protocol': 'qwen_mt'}, request)


def test_uniformly_styled_sentence_keeps_context_across_a_soft_break(monkeypatch):
    source = 'Get to know the controls\nof the demo case'
    payload = {'value': source, 'run_lengths': [24, 16], 'run_styles': [{'bold': '1'}, {'bold': '1'}]}
    assert paragraph_line_groups(payload) == [2]
    payload['run_styles'][1] = {'bold': '0'}
    assert paragraph_line_groups(payload) == [1, 1]
    seen = []

    def fake_qwen(config, request, progress):
        seen.extend(text for _, text in request.units)
        return {key: '了解演示箱的操作控件' for key, _ in request.units}

    monkeypatch.setattr(providers, 'translate_qwen', fake_qwen)
    request = providers.TranslationRequest([('sentence', source)], 'English', 'Chinese', [], [], True, {'sentence': [2]})
    result = providers.translate({'protocol': 'qwen_mt'}, request)['sentence']
    assert seen == ['Get to know the controls of the demo case']
    assert result.replace('\n', '') == '了解演示箱的操作控件'
    assert result.count('\n') == 1


def test_frozen_renderer_uses_executable_worker_instead_of_a_missing_script(monkeypatch, tmp_path):
    monkeypatch.setattr(sys, 'frozen', True, raising=False)
    command = documents._ppt_render_command(tmp_path / 'source.pptx', tmp_path / 'page.png', '1')
    assert command == [sys.executable, '--render-ppt-preview', str(tmp_path / 'source.pptx'), str(tmp_path / 'page.png'), '1']


def test_renderer_launcher_does_not_start_web_server(monkeypatch, tmp_path):
    launcher = Path(__file__).resolve().parents[1] / 'scripts' / 'run_sitrain.py'
    spec = importlib.util.spec_from_file_location('renderer_launcher_test', launcher)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    calls = []
    monkeypatch.setattr(sys, 'argv', ['SITRAIN-Teacher.exe', '--render-ppt-preview', 'source.pptx', 'page.png', '1'])
    monkeypatch.setattr(ppt_render, 'main', lambda *args: calls.append(args) or 0)
    monkeypatch.setenv('SITRAIN_DATA_ROOT', str(tmp_path / 'must-not-be-created'))
    monkeypatch.setattr(module.webbrowser, 'open', lambda *args: pytest.fail('Renderer opened a browser'))
    assert module.main() == 0
    assert calls == [('source.pptx', 'page.png', '1')]
    assert not (tmp_path / 'must-not-be-created').exists()


def test_preview_rejects_simplified_rendering_and_ignores_old_cache(monkeypatch, tmp_path):
    cache = tmp_path / 'cache'; cache.mkdir()
    stale = cache / 'page-1-rendered.png'
    stale.write_bytes(b'old-simplified-preview')
    monkeypatch.setattr(documents, '_render_pptx_with_powerpoint', lambda *args: False)
    monkeypatch.setattr(documents, '_render_pptx_with_libreoffice', lambda *args: None)
    with pytest.raises(RuntimeError, match='保留原版式'):
        documents.render_preview(tmp_path / 'source.pptx', 1, cache)
    assert stale.read_bytes() == b'old-simplified-preview'
    assert not (cache / documents.PPT_RENDER_CACHE_VERSION / stale.name).exists()
    calls = []

    def render(source, page, target):
        calls.append(page)
        Image.new('RGB', (160, 90), 'navy').save(target)
        return True

    monkeypatch.setattr(documents, '_render_pptx_with_powerpoint', render)
    target = documents.render_preview(tmp_path / 'source.pptx', 1, cache)
    assert target.parent.name == documents.PPT_RENDER_CACHE_VERSION
    assert documents.render_preview(tmp_path / 'source.pptx', 1, cache) == target
    assert calls == [1]


def test_deck_preparation_uses_the_same_new_renderer_cache(monkeypatch, tmp_path):
    calls = []

    def render(source, cache):
        calls.append(cache)
        for page in (1, 2): Image.new('RGB', (160, 90), 'navy').save(cache / f'page-{page}-rendered.png')
        return True

    monkeypatch.setattr(documents, '_render_pptx_deck_with_powerpoint', render)
    documents.prepare_preview_cache(tmp_path / 'source.pptx', tmp_path / 'cache', 2)
    assert calls == [tmp_path / 'cache' / documents.PPT_RENDER_CACHE_VERSION]
