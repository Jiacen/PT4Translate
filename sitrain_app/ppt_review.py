"""PPTX review edits: patch selected XML parts, never reserialize the whole deck.

python-pptx is used only to read geometry. Saving through it would discard some
Office extensions. All untouched ZIP members are copied byte-for-byte instead.
"""
from __future__ import annotations

from copy import deepcopy
from difflib import SequenceMatcher
from functools import lru_cache
from io import BytesIO
from pathlib import Path
import posixpath
import re
import uuid
import zipfile

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def xml(data):
    return etree.fromstring(data, etree.XMLParser(resolve_entities=False, no_network=True))


def serialize(root):
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


@lru_cache(maxsize=48)
def page_objects(path: str, stamp: int, page: int):
    deck = Presentation(path)
    if not 1 <= page <= len(deck.slides):
        raise ValueError("页码超出范围")
    slide = deck.slides[page - 1]
    objects = []
    # Keep a structural reference back to the extraction pipeline.  It is not
    # based on visible text (which may occur many times on the same slide), so
    # review actions can safely find the original translation unit later.
    with zipfile.ZipFile(path) as package:
        slide_part = str(slide.part.partname).lstrip("/")
        slide_xml = xml(package.read(slide_part))
    text_box_by_shape = {}
    for index, tx_body in enumerate(slide_xml.xpath(".//p:txBody", namespaces=NS), start=1):
        shape = tx_body.getparent()
        c_nv_pr = shape.find("p:nvSpPr/p:cNvPr", NS) if shape is not None else None
        if c_nv_pr is not None:
            text_box_by_shape[c_nv_pr.get("id")] = index
    table_by_shape = {}
    for index, table in enumerate(slide_xml.xpath(".//a:tbl", namespaces=NS), start=1):
        frame = table.xpath("ancestor::p:graphicFrame[1]", namespaces=NS)
        c_nv_pr = frame[0].find("p:nvGraphicFramePr/p:cNvPr", NS) if frame else None
        if c_nv_pr is not None:
            table_by_shape[c_nv_pr.get("id")] = index

    def add(shape, box, cell=None):
        element = cell._tc if cell is not None else shape._element
        paragraphs = element.xpath("./a:txBody/a:p" if cell is not None else "./p:txBody/a:p")
        texts = ["".join("\n" if etree.QName(n).localname == "br" else (n.text or "")
                         for n in p.iter() if etree.QName(n).localname in {"t", "br"}) for p in paragraphs]
        picture = shape.shape_type == MSO_SHAPE_TYPE.PICTURE
        if not picture and not paragraphs:
            return
        sid = str(shape.shape_id)
        unit_refs = []
        if not picture:
            if cell is not None:
                table_index = table_by_shape.get(sid)
                if table_index:
                    unit_refs = [{"type": "table_cell_paragraph", "table_index": table_index,
                                  "row_index": cell._tc.getparent().getparent().index(cell._tc.getparent()),
                                  "cell_index": cell._tc.getparent().index(cell._tc),
                                  "paragraph_index": index}
                                 for index in range(len(paragraphs))]
            elif sid in text_box_by_shape:
                unit_refs = [{"type": "text_paragraph", "text_box_index": text_box_by_shape[sid],
                              "paragraph_index": index} for index in range(len(paragraphs))]
        objects.append({"id": sid, "kind": "image" if picture else "text", "name": shape.name,
                        "box": [box[0]/deck.slide_width, box[1]/deck.slide_height,
                                box[2]/deck.slide_width, box[3]/deck.slide_height],
                        "rotation": shape.rotation, "paragraphs": texts, "unit_refs": unit_refs})

    def visit(shapes, sx=1, sy=1, ox=0, oy=0):
        for shape in shapes:
            left, top = shape.left or 0, shape.top or 0
            width, height = shape.width or 1, shape.height or 1
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                x = shape._element.grpSpPr.xfrm
                gx = width / max(1, x.chExt.cx)
                gy = height / max(1, x.chExt.cy)
                visit(shape.shapes, sx*gx, sy*gy, ox+sx*(left-gx*x.chOff.x), oy+sy*(top-gy*x.chOff.y))
            elif shape.has_table:
                y = top
                for ri, row in enumerate(shape.table.rows):
                    x = left
                    for ci, cell in enumerate(row.cells):
                        w = shape.table.columns[ci].width
                        if not cell.is_spanned:
                            cw = sum(shape.table.columns[k].width for k in range(ci, ci+cell.span_width))
                            ch = sum(shape.table.rows[k].height for k in range(ri, ri+cell.span_height))
                            add(shape, (ox+sx*x, oy+sy*y, sx*cw, sy*ch), cell)
                            objects[-1]["id"] = f"{shape.shape_id}:{ri}:{ci}"
                            objects[-1]["name"] += f" · 单元格 {ri+1},{ci+1}"
                        x += w
                    y += row.height
            else:
                add(shape, (ox+sx*left, oy+sy*top, sx*width, sy*height))
    visit(slide.shapes)
    return {"page": page, "width": deck.slide_width, "height": deck.slide_height,
            "part": str(slide.part.partname).lstrip("/"), "objects": objects}


def _paragraph_text(p):
    return "".join("\n" if etree.QName(n).localname == "br" else (n.text or "")
                   for n in p.iter() if etree.QName(n).localname in {"t", "br"})


def replace_text(p, text):
    """Preserve unchanged runs exactly; inserted characters inherit nearby style."""
    old = _paragraph_text(p)
    if old == text:
        return
    if len(text) > 20000 or any(ord(c) < 32 and c not in "\n\t" for c in text):
        raise ValueError("文字过长或包含无效控制字符")
    tokens, owners = [], []
    for node in p:
        kind = etree.QName(node).localname
        if kind not in {"r", "br", "fld"}:
            continue
        value = "\n" if kind == "br" else "".join(node.xpath(".//a:t/text()", namespaces=NS))
        owners.extend([len(tokens)] * len(value))
        tokens.append(node)
    # Fields (page numbers/dates) must not be silently converted to ordinary text.
    if any(etree.QName(n).localname == "fld" for n in tokens):
        raise ValueError("此文字包含自动字段，请在 PowerPoint 中修改该字段")
    mapped = []
    for tag, i, j, a, b in SequenceMatcher(None, old, text, autojunk=False).get_opcodes():
        if tag == "equal":
            mapped.extend(owners[i:j])
        elif tag in {"insert", "replace"}:
            mapped.extend([owners[min(i, len(owners)-1)] if owners else -1] * (b-a))
    for node in tokens:
        p.remove(node)
    end = p.find("a:endParaRPr", NS)
    offset = p.index(end) if end is not None else len(p)
    start = 0
    while start < len(text):
        finish = start+1
        if text[start] != "\n":
            while finish < len(text) and text[finish] != "\n" and mapped[finish] == mapped[start]:
                finish += 1
        node = etree.Element(f"{{{NS['a']}}}{'br' if text[start] == chr(10) else 'r'}")
        owner = tokens[mapped[start]] if mapped[start] >= 0 else None
        style = owner.find("a:rPr", NS) if owner is not None else None
        if style is not None:
            node.append(deepcopy(style))
        if text[start] != "\n":
            t = etree.SubElement(node, f"{{{NS['a']}}}t")
            t.text = text[start:finish]
        p.insert(offset, node)
        offset += 1
        start = finish


def edit_pptx(source: Path, output: Path, page: int, object_id: str,
              paragraphs: list[str] | None = None, image: bytes | None = None):
    data = page_objects(str(source), source.stat().st_mtime_ns, page)
    obj = next((o for o in data["objects"] if o["id"] == object_id), None)
    if not obj:
        raise ValueError("未找到可编辑的对象，请刷新页面")
    part = data["part"]
    with zipfile.ZipFile(source) as z:
        root = xml(z.read(part))
        sid, *cell = object_id.split(":")
        candidates = root.xpath(".//p:cNvPr[@id=$sid]/../..", namespaces=NS, sid=sid)
        if len(candidates) != 1:
            raise ValueError("对象定位不唯一")
        shape = candidates[0]
        changes = {}
        if paragraphs is not None:
            if obj["kind"] != "text":
                raise ValueError("该对象不是文本")
            if cell:
                ri, ci = map(int, cell)
                target = shape.find("a:graphic/a:graphicData/a:tbl", NS).findall("a:tr", NS)[ri].findall("a:tc", NS)[ci]
                ps = target.findall("a:txBody/a:p", NS)
            else:
                ps = shape.findall("p:txBody/a:p", NS)
            if len(ps) != len(paragraphs):
                raise ValueError("段落数量已变化，请刷新；段内换行可以直接输入")
            for p, text in zip(ps, paragraphs):
                replace_text(p, text)
        elif image is not None:
            if obj["kind"] != "image" or len(image) > 10*1024*1024:
                raise ValueError("请选择图片对象，上传图片不能超过 10 MB")
            with Image.open(BytesIO(image)) as im:
                if im.format not in {"PNG", "JPEG", "WEBP"} or im.width*im.height > 25_000_000:
                    raise ValueError("仅支持 PNG、JPEG、WebP，图片不能超过 2500 万像素")
                im.load()
                buf = BytesIO()
                im.convert("RGBA").save(buf, format="PNG")
            relpart = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part)+".rels")
            rels = xml(z.read(relpart))
            rid = "rIdReview" + uuid.uuid4().hex
            media = "ppt/media/review-" + uuid.uuid4().hex + ".png"
            etree.SubElement(rels, f"{{{REL}}}Relationship", Id=rid,
                             Type=NS["r"]+"/image", Target=posixpath.relpath(media, posixpath.dirname(part)))
            blip = shape.find("p:blipFill/a:blip", NS)
            blip.set(f"{{{NS['r']}}}embed", rid)
            blip.attrib.pop(f"{{{NS['r']}}}link", None)
            # Remove alternate SVG/image references that would otherwise still display the old image.
            for ext in list(blip):
                if etree.QName(ext).localname == "extLst":
                    blip.remove(ext)
            changes[relpart] = serialize(rels)
            changes[media] = buf.getvalue()
            types = xml(z.read("[Content_Types].xml"))
            if not any(n.get("Extension", "").lower() == "png" for n in types):
                etree.SubElement(types, f"{{{CT}}}Default", Extension="png", ContentType="image/png")
                changes["[Content_Types].xml"] = serialize(types)
        else:
            raise ValueError("没有提供修改内容")
        changes[part] = serialize(root)
        output.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output, "x", compression=zipfile.ZIP_DEFLATED) as out:
            for item in z.infolist():
                out.writestr(item, changes.pop(item.filename, z.read(item.filename)))
            for name, content in changes.items():
                out.writestr(name, content)
    # Fail before registering a version if the package cannot be opened.
    Presentation(output)


def edit_pptx_unit_texts(source: Path, output: Path, updates: list[dict]) -> list[int]:
    """Patch extracted PPT text units in one immutable copy.

    This is used by the terminology safety net.  It intentionally works from
    the structural locator captured at extraction time, never from a search of
    visible text.  Consequently a repeated label on a slide cannot be changed
    by accident.  The caller has already excluded objects a teacher edited.
    """
    with zipfile.ZipFile(source) as z:
        slide_parts = sorted((item.filename for item in z.infolist()
                              if item.filename.startswith("ppt/slides/slide") and item.filename.endswith(".xml")),
                             key=lambda name: int(re.search(r"slide(\d+)\.xml$", name).group(1)))
        roots: dict[str, etree._Element] = {}
        changed: dict[str, bytes] = {}
        changed_pages: set[int] = set()

        def part_root(part: str):
            if part not in roots:
                roots[part] = xml(z.read(part))
            return roots[part]

        for update in updates:
            payload = update.get("source_payload") or {}
            page = update.get("page") or payload.get("slide_index")
            if not isinstance(page, int) or not 1 <= page <= len(slide_parts):
                continue
            kind = update.get("unit_type") or payload.get("type")
            if kind == "notes":
                part = slide_parts[page - 1].replace("slides/slide", "notesSlides/notesSlide")
                if part not in z.namelist():
                    continue
                paragraphs = part_root(part).xpath(".//a:p", namespaces=NS)
                index = payload.get("paragraph_index")
            elif kind == "text_paragraph":
                part = slide_parts[page - 1]
                bodies = part_root(part).xpath(".//p:txBody", namespaces=NS)
                box_index, index = payload.get("text_box_index"), payload.get("paragraph_index")
                if not isinstance(box_index, int) or not 1 <= box_index <= len(bodies):
                    continue
                paragraphs = bodies[box_index - 1].xpath("./a:p", namespaces=NS)
            elif kind == "table_cell_paragraph":
                part = slide_parts[page - 1]
                tables = part_root(part).xpath(".//a:tbl", namespaces=NS)
                table_index = payload.get("table_index")
                if not isinstance(table_index, int) or not 1 <= table_index <= len(tables):
                    continue
                rows = tables[table_index - 1].xpath(".//a:tr", namespaces=NS)
                row_index, cell_index, index = payload.get("row_index"), payload.get("cell_index"), payload.get("paragraph_index")
                if not isinstance(row_index, int) or not isinstance(cell_index, int) or not 0 <= row_index < len(rows):
                    continue
                cells = rows[row_index].xpath(".//a:tc", namespaces=NS)
                if not 0 <= cell_index < len(cells):
                    continue
                paragraphs = cells[cell_index].xpath(".//a:p", namespaces=NS)
            else:
                continue
            if not isinstance(index, int) or not 0 <= index < len(paragraphs):
                continue
            replace_text(paragraphs[index], (update.get("text") or "").replace("␊", "\n").replace("␍", "\r"))
            changed[part] = serialize(roots[part])
            changed_pages.add(page)

        if not changed:
            return []
        output.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output, "x", compression=zipfile.ZIP_DEFLATED) as out:
            for item in z.infolist():
                out.writestr(item, changed.get(item.filename, z.read(item.filename)))
    # Reject a damaged package before it is recorded as a review version.
    Presentation(output)
    return sorted(changed_pages)
