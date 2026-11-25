"""Read physical Word pages without changing the document or guessing image order."""
from __future__ import annotations

import json
import posixpath
import uuid
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def bookmarked_copy(source: Path, target: Path) -> dict[str, list[str]]:
    """Add zero-width location markers to a disposable copy, never the source.

    Media relationships are read directly from OOXML. Bookmarks let Word
    return their real layout pages without exporting every picture via COM.
    """
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    with zipfile.ZipFile(source) as archive:
        document = etree.fromstring(archive.read("word/document.xml"), parser)
        relationships = etree.fromstring(archive.read("word/_rels/document.xml.rels"), parser)
        images = {}
        for rel in relationships:
            if rel.get("Type", "").endswith("/image") and rel.get("TargetMode") != "External":
                target_part = posixpath.normpath(posixpath.join("word", rel.get("Target", ""))).lstrip("/")
                if target_part in archive.namelist() and target_part.startswith("word/media/"):
                    images[rel.get("Id")] = target_part
        bookmarks = {}
        prefix = "sitrain_" + uuid.uuid4().hex[:8] + "_"
        marker = max((int(e.get(f"{{{W}}}id", "0")) for e in document.iter(f"{{{W}}}bookmarkStart")), default=0) + 1
        for run in list(document.iter(f"{{{W}}}r")):
            parts = set()
            for element in run.iter():
                owner = next((ancestor for ancestor in element.iterancestors() if ancestor.tag == f"{{{W}}}r"), None)
                if owner is not run:
                    continue
                for attribute in ("embed", "id"):
                    image = images.get(element.get(f"{{{R}}}{attribute}"))
                    if image:
                        parts.add(image)
            if not parts:
                continue
            name = prefix + str(marker)
            start = etree.Element(f"{{{W}}}bookmarkStart", {f"{{{W}}}id": str(marker), f"{{{W}}}name": name})
            end = etree.Element(f"{{{W}}}bookmarkEnd", {f"{{{W}}}id": str(marker)})
            run.addprevious(start)
            run.addnext(end)
            bookmarks[name] = sorted(parts)
            marker += 1
        with zipfile.ZipFile(target, "x", zipfile.ZIP_DEFLATED) as output:
            for info in archive.infolist():
                output.writestr(info, etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)
                                if info.filename == "word/document.xml" else archive.read(info.filename))
    return bookmarks


def index_word_images(source: Path) -> dict:
    import pythoncom
    import win32com.client

    working = source.parent / f".word-page-index-{uuid.uuid4().hex}.docx"
    # Caller supplies an immutable published source. Only this explicitly
    # named temporary file is written and removed.
    bookmarks = bookmarked_copy(source, working)
    app = document = None
    previous_links = None
    result = {"status": "ready", "pages": {}, "locations": {}, "total_pages": 0}
    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        app.AutomationSecurity = 3
        previous_links = app.Options.UpdateLinksAtOpen
        app.Options.UpdateLinksAtOpen = False
        document = app.Documents.Open(str(working.resolve()), ReadOnly=True,
                                      AddToRecentFiles=False, ConfirmConversions=False)
        document.Repaginate()
        result["total_pages"] = int(document.Content.Information(4))
        for bookmark, names in bookmarks.items():
            try:
                region = document.Bookmarks.Item(bookmark).Range
                page = int(region.Information(3))
                printed = int(region.Information(1))
                if page < 1:
                    continue
                for name in names:
                    pages = result["pages"].setdefault(name, [])
                    if page not in pages:
                        pages.append(page)
                    locations = result["locations"].setdefault(name, [])
                    entry = {"page": page, "printed_page": printed, "position": int(region.Start)}
                    if entry not in locations:
                        locations.append(entry)
            except Exception:
                continue
        for pages in result["pages"].values():
            pages.sort()
        if set(result["pages"]) != {part for parts in bookmarks.values() for part in parts}:
            result["status"] = "partial"
        return result
    finally:
        try:
            if document is not None:
                document.Close(False)
        finally:
            try:
                if app is not None:
                    if previous_links is not None:
                        app.Options.UpdateLinksAtOpen = previous_links
                    app.Quit()
            finally:
                pythoncom.CoUninitialize()
                if working.exists():
                    working.unlink()


def main(source: str, target: str) -> int:
    try:
        result = index_word_images(Path(source))
        output = Path(target)
        output.parent.mkdir(parents=True, exist_ok=True)
        temporary = output.with_suffix(".tmp")
        temporary.write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
        temporary.replace(output)
        return 0
    except Exception:
        return 1
