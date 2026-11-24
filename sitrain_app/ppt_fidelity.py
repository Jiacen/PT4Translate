"""Fail reconstruction if a translation changes protected PPT package structure."""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

from lxml import etree

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}
TEXT_PART = re.compile(r"ppt/(slides|notesSlides|slideMasters|slideLayouts|notesMasters|handoutMasters|diagrams|charts|comments)/[^/]+\.xml$")


def paragraph_line_groups(payload: dict) -> list[int]:
    """Keep a sentence together across soft breaks when its run style is equal."""
    lines = str(payload.get("value", "")).replace("␊", "\n").replace("␍", "\n").split("\n")
    lengths, styles = payload.get("run_lengths", []), payload.get("run_styles", [])
    if len(lines) < 2 or not lengths or len(lengths) != len(styles):
        return [1] * len(lines)
    owners = [index for index, length in enumerate(lengths) for _ in range(length)]
    position, groups = 0, [1]
    for line in lines[:-1]:
        position += len(line)
        if 0 < position < len(owners) and styles[owners[position - 1]] == styles[owners[position]]:
            groups[-1] += 1
        else:
            groups.append(1)
    return groups


def restore_soft_breaks(source_lines: list[str], translated: str) -> str:
    """Rewrap one uniformly styled sentence without changing its wording.

    A heading/subheading style boundary is translated separately. Inside a
    uniformly styled sentence, restoring a line break must not force the model
    to translate a trailing fragment such as 'of the demo case' independently.
    """
    flat = re.sub(r"[\r\n]+", " ", translated)
    lengths = [len(line) for line in source_lines]
    total = max(1, sum(lengths))
    cuts, consumed = [0], 0
    for length in lengths[:-1]:
        consumed += length
        proposed = max(cuts[-1], min(len(flat), round(len(flat) * consumed / total)))
        # Prefer punctuation/word boundaries; never split an ASCII model name.
        radius = max(3, len(flat) // 10)
        candidates = [i for i in range(max(cuts[-1], proposed-radius), min(len(flat), proposed+radius)+1)
                      if i == 0 or i == len(flat) or flat[i-1] in '，；：、,;: ' or flat[i].isspace()]
        if candidates:
            proposed = min(candidates, key=lambda i: abs(i-proposed))
        elif 0 < proposed < len(flat) and re.match(r'[A-Za-z0-9_-]', flat[proposed-1]) and re.match(r'[A-Za-z0-9_-]', flat[proposed]):
            while proposed < len(flat) and re.match(r'[A-Za-z0-9_-]', flat[proposed]):
                proposed += 1
        cuts.append(proposed)
    cuts.append(len(flat))
    return '\n'.join(flat[a:b].strip() for a, b in zip(cuts, cuts[1:]))


def _structure(content: bytes) -> bytes:
    root = etree.fromstring(content, etree.XMLParser(resolve_entities=False, load_dtd=False, no_network=True))
    for node in root.xpath(".//a:t|.//p:text|.//c:strCache//c:v|.//c:strLit//c:v", namespaces=NS):
        node.text = ""
    for node in root.xpath(".//p:cNvPr", namespaces=NS):
        for attribute in ("descr", "title"):
            node.attrib.pop(attribute, None)
    return etree.tostring(root, method="c14n")


def verify_pptx_preservation(original: Path, translated: Path) -> dict:
    """Media, links, themes, shapes, bullets and run styles are immutable here.

    Only translated text and accessibility descriptions may differ. This check
    runs before a machine version is registered, so a broken package cannot be
    reported as a successful translation.
    """
    protected = visual = 0
    with zipfile.ZipFile(original) as before, zipfile.ZipFile(translated) as after:
        if set(before.namelist()) != set(after.namelist()) or len(after.namelist()) != len(set(after.namelist())):
            raise RuntimeError("PPT 格式保真检查失败：文件内部资源发生增减，原件已保留。")
        for name in before.namelist():
            source, result = before.read(name), after.read(name)
            if TEXT_PART.fullmatch(name):
                visual += 1
                same = source == result or _structure(source) == _structure(result)
            else:
                protected += 1
                same = source == result
            if not same:
                raise RuntimeError(f"PPT 格式保真检查失败：{name} 的非文字内容发生变化，原件已保留。")
    return {"policy": "preserve_original", "protected_parts": protected, "visual_parts": visual}
