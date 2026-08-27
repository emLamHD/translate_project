from __future__ import annotations

import re
import zipfile
from pathlib import Path

from lxml import etree

from translator.document.visual_objects import paragraph_has_visual_or_embedded
from translator.models import TextUnit
from translator.translation.memory import normalize_source, source_hash

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = f"{{{NS['w']}}}"
VI_MARKS = re.compile(r"[À-ỹĐđ]")
WORD = re.compile(r"[A-Za-zÀ-ỹĐđ]{2,}")
NUMBERISH = re.compile(r"^[\s\d.,:;()%+\-–—≤≥=×/*µ°²³\[\]]+$")


def paragraph_text(paragraph: etree._Element) -> str:
    return "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))


def deterministic_language_route(text: str, language: str) -> bool:
    text = normalize_source(text)
    if language != "vi" or not text or NUMBERISH.fullmatch(text) or not WORD.search(text):
        return False
    hints = ("MUC ", "PHAM VI", "THIET BI", "HOA CHAT", "DUNG DICH", "MAU ", "NGUOI ")
    return bool(VI_MARKS.search(text) or any(item in text.upper() for item in hints))


def load_document_root(path: Path) -> etree._Element:
    with zipfile.ZipFile(path) as zf:
        return etree.fromstring(zf.read("word/document.xml"))


def extract_units(
    path: Path, source_language: str
) -> tuple[list[TextUnit], dict[str, etree._Element]]:
    root = load_document_root(path)
    units: list[TextUnit] = []
    elements: dict[str, etree._Element] = {}
    for index, paragraph in enumerate(root.xpath(".//w:p", namespaces=NS)):
        text = normalize_source(paragraph_text(paragraph))
        visual = paragraph_has_visual_or_embedded(paragraph)
        if not text:
            continue
        eligible = deterministic_language_route(text, source_language)
        if not eligible and not visual:
            continue
        unit_id = f"document.xml:p{index}:{source_hash(text)[:12]}"
        reason = "skip_visual_or_embedded_object" if visual else None
        unit = TextUnit(unit_id, "word/document.xml", text, text, source_hash(text), visual, reason)
        units.append(unit)
        elements[unit_id] = paragraph
    return units, elements
