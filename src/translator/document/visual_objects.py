from __future__ import annotations

import hashlib
import zipfile
from dataclasses import dataclass
from pathlib import Path

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {
    "w": W_NS,
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}
VISUAL_XPATH = (
    ".//w:drawing|.//w:pict|.//v:*|.//o:OLEObject|.//w:object|"
    ".//m:oMath|.//m:oMathPara|.//w:txbxContent"
)


@dataclass(frozen=True)
class VisualInventory:
    counts: dict[str, int]
    part_hashes: dict[str, str]
    relationship_hashes: dict[str, str]
    anchor_hashes: tuple[str, ...]


def paragraph_has_visual_or_embedded(paragraph: etree._Element) -> bool:
    return bool(paragraph.xpath(VISUAL_XPATH, namespaces=NS))


def _sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def inspect_visuals(path: Path) -> VisualInventory:
    counts = {"drawing": 0, "pict_vml": 0, "ole_object": 0, "equation": 0, "textbox": 0}
    part_hashes: dict[str, str] = {}
    relationship_hashes: dict[str, str] = {}
    anchors: list[str] = []
    with zipfile.ZipFile(path) as zf:
        for name in zf.namelist():
            data = zf.read(name)
            if name.startswith(
                ("word/media/", "word/charts/", "word/embeddings/", "word/diagrams/")
            ):
                part_hashes[name] = _sha(data)
            if name.startswith("word/_rels/") or (name.startswith("word/") and "/_rels/" in name):
                relationship_hashes[name] = _sha(data)
            if not name.startswith("word/") or not name.endswith(".xml"):
                continue
            try:
                root = etree.fromstring(data)
            except etree.XMLSyntaxError:
                continue
            counts["drawing"] += len(root.xpath(".//w:drawing", namespaces=NS))
            counts["pict_vml"] += len(root.xpath(".//w:pict|.//v:*", namespaces=NS))
            counts["ole_object"] += len(root.xpath(".//o:OLEObject|.//w:object", namespaces=NS))
            counts["equation"] += len(root.xpath(".//m:oMath|.//m:oMathPara", namespaces=NS))
            counts["textbox"] += len(root.xpath(".//w:txbxContent", namespaces=NS))
            for node in root.xpath(
                ".//w:drawing|.//w:pict|.//w:object|.//m:oMath|.//m:oMathPara", namespaces=NS
            ):
                anchors.append(_sha(etree.tostring(node, method="c14n")))
    return VisualInventory(counts, part_hashes, relationship_hashes, tuple(anchors))
