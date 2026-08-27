from __future__ import annotations

import zipfile
from dataclasses import asdict
from pathlib import Path

from lxml import etree

from translator.document.extract import paragraph_text
from translator.document.visual_objects import inspect_visuals
from translator.errors import StructuralQAError

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def _root(path: Path) -> etree._Element:
    with zipfile.ZipFile(path) as zf:
        return etree.fromstring(zf.read("word/document.xml"))


def structural_qa(source: Path, output: Path, expected_targets: list[str]) -> dict[str, object]:
    with zipfile.ZipFile(output) as zf:
        corrupt = zf.testzip()
    if corrupt:
        raise StructuralQAError(f"Corrupt ZIP member: {corrupt}")
    before, after = _root(source), _root(output)
    before_text = [
        paragraph_text(p) for p in before.xpath(".//w:p", namespaces=NS) if paragraph_text(p)
    ]
    after_text = [
        paragraph_text(p) for p in after.xpath(".//w:p", namespaces=NS) if paragraph_text(p)
    ]
    cursor = 0
    for text in before_text:
        try:
            cursor = after_text.index(text, cursor) + 1
        except ValueError as error:
            raise StructuralQAError(f"Source text lost or reordered: {text[:80]}") from error
    for target in expected_targets:
        if after_text.count(target) != 1:
            raise StructuralQAError(f"Approved target count is not one: {target[:80]}")
    if any("[TOKEN_" in item for item in after_text):
        raise StructuralQAError("Internal token placeholder leaked into output")
    before_tables = (
        len(before.xpath(".//w:tbl", namespaces=NS)),
        len(before.xpath(".//w:tr", namespaces=NS)),
        len(before.xpath(".//w:tc", namespaces=NS)),
    )
    after_tables = (
        len(after.xpath(".//w:tbl", namespaces=NS)),
        len(after.xpath(".//w:tr", namespaces=NS)),
        len(after.xpath(".//w:tc", namespaces=NS)),
    )
    if after_tables != before_tables:
        raise StructuralQAError(f"Table/row/cell counts changed: {before_tables} -> {after_tables}")
    visuals_before, visuals_after = inspect_visuals(source), inspect_visuals(output)
    if visuals_before != visuals_after:
        raise StructuralQAError("Visual, embedded, relationship, or anchor inventory changed")
    return {
        "status": "PASS",
        "zip_integrity": "PASS",
        "source_text_preserved": True,
        "table_row_cell_counts_before": before_tables,
        "table_row_cell_counts_after": after_tables,
        "visuals_before": asdict(visuals_before),
        "visuals_after": asdict(visuals_after),
        "media_hash_result": "PASS",
        "relationship_result": "PASS",
    }
