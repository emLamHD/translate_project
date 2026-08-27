from __future__ import annotations

import os
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree

from translator.formatting.apply import apply_document_formatting, format_target_paragraph
from translator.formatting.profiles import FormattingProfile
from translator.models import Resolution, TextUnit

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = f"{{{NS['w']}}}"


def _replace_text(paragraph: etree._Element, replacement: str) -> None:
    nodes = paragraph.xpath(".//w:t", namespaces=NS)
    if not nodes:
        run = etree.SubElement(paragraph, W + "r")
        nodes = [etree.SubElement(run, W + "t")]
    nodes[0].text = replacement
    nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for node in nodes[1:]:
        node.text = ""


def _target_properties(paragraph: etree._Element) -> None:
    ppr = paragraph.find(W + "pPr")
    if ppr is None:
        ppr = etree.Element(W + "pPr")
        paragraph.insert(0, ppr)
    numpr = ppr.find(W + "numPr")
    if numpr is None:
        numpr = etree.SubElement(ppr, W + "numPr")
    for child in list(numpr):
        numpr.remove(child)
    numid = etree.SubElement(numpr, W + "numId")
    numid.set(W + "val", "0")
    for run in paragraph.xpath(".//w:r", namespaces=NS):
        rpr = run.find(W + "rPr")
        if rpr is None:
            rpr = etree.Element(W + "rPr")
            run.insert(0, rpr)
        for name in ("i", "iCs"):
            node = rpr.find(W + name)
            if node is None:
                node = etree.SubElement(rpr, W + name)
            node.set(W + "val", "1")


def write_atomic_docx(
    source: Path,
    destination: Path,
    units: list[TextUnit],
    resolutions: list[Resolution],
    show_missing_markers: bool,
    profile: FormattingProfile,
) -> dict[str, int]:
    with zipfile.ZipFile(source) as zin:
        members = {name: zin.read(name) for name in zin.namelist()}
    root = etree.fromstring(members["word/document.xml"])
    paragraphs = root.xpath(".//w:p", namespaces=NS)
    by_id = {resolution.unit_id: resolution for resolution in resolutions}
    inserted = 0
    for unit in units:
        resolution = by_id[unit.unit_id]
        if resolution.status == "skipped":
            continue
        if resolution.status == "manual_translation_required" and not show_missing_markers:
            continue
        target = resolution.target_text or "MANUAL_TRANSLATION_REQUIRED"
        index = int(unit.unit_id.split(":p", 1)[1].split(":", 1)[0])
        source_paragraph = paragraphs[index]
        clone = deepcopy(source_paragraph)
        _replace_text(clone, target)
        _target_properties(clone)
        format_target_paragraph(clone, profile)
        source_paragraph.addnext(clone)
        inserted += 1
    format_stats = apply_document_formatting(root, profile)
    members["word/document.xml"] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    destination.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(
        prefix=destination.name + ".", suffix=".tmp", dir=destination.parent
    )
    os.close(fd)
    temp_path = Path(temporary)
    try:
        with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in sorted(members):
                zout.writestr(name, members[name])
        os.replace(temp_path, destination)
    finally:
        if temp_path.exists():
            temp_path.unlink()
    return {"inserted_targets": inserted, **format_stats}
