from __future__ import annotations

import zipfile
from pathlib import Path

from conftest import write_tm
from lxml import etree

from translator.config import PipelineConfig
from translator.document.visual_objects import inspect_visuals
from translator.pipeline.orchestrator import run_pipeline

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def root(path: Path) -> etree._Element:
    with zipfile.ZipFile(path) as zf:
        return etree.fromstring(zf.read("word/document.xml"))


def test_synthetic_no_ai_pipeline(synthetic_docx: Path, approved_entry, tmp_path: Path) -> None:
    tm = write_tm(tmp_path / "approved_tm.json", [approved_entry])
    out = tmp_path / "private_output"
    config = PipelineConfig("no-ai", "vi", "en", "clean", tm, out)
    output, report_path, report = run_pipeline(synthetic_docx, config)
    assert output.name.endswith("_no-ai_INCOMPLETE.docx")
    assert report["runtime_ai_used"] is False
    assert report["external_translation_calls"] == 0
    assert report["outbound_document_content_calls"] == 0
    assert report["models_loaded"] == []
    assert report["translated_units"] == 1
    assert report["manual_translation_required"] >= 1
    assert report["visual_units_skipped"] >= 2
    assert report["structural_qa"]["status"] == "PASS"
    assert inspect_visuals(synthetic_docx) == inspect_visuals(output)
    before, after = root(synthetic_docx), root(output)
    before_num = before.xpath(
        ".//w:p[w:pPr/w:pStyle[@w:val='ListNumber']]/w:pPr/w:numPr", namespaces=NS
    )
    after_num = after.xpath(
        ".//w:p[w:pPr/w:pStyle[@w:val='ListNumber']]/w:pPr/w:numPr", namespaces=NS
    )
    assert len(after_num) >= len(before_num)
    source_list_before = next(
        paragraph
        for paragraph in before.xpath(".//w:p", namespaces=NS)
        if "Mục đánh số" in "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))
    )
    source_list_after = next(
        paragraph
        for paragraph in after.xpath(".//w:p", namespaces=NS)
        if "Mục đánh số" in "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))
    )
    assert etree.tostring(source_list_before.find("w:pPr", NS), method="c14n") == etree.tostring(
        source_list_after.find("w:pPr", NS), method="c14n"
    )
    targets = [
        p
        for p in after.xpath(".//w:p", namespaces=NS)
        if "Accurately weigh" in "".join(p.xpath(".//w:t/text()", namespaces=NS))
    ]
    assert len(targets) == 1
    assert targets[0].xpath(".//w:rPr/w:i", namespaces=NS)
    assert targets[0].xpath("./w:pPr/w:numPr/w:numId[@w:val='0']", namespaces=NS)
    heading = next(
        paragraph
        for paragraph in after.xpath(".//w:p", namespaces=NS)
        if "PHẠM VI" in "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))
    )
    assert heading.xpath("./w:pPr/w:keepNext", namespaces=NS)
    assert not list(output.parent.glob("*.tmp"))
    assert report_path.exists()


def test_formatter_does_not_change_source_text(synthetic_docx: Path, tmp_path: Path) -> None:
    config = PipelineConfig("no-ai", "vi", "en", "clean", None, tmp_path / "out")
    output, _report_path, _report = run_pipeline(synthetic_docx, config)
    source_text = root(synthetic_docx).xpath(".//w:t/text()", namespaces=NS)
    output_text = root(output).xpath(".//w:t/text()", namespaces=NS)
    cursor = 0
    for text in source_text:
        cursor = output_text.index(text, cursor) + 1


def test_no_ocr_or_network_translation_code_exists() -> None:
    all_code = "\n".join(
        path.read_text(encoding="utf-8") for path in Path("src/translator").rglob("*.py")
    )
    assert "pytesseract.image_to_string" not in all_code
    assert "translate.googleapis.com" not in all_code
    assert "urllib.request" not in all_code
