from __future__ import annotations

import json
from pathlib import Path

from lxml import etree

from translator.formatting.apply import apply_document_formatting
from translator.formatting.profiles import load_profile

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def test_etech_profile_applies_table_rules_without_text_change(tmp_path: Path) -> None:
    config = tmp_path / "format.json"
    config.write_text(
        json.dumps(
            {
                "target_italic": True,
                "target_space_after_twips": 80,
                "heading_keep_with_next": True,
                "table_cell_margin_twips": 75,
            }
        ),
        encoding="utf-8",
    )
    root = etree.fromstring(
        f"<w:document xmlns:w='{W_NS}'><w:body><w:tbl><w:tr><w:tc>"
        "<w:p><w:r><w:t>UNCHANGED</w:t></w:r></w:p>"
        "</w:tc></w:tr></w:tbl></w:body></w:document>"
    )
    before = root.xpath(".//w:t/text()", namespaces=NS)
    stats = apply_document_formatting(root, load_profile("etech-sop", config))
    assert stats["tables_formatted"] == 1
    assert root.xpath(".//w:tcMar/w:top[@w:w='75']", namespaces=NS)
    assert root.xpath(".//w:t/text()", namespaces=NS) == before
