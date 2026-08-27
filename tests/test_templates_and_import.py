from __future__ import annotations

import argparse
import csv
from pathlib import Path

import pytest

from translator.cli import import_tm
from translator.errors import ProvenanceError, TokenProtectionError
from translator.models import TextUnit, TMEntry
from translator.translation.memory import TranslationMemory, source_hash
from translator.translation.no_ai import NoAITranslator
from translator.translation.protect import validate_tokens


def test_approved_normalized_template() -> None:
    source = "Cân chính xác 5 g mẫu."
    entry = TMEntry(
        "vi",
        "en",
        source,
        "Accurately weigh 5 g of the sample.",
        source_hash(source),
        "owner_manual",
        "template:1",
        "Owner",
        "now",
        True,
        template=True,
    )
    candidate = "Cân chính xác 10 g mẫu."
    unit = TextUnit("u", "doc", candidate, candidate, source_hash(candidate), False)
    result = NoAITranslator(TranslationMemory([entry]), "vi", "en").resolve(unit)
    assert result.status == "translated"
    assert "10 g" in (result.target_text or "")


def test_token_mismatch_fails() -> None:
    with pytest.raises(TokenProtectionError):
        validate_tokens("Cân 5 g mẫu DOC.TEST-01.", "Weigh the sample.")


def test_fuzzy_match_not_used(approved_entry: TMEntry) -> None:
    text = "Cân cẩn thận 5 g mẫu."
    item = TextUnit("u", "doc", text, text, source_hash(text), False)
    assert (
        NoAITranslator(TranslationMemory([approved_entry]), "vi", "en").resolve(item).status
        == "manual_translation_required"
    )


def test_tm_import_conflict(tmp_path: Path) -> None:
    csv_path = tmp_path / "approved.csv"
    with csv_path.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(
            stream, fieldnames=["unit_id", "source_text", "target_text", "approved_by"]
        )
        writer.writeheader()
        writer.writerow(
            {
                "unit_id": "1",
                "source_text": "Cân 5 g mẫu.",
                "target_text": "Weigh 5 g of sample.",
                "approved_by": "Owner",
            }
        )
        writer.writerow(
            {
                "unit_id": "2",
                "source_text": "Cân 5 g mẫu.",
                "target_text": "Accurately weigh 5 g.",
                "approved_by": "Owner",
            }
        )
    args = argparse.Namespace(
        input=str(csv_path),
        output=str(tmp_path / "tm.json"),
        audit_report=str(tmp_path / "audit.json"),
        source="vi",
        target="en",
        approved_by="Owner",
    )
    with pytest.raises(ProvenanceError):
        import_tm(args)
