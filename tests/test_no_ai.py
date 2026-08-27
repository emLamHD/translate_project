from __future__ import annotations

import socket
import sys

import pytest

from translator.errors import NetworkAccessBlocked
from translator.models import RuntimeEvidence, TextUnit, TMEntry
from translator.qa.privacy import NetworkGuard, assert_no_models_loaded
from translator.translation.memory import TranslationMemory, source_hash
from translator.translation.no_ai import MANUAL_TRANSLATION_REQUIRED, NoAITranslator


def unit(text: str = "Cân chính xác 5 g mẫu.") -> TextUnit:
    return TextUnit("u1", "word/document.xml", text, text, source_hash(text), False)


def test_exact_approved_match(approved_entry: TMEntry) -> None:
    result = NoAITranslator(TranslationMemory([approved_entry]), "vi", "en").resolve(unit())
    assert result.status == "translated"
    assert result.resolution_type == "approved_exact"


@pytest.mark.parametrize(
    "provenance", ["claude_silver_reference", "google_machine_draft", "unknown"]
)
def test_unapproved_provenance_excluded(approved_entry: TMEntry, provenance: str) -> None:
    entry = TMEntry(
        **{**approved_entry.to_dict(), "provenance_type": provenance, "approved": False}
    )
    result = NoAITranslator(TranslationMemory([entry]), "vi", "en").resolve(unit())
    assert result.status == "manual_translation_required"


def test_approved_flag_required(approved_entry: TMEntry) -> None:
    entry = TMEntry(**{**approved_entry.to_dict(), "approved": False})
    assert (
        NoAITranslator(TranslationMemory([entry]), "vi", "en").resolve(unit()).status
        == "manual_translation_required"
    )


def test_unmatched_returns_manual() -> None:
    result = NoAITranslator(TranslationMemory.empty(), "vi", "en").resolve(unit())
    assert result.resolution_type == MANUAL_TRANSLATION_REQUIRED


def test_source_equals_target_rejected(approved_entry: TMEntry) -> None:
    entry = TMEntry(**{**approved_entry.to_dict(), "target_text": approved_entry.source_normalized})
    assert (
        NoAITranslator(TranslationMemory([entry]), "vi", "en").resolve(unit()).status
        == "manual_translation_required"
    )


def test_visual_unit_skipped() -> None:
    item = TextUnit("u", "document", "Nhãn hình", "Nhãn hình", source_hash("Nhãn hình"), True)
    assert NoAITranslator(TranslationMemory.empty(), "vi", "en").resolve(item).status == "skipped"


def test_runtime_network_guard() -> None:
    evidence = RuntimeEvidence()
    with NetworkGuard(evidence), pytest.raises(NetworkAccessBlocked):
        socket.create_connection(("example.com", 443))
    assert evidence.blocked_network_attempts == 1


@pytest.mark.parametrize("module", ["anthropic", "openai", "transformers"])
def test_no_ai_mode_rejects_loaded_ai_module(monkeypatch: pytest.MonkeyPatch, module: str) -> None:
    monkeypatch.setitem(sys.modules, module, object())
    with pytest.raises(RuntimeError):
        assert_no_models_loaded()
