from __future__ import annotations

import hashlib
import json
import re
from pathlib import Path

from translator.errors import ProvenanceError, TokenProtectionError
from translator.models import TMEntry
from translator.translation.protect import template_signature, validate_tokens
from translator.translation.provenance import usable_in_no_ai


def normalize_source(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def source_hash(text: str) -> str:
    return hashlib.sha256(normalize_source(text).encode("utf-8")).hexdigest()


class TranslationMemory:
    def __init__(self, entries: list[TMEntry]) -> None:
        self.entries = entries
        self._exact: dict[tuple[str, str, str], TMEntry] = {}
        self._templates: list[TMEntry] = []
        for entry in entries:
            expected = source_hash(entry.source_normalized)
            if entry.source_hash != expected:
                raise ProvenanceError(f"Source hash mismatch: {entry.provenance_id}")
            if not usable_in_no_ai(entry):
                continue
            key = (entry.source_language, entry.target_language, entry.source_normalized)
            if key in self._exact and self._exact[key].target_text != entry.target_text:
                raise ProvenanceError(
                    f"Conflicting approved entries for {entry.source_normalized!r}"
                )
            if entry.template:
                self._templates.append(entry)
            else:
                self._exact[key] = entry

    @classmethod
    def empty(cls) -> TranslationMemory:
        return cls([])

    @classmethod
    def load(cls, path: Path | None) -> TranslationMemory:
        if path is None:
            return cls.empty()
        payload = json.loads(path.read_text(encoding="utf-8"))
        raw = payload["entries"] if isinstance(payload, dict) else payload
        return cls([TMEntry(**item) for item in raw])

    def exact(self, source: str, source_lang: str, target_lang: str) -> TMEntry | None:
        return self._exact.get((source_lang, target_lang, normalize_source(source)))

    def template(
        self, source: str, source_lang: str, target_lang: str
    ) -> tuple[TMEntry, str] | None:
        signature, tokens = template_signature(normalize_source(source))
        for entry in self._templates:
            if entry.source_language != source_lang or entry.target_language != target_lang:
                continue
            approved_signature, approved_tokens = template_signature(entry.source_normalized)
            if signature != approved_signature or len(tokens) != len(approved_tokens):
                continue
            target_signature, target_tokens = template_signature(entry.target_text)
            if len(target_tokens) != len(tokens):
                raise TokenProtectionError(
                    f"Approved template token mismatch: {entry.provenance_id}"
                )
            result = target_signature
            for index, token in enumerate(tokens):
                result = result.replace(f"[TOKEN_{index}]", token)
            validate_tokens(source, result)
            return entry, result
        return None
