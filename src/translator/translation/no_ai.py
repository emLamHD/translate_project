from __future__ import annotations

from translator.models import Resolution, TextUnit
from translator.translation.memory import TranslationMemory
from translator.translation.protect import validate_tokens

MANUAL_TRANSLATION_REQUIRED = "MANUAL_TRANSLATION_REQUIRED"


class NoAITranslator:
    def __init__(
        self, memory: TranslationMemory, source_language: str, target_language: str
    ) -> None:
        self.memory = memory
        self.source_language = source_language
        self.target_language = target_language

    def resolve(self, unit: TextUnit) -> Resolution:
        if unit.visual_or_embedded:
            return Resolution(unit.unit_id, "skipped", None, None, "skip_visual_or_embedded_object")
        exact = self.memory.exact(unit.source_text, self.source_language, self.target_language)
        if exact is not None:
            if exact.target_text.strip() == unit.source_text.strip():
                return Resolution(
                    unit.unit_id,
                    "manual_translation_required",
                    None,
                    None,
                    "source_equals_target_rejected",
                )
            validate_tokens(unit.source_text, exact.target_text)
            return Resolution(
                unit.unit_id, "translated", exact.target_text, exact.provenance_id, "approved_exact"
            )
        template = self.memory.template(
            unit.source_text, self.source_language, self.target_language
        )
        if template is not None:
            entry, target = template
            return Resolution(
                unit.unit_id, "translated", target, entry.provenance_id, "approved_template"
            )
        return Resolution(
            unit.unit_id, "manual_translation_required", None, None, MANUAL_TRANSLATION_REQUIRED
        )
