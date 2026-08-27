from __future__ import annotations

import re
import tempfile
from dataclasses import asdict
from pathlib import Path

from translator.config import PipelineConfig
from translator.document.extract import extract_units
from translator.document.inject import write_atomic_docx
from translator.document.normalize import normalize_to_docx
from translator.errors import IncompleteOutputNameError
from translator.formatting.profiles import load_profile
from translator.models import RuntimeEvidence
from translator.qa.privacy import NetworkGuard, assert_no_models_loaded
from translator.qa.reports import write_private_report
from translator.qa.structural import structural_qa
from translator.translation.memory import TranslationMemory
from translator.translation.no_ai import NoAITranslator

RELEASE_WORDS = re.compile(r"(?:final|release|approved)", re.I)


def _output_name(source: Path, incomplete: bool) -> str:
    suffix = "_no-ai_INCOMPLETE.docx" if incomplete else "_no-ai.docx"
    return source.stem + suffix


def run_pipeline(source: Path, config: PipelineConfig) -> tuple[Path, Path, dict[str, object]]:
    config.validate()
    evidence = RuntimeEvidence()
    config.output_dir.mkdir(parents=True, exist_ok=True)
    with NetworkGuard(evidence), tempfile.TemporaryDirectory(prefix="translator-no-ai-") as temp:
        normalized = normalize_to_docx(source.resolve(), Path(temp))
        memory = TranslationMemory.load(config.translation_memory)
        units, _elements = extract_units(normalized, config.source_language)
        resolver = NoAITranslator(memory, config.source_language, config.target_language)
        resolutions = [resolver.resolve(unit) for unit in units]
        translated = [item for item in resolutions if item.status == "translated"]
        missing = [item for item in resolutions if item.status == "manual_translation_required"]
        skipped = [item for item in resolutions if item.status == "skipped"]
        profile = load_profile(config.format_profile, config.formatting_config)
        incomplete = bool(missing)
        output = config.output_dir / _output_name(source, incomplete)
        if incomplete and RELEASE_WORDS.search(output.name):
            raise IncompleteOutputNameError(output.name)
        format_result = write_atomic_docx(
            normalized, output, units, resolutions, config.show_missing_markers, profile
        )
        qa = structural_qa(
            normalized, output, [item.target_text for item in translated if item.target_text]
        )
        evidence.models_loaded = assert_no_models_loaded()
        report: dict[str, object] = {
            "status": "CONDITIONAL_PASS" if incomplete else "PASS",
            "execution_profile": "no-ai",
            "source_language": config.source_language,
            "target_language": config.target_language,
            "format_profile": config.format_profile,
            "runtime_ai_used": evidence.runtime_ai_used,
            "external_translation_calls": evidence.external_translation_calls,
            "outbound_document_content_calls": evidence.outbound_document_content_calls,
            "models_loaded": evidence.models_loaded,
            "blocked_network_attempts": evidence.blocked_network_attempts,
            "total_units": len(units),
            "translated_units": len(translated),
            "manual_translation_required": len(missing),
            "visual_units_skipped": len(skipped),
            "translation_coverage": (len(translated) / (len(translated) + len(missing)))
            if translated or missing
            else 1.0,
            **format_result,
            "output_marked_incomplete": incomplete,
            "visual_content_policy": "PRESERVED_UNCHANGED_AND_SKIPPED",
            "visual_content_translated": False,
            "ocr_used": False,
            "structural_qa": qa,
            "resolutions": [asdict(item) for item in resolutions],
        }
        report_path = config.output_dir / f"{source.stem}_no-ai_report.json"
        write_private_report(report_path, report)
        return output, report_path, report
