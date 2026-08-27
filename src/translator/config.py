from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Literal

ExecutionProfile = Literal["no-ai"]
FormatProfile = Literal["preserve", "clean", "etech-sop"]


@dataclass(frozen=True)
class PipelineConfig:
    execution_profile: ExecutionProfile
    source_language: str
    target_language: str
    format_profile: FormatProfile
    translation_memory: Path | None
    output_dir: Path
    show_missing_markers: bool = False
    formatting_config: Path | None = None

    def validate(self) -> None:
        if self.execution_profile != "no-ai":
            raise ValueError("Only --execution-profile no-ai is supported")
        if self.source_language == "auto":
            raise ValueError("The production blind run requires an explicit source language")
        if self.format_profile == "etech-sop" and self.formatting_config is None:
            raise ValueError("etech-sop requires --format-config")
