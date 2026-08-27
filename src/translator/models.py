from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any, Literal

ProvenanceType = Literal[
    "human_approved",
    "owner_manual",
    "claude_silver_reference",
    "google_machine_draft",
    "unknown",
]


@dataclass(frozen=True)
class TMEntry:
    source_language: str
    target_language: str
    source_normalized: str
    target_text: str
    source_hash: str
    provenance_type: ProvenanceType
    provenance_id: str
    approved_by: str | None
    approved_at: str | None
    approved: bool
    version: int = 1
    template: bool = False

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class TextUnit:
    unit_id: str
    part_name: str
    source_text: str
    source_normalized: str
    source_hash: str
    visual_or_embedded: bool
    reason: str | None = None


@dataclass(frozen=True)
class Resolution:
    unit_id: str
    status: Literal["translated", "manual_translation_required", "skipped"]
    target_text: str | None
    provenance_id: str | None
    resolution_type: str


@dataclass
class RuntimeEvidence:
    runtime_ai_used: bool = False
    external_translation_calls: int = 0
    outbound_document_content_calls: int = 0
    models_loaded: list[str] = field(default_factory=list)
    blocked_network_attempts: int = 0
