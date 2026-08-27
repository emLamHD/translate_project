from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class FormattingProfile:
    name: str
    target_italic: bool = True
    target_space_after_twips: int = 80
    heading_keep_with_next: bool = True
    table_cell_margin_twips: int | None = None


def load_profile(name: str, config: Path | None = None) -> FormattingProfile:
    if name == "preserve":
        return FormattingProfile(
            "preserve", target_space_after_twips=0, heading_keep_with_next=False
        )
    if name == "clean":
        # Source table metrics are layout-sensitive. The generic clean profile
        # does not alter them; an Owner-approved etech-sop profile may do so.
        return FormattingProfile("clean", table_cell_margin_twips=None)
    if name != "etech-sop" or config is None:
        raise ValueError(f"Unknown or incomplete format profile: {name}")
    payload = json.loads(config.read_text(encoding="utf-8"))
    return FormattingProfile(name="etech-sop", **payload)
