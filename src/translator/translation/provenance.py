from __future__ import annotations

from translator.models import TMEntry

ALLOWED_NO_AI_PROVENANCE = {"human_approved", "owner_manual"}


def usable_in_no_ai(entry: TMEntry) -> bool:
    return entry.approved and entry.provenance_type in ALLOWED_NO_AI_PROVENANCE
