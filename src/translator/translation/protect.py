from __future__ import annotations

import re
from collections import Counter

from translator.errors import TokenProtectionError

TOKEN_RE = re.compile(
    r"(?:[A-Z]{2,}(?:[.-][A-Z0-9]+)+)|"
    r"(?:\b(?:Pb|Cd|Cu|Zn|Fe|Pd|Mg|NaCl|HCl|HNO3|H2O2|H2SO4|H3BO3|HF|NH4H2PO4|Pd\(NO3\)2|Mg\(NO3\)2)\b)|"
    r"(?:\b[A-Z][a-z]?\d+(?:[A-Z][a-z]?\d*)*\b)|"
    r"(?:[<>≤≥]?\s*-?\d+(?:[.,]\d+)?\s*(?:%|°C|µg/L|µg/kg|mg/kg|mg/L|mL|ml|µL|g|kg|ppm|ppb|nm|h|min)?)|"
    r"(?:[A-Za-z]\d?\s*=\s*[^,;:]+)"
)


def protected_tokens(text: str) -> tuple[str, ...]:
    return tuple(match.group(0).strip() for match in TOKEN_RE.finditer(text))


def validate_tokens(source: str, target: str) -> None:
    source_tokens = Counter(protected_tokens(source))
    target_tokens = Counter(protected_tokens(target))
    # Chemical names may legitimately be reordered, but immutable literal tokens
    # found in the source must survive exactly in the approved target.
    missing = source_tokens - target_tokens
    if missing:
        raise TokenProtectionError(f"Protected tokens missing from target: {dict(missing)}")


def template_signature(text: str) -> tuple[str, tuple[str, ...]]:
    tokens: list[str] = []

    def replace(match: re.Match[str]) -> str:
        tokens.append(match.group(0).strip())
        return f"[TOKEN_{len(tokens) - 1}]"

    return TOKEN_RE.sub(replace, text), tuple(tokens)
