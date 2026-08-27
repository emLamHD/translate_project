"""Stable public exceptions."""


class TranslatorError(Exception):
    """Base error for deterministic translator failures."""


class ConfigurationError(TranslatorError):
    """Configuration is unsafe or invalid."""


class ProvenanceError(TranslatorError):
    """Translation provenance is invalid or conflicts."""


class TokenProtectionError(TranslatorError):
    """Protected numbers, units, codes, or formulae do not match."""


class StructuralQAError(TranslatorError):
    """DOCX structure changed outside the allowlist."""


class IncompleteOutputNameError(TranslatorError):
    """An incomplete result was given a release-like filename."""


class NetworkAccessBlocked(TranslatorError):
    """Outbound network access was attempted in no-AI mode."""
