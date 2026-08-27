from __future__ import annotations

import socket
from contextlib import AbstractContextManager
from types import TracebackType
from typing import Any

from translator.errors import NetworkAccessBlocked
from translator.models import RuntimeEvidence


class NetworkGuard(AbstractContextManager["NetworkGuard"]):
    """Process-local outbound socket guard used for the complete no-AI run."""

    def __init__(self, evidence: RuntimeEvidence) -> None:
        self.evidence = evidence
        self._original_connect: Any = None
        self._original_create_connection: Any = None

    def __enter__(self) -> NetworkGuard:
        self._original_connect = socket.socket.connect
        self._original_create_connection = socket.create_connection
        evidence = self.evidence

        def blocked_connect(*_args: Any, **_kwargs: Any) -> None:
            evidence.blocked_network_attempts += 1
            raise NetworkAccessBlocked("Outbound network is disabled in no-AI mode")

        socket.socket.connect = blocked_connect  # type: ignore[method-assign]
        socket.create_connection = blocked_connect  # type: ignore[assignment]
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        socket.socket.connect = self._original_connect  # type: ignore[method-assign]
        socket.create_connection = self._original_create_connection


BANNED_RUNTIME_IMPORTS = (
    "anthropic",
    "openai",
    "transformers",
    "torch",
    "sentence_transformers",
    "googletrans",
    "deep_translator",
    "argostranslate",
    "easyocr",
    "pytesseract",
)


def assert_no_models_loaded() -> list[str]:
    import sys

    loaded = [name for name in BANNED_RUNTIME_IMPORTS if name in sys.modules]
    if loaded:
        raise RuntimeError(f"Forbidden AI/ML modules loaded: {loaded}")
    return loaded
