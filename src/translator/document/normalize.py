from __future__ import annotations

import subprocess
from pathlib import Path


def normalize_to_docx(source: Path, work_dir: Path) -> Path:
    if source.suffix.lower() == ".docx":
        return source
    if source.suffix.lower() != ".doc":
        raise ValueError("Only DOC and DOCX inputs are supported")
    work_dir.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [
            "/usr/bin/soffice",
            "--headless",
            "--convert-to",
            "docx",
            "--outdir",
            str(work_dir),
            str(source),
        ],
        check=True,
        capture_output=True,
        text=True,
    )
    result = work_dir / f"{source.stem}.docx"
    if not result.exists():
        raise RuntimeError(f"LibreOffice did not create {result}")
    return result
