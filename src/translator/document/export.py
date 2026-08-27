from __future__ import annotations

import subprocess
from pathlib import Path


def render_pdf(docx: Path, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [
            "/usr/bin/soffice",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(output_dir),
            str(docx),
        ],
        check=True,
        capture_output=True,
        text=True,
    )
    result = output_dir / f"{docx.stem}.pdf"
    if not result.exists():
        raise RuntimeError(f"LibreOffice did not render {result}")
    return result
