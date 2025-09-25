from __future__ import annotations

import os
import shutil
import subprocess
import tempfile


class GhostscriptError(RuntimeError):
    """Raised when Ghostscript conversion fails."""


def find_ghostscript() -> str:
    candidates = ["gs", "gswin64c", "gswin32c"]
    for candidate in candidates:
        path = shutil.which(candidate)
        if path:
            return path
    raise GhostscriptError(
        "Ghostscript executable not found. Install Ghostscript and ensure it is available in PATH."
    )


def convert_eps_to_pdf_bytes(eps_path: str) -> bytes:
    if not os.path.isfile(eps_path):
        raise GhostscriptError(f"EPS file not found: {eps_path}")
    ghostscript = find_ghostscript()
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_output:
        output_path = temp_output.name
    try:
        command = [
            ghostscript,
            "-dSAFER",
            "-dBATCH",
            "-dNOPAUSE",
            "-sDEVICE=pdfwrite",
            "-dEPSCrop",
            f"-sOutputFile={output_path}",
            eps_path,
        ]
        result = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )
        if result.returncode != 0:
            error = result.stderr.decode("utf-8", errors="ignore") or "Ghostscript conversion failed"
            raise GhostscriptError(error)
        with open(output_path, "rb") as pdf_file:
            return pdf_file.read()
    finally:
        try:
            os.remove(output_path)
        except OSError:
            pass
