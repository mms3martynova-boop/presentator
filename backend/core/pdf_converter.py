"""Конвертация PPTX → PDF через LibreOffice headless."""

from __future__ import annotations

import logging
import os
import shutil
import subprocess
from pathlib import Path

logger = logging.getLogger(__name__)

DEFAULT_TIMEOUT = float(os.getenv("PRES_LIBREOFFICE_TIMEOUT", "60"))
LIBREOFFICE_BIN = os.getenv("PRES_LIBREOFFICE_BIN", "libreoffice")


class PdfConversionError(Exception):
    pass


def is_libreoffice_available() -> bool:
    return shutil.which(LIBREOFFICE_BIN) is not None


def convert_to_pdf(pptx_path: str | Path, output_dir: str | Path) -> Path:
    pptx_path = Path(pptx_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    command = [
        LIBREOFFICE_BIN,
        "--headless",
        f"-env:UserInstallation=file://{output_dir}/lo_profile",
        "--convert-to",
        "pdf",
        "--outdir",
        str(output_dir),
        str(pptx_path),
    ]
    try:
        result = subprocess.run(command, capture_output=True, timeout=DEFAULT_TIMEOUT, check=False)
    except FileNotFoundError as exc:
        raise PdfConversionError("LibreOffice не установлен") from exc
    except subprocess.TimeoutExpired as exc:
        raise PdfConversionError(
            f"Конвертация в PDF превысила таймаут {DEFAULT_TIMEOUT} секунд"
        ) from exc
    pdf_path = output_dir / f"{pptx_path.stem}.pdf"
    if result.returncode != 0 or not pdf_path.exists():
        details = result.stderr.decode(errors="replace").strip()[:300]
        raise PdfConversionError(f"LibreOffice не смог сконвертировать файл: {details}")
    return pdf_path


def convert_to_pdf_async(pptx_path: str | Path, output_dir: str | Path, callback) -> None:
    """Запускает конвертацию и вызывает callback(pdf_path | PdfConversionError)."""
    try:
        callback(convert_to_pdf(pptx_path, output_dir))
    except PdfConversionError as exc:
        callback(exc)


def get_page_count(pdf_path: str | Path) -> int:
    try:
        from pypdf import PdfReader
    except ImportError:
        try:
            from PyPDF2 import PdfReader  # type: ignore[no-redef]
        except ImportError:
            logger.warning("pypdf/PyPDF2 не установлены — количество страниц неизвестно")
            return 0
    with Path(pdf_path).open("rb") as stream:
        return len(PdfReader(stream).pages)
