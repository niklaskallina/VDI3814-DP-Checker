"""Dateien einlesen: PDF (auch mehrseitig/gescannt), PNG, JPEG -> Seitenbilder."""

from __future__ import annotations

import hashlib
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator

from PIL import Image

from ..config import SETTINGS

SUPPORTED_SUFFIXES = {".pdf", ".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"}


@dataclass
class PageImage:
    """Eine gerasterte Seite inkl. Herkunft."""

    source_path: Path
    page_index: int          # 0-basiert
    image: Image.Image
    page_count: int = 1

    @property
    def file_name(self) -> str:
        return self.source_path.name

    @property
    def stem(self) -> str:
        return self.source_path.stem


def file_hash(path: Path) -> str:
    """SHA-256 der Quelldatei - Grundlage fuer die Duplikatserkennung beim Import."""
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _load_pdf(path: Path, dpi: int) -> Iterator[PageImage]:
    import fitz  # PyMuPDF

    # PyMuPDF statt pdf2image: keine externe Poppler-Installation noetig,
    # das vereinfacht das Setup auf Windows-Rechnern erheblich.
    with fitz.open(path) as doc:
        zoom = dpi / 72.0
        matrix = fitz.Matrix(zoom, zoom)
        for index, page in enumerate(doc):
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            yield PageImage(path, index, image, page_count=doc.page_count)


def _load_image(path: Path) -> Iterator[PageImage]:
    with Image.open(path) as img:
        frames = getattr(img, "n_frames", 1)
        for index in range(frames):
            img.seek(index)
            yield PageImage(path, index, img.convert("RGB"), page_count=frames)


def load_pages(path: str | Path, dpi: int | None = None) -> Iterator[PageImage]:
    """Liefert alle Seiten einer Datei als RGB-Bild.

    Wirft ValueError bei nicht unterstuetzten Dateitypen; Lesefehler einzelner
    Dateien werden bewusst nicht abgefangen, sondern von der Pipeline
    behandelt (dort landet die Datei mit Fehlermeldung im Protokoll).
    """
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_SUFFIXES:
        raise ValueError(f"Nicht unterstuetzter Dateityp: {path.name} ({suffix})")
    if suffix == ".pdf":
        yield from _load_pdf(path, dpi or SETTINGS.render_dpi)
    else:
        yield from _load_image(path)


def render_pdf_page(path: str | Path, page_index: int, dpi: int | None = None) -> Image.Image:
    """Rendert genau eine PDF-Seite.

    Bewusst einzeln: bei mehrseitigen Scans wuerde ein Durchlauf durch alle
    Seiten je Seite die Verarbeitung quadratisch verlangsamen.
    """
    import fitz  # PyMuPDF

    zoom = (dpi or SETTINGS.render_dpi) / 72.0
    with fitz.open(path) as doc:
        page = doc[page_index]
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def collect_files(paths: list[str | Path]) -> list[Path]:
    """Sammelt aus Dateien/Ordnern alle unterstuetzten Dateien (Batch-Import)."""
    files: list[Path] = []
    for raw in paths:
        path = Path(raw)
        if path.is_dir():
            for child in sorted(path.rglob("*")):
                if child.is_file() and child.suffix.lower() in SUPPORTED_SUFFIXES:
                    files.append(child)
        elif path.is_file() and path.suffix.lower() in SUPPORTED_SUFFIXES:
            files.append(path)
    return files
