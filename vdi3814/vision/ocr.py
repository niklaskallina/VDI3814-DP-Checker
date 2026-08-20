"""Ergaenzendes lokales OCR (Tesseract).

Zweck: Textgenauigkeit dort, wo das Vision-Modell unscharf ist - vor allem
fuer die Fassungserkennung (Stichworte im Kopfbereich) und als Gegenprobe
fuer Datenpunktbezeichnungen. Fehlt Tesseract, laeuft das Tool ohne OCR
weiter (nur mit entsprechendem Hinweis).
"""

from __future__ import annotations

import functools
import logging
import os

from PIL import Image

from ..config import SETTINGS, bundled_tesseract
from ..ingest.preprocess import prepare_for_ocr

log = logging.getLogger(__name__)


@functools.lru_cache(maxsize=1)
def available() -> bool:
    """Prueft die Texterkennung und richtet sie bei Bedarf selbst ein.

    Reihenfolge: ausdruecklich gesetzter Pfad -> mitgeliefertes Tesseract
    (neben der EXE) -> im System installiertes Tesseract.
    """
    try:
        import pytesseract
    except ImportError:
        return False

    if SETTINGS.tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = SETTINGS.tesseract_cmd
    else:
        mitgeliefert = bundled_tesseract()
        if mitgeliefert is not None:
            pytesseract.pytesseract.tesseract_cmd = str(mitgeliefert)
            sprachdaten = mitgeliefert.parent / "tessdata"
            if sprachdaten.is_dir():
                os.environ.setdefault("TESSDATA_PREFIX", str(sprachdaten))

    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception as exc:      # pragma: no cover - haengt von der Installation ab
        log.info("Tesseract nicht verfuegbar: %s", exc)
        return False


def sprachen() -> list[str]:
    """Installierte OCR-Sprachpakete."""
    if not available():
        return []
    import pytesseract

    try:
        return list(pytesseract.get_languages(config=""))
    except Exception:             # pragma: no cover
        return []


def page_text(image: Image.Image, lang: str | None = None) -> str:
    """Volltext einer Seite (leerer String, wenn OCR nicht verfuegbar ist)."""
    if not available():
        return ""
    import pytesseract

    try:
        return pytesseract.image_to_string(prepare_for_ocr(image), lang=lang or SETTINGS.ocr_lang)
    except Exception as exc:      # pragma: no cover
        log.warning("OCR fehlgeschlagen: %s", exc)
        return ""
