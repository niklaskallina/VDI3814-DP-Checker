"""Extraktion aus PDFs mit Textebene (aus Excel/CAD gedruckte Funktionslisten).

Die meisten GA-Funktionslisten liegen als PDF-Druck einer Excel-Vorlage vor.
Dann sind alle Zellwerte als Text mit exakten Koordinaten vorhanden - eine
Bilderkennung waere hier unnoetig ungenau.

Dieses Modul liefert nur die PDF-spezifischen Teile:
* die Woerter mit Koordinaten aus der Textebene,
* die Unterkante des Datenbereichs aus den gezeichneten Spaltentrennern,
* eine Seitenklassifikation ohne Modell.

Die eigentliche Rekonstruktion der Tabelle uebernimmt wordgrid.py, das
dieselbe Logik auch auf OCR-Woerter gescannter Seiten anwendet.
"""

from __future__ import annotations

import logging
from collections import Counter
from pathlib import Path

from ..textutil import normalize
from .base import ExtractionMode, RawTable
from .wordgrid import Word, extract_table_from_words

log = logging.getLogger(__name__)


def pdf_has_text_layer(path: str | Path, min_chars: int = 200) -> bool:
    """True, wenn mindestens eine Seite eine brauchbare Textebene besitzt."""
    import pymupdf

    with pymupdf.open(path) as doc:
        for page in doc:
            if len(page.get_text().strip()) >= min_chars:
                return True
    return False


def _vertical_segments(page) -> list[tuple[float, float, float]]:
    """Alle senkrechten Linien der Seite als (x, y_oben, y_unten)."""
    try:
        drawings = page.get_drawings()
    except Exception:                       # pragma: no cover - defensiv
        return []

    segments: list[tuple[float, float, float]] = []
    for item in drawings:
        for operation in item.get("items", []):
            if operation[0] == "l":
                start, end = operation[1], operation[2]
                x0, x1 = start.x, end.x
                y0, y1 = min(start.y, end.y), max(start.y, end.y)
            elif operation[0] == "re":
                rect = operation[1]
                x0, x1 = rect.x0, rect.x1
                y0, y1 = rect.y0, rect.y1
            else:
                continue
            if abs(x1 - x0) < 1.5 and (y1 - y0) > 20.0:
                segments.append((x0, y0, y1))
    return segments


def make_table_bottom_finder(page):
    """Liefert eine Funktion, die die Unterkante des Datenbereichs bestimmt.

    Die senkrechten Trennlinien der Funktionsspalten laufen genau ueber den
    Datenbereich: sie beginnen an der Unterkante der Kopfzeilen und enden am
    Tabellenende. Damit ist die Grenze zwischen Tabelle und Fussbereich exakt
    bestimmt - Postleitzahlen, Revisionsstaende oder Planerangaben koennen
    nicht mehr als Datenpunkte gelesen werden.
    """
    segments = _vertical_segments(page)

    def finder(body_top: float, x_left: float, x_right: float,
               pitch: float, row_pitch: float) -> float | None:
        if not segments:
            return None
        # Toleranz aus der gemessenen Geometrie der Seite, keine feste Groesse.
        start_tolerance = max(pitch * 0.5, row_pitch * 0.9)
        links, rechts = x_left - pitch, x_right + pitch
        bottoms = [
            round(y1, 1) for x0, y0, y1 in segments
            if links <= x0 <= rechts
            and abs(y0 - body_top) <= start_tolerance
            and y1 > body_top + start_tolerance
        ]
        if not bottoms:
            return None
        # Die Spaltentrenner enden alle auf derselben Hoehe - haeufigster Wert.
        haeufigkeit = Counter(bottoms)
        hoechste = max(haeufigkeit.values())
        return max(value for value, count in haeufigkeit.items() if count == hoechste)

    return finder


def extract_pdf_text(path: str | Path, page_index: int) -> RawTable | None:
    """Extrahiert eine PDF-Seite ueber ihre Textebene.

    Rueckgabe None, wenn die Seite kein erkennbares Abschnitt-/Spalte-Raster
    hat (dann ist es vermutlich ein Regelschema oder ein Scan).
    """
    import pymupdf

    with pymupdf.open(path) as doc:
        page = doc[page_index]
        words: list[Word] = page.get_text("words")
        page_width = page.rect.width
        finder = make_table_bottom_finder(page)

        return extract_table_from_words(
            words, page_width, page_index, ExtractionMode.PDF_TEXT,
            table_bottom_finder=finder,
        )


SCHEMA_KEYWORDS = (
    "regelschema", "prinzipschaltbild", "anlagenschema", "strangschema",
    "funktionsschema", "schaltplan", "r+i", "rls-schema",
)
TABLE_KEYWORDS = ("summe funktionen", "funktionsliste", "datenpunktliste",
                  "verarbeitungsfunktionen", "anwendungsfunktionen")


def classify_page_text(path: str | Path, page_index: int) -> tuple[str, float, str]:
    """Klassifiziert eine PDF-Seite ohne Modell, allein aus der Textebene.

    Rueckgabe: (typ, konfidenz, begruendung) mit typ aus
    {"funktionsliste", "regelschema", "sonstiges", "unbekannt"}.
    "unbekannt" heisst: bitte OCR oder das Vision-Modell fragen.
    """
    import pymupdf

    with pymupdf.open(path) as doc:
        page = doc[page_index]
        text = page.get_text()
        words = page.get_text("words")
        try:
            drawings = len(page.get_drawings())
        except Exception:                       # pragma: no cover
            drawings = 0

    lowered = normalize(text)
    if len(text.strip()) < 30:
        return "unbekannt", 0.0, "Keine verwertbare Textebene (vermutlich Scan)"
    if any(normalize(keyword) in lowered for keyword in SCHEMA_KEYWORDS):
        return "regelschema", 0.85, "Schema-Stichwort in der Textebene gefunden"
    if any(normalize(keyword) in lowered for keyword in TABLE_KEYWORDS):
        return "funktionsliste", 0.7, "Tabellen-Stichwort gefunden, aber kein Abschnitt-/Spalte-Raster"
    if drawings > 150 and len(words) < 400:
        return "regelschema", 0.7, f"{drawings} Vektorpfade bei nur {len(words)} Woertern"
    return "sonstiges", 0.5, "Kein Tabellenraster erkennbar"
