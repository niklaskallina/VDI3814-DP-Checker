"""Gemeinsames Zwischenformat und Hilfen fuer alle Extraktions-Backends."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

from ..models import ColumnHeader, DataPointRow, DocumentMetadata, Footnote
from ..textutil import normalize, normalize_address


class ExtractionMode(str, Enum):
    EXCEL = "excel"
    PDF_TEXT = "pdf_text"
    VISION = "vision"


@dataclass
class RawTable:
    """Roh extrahierte Tabelle einer Seite - noch ohne Profil-Zuordnung."""

    page_index: int = 0
    mode: ExtractionMode = ExtractionMode.VISION
    columns: list[ColumnHeader] = field(default_factory=list)
    rows: list[DataPointRow] = field(default_factory=list)
    footnotes: list[Footnote] = field(default_factory=list)
    metadata: DocumentMetadata = field(default_factory=DocumentMetadata)
    texts: list[str] = field(default_factory=list)   # Grundlage der Fassungserkennung
    warnings: list[str] = field(default_factory=list)

    @property
    def has_data(self) -> bool:
        return bool(self.columns) and any(r.is_countable for r in self.rows)


# --------------------------------------------------------------------------
# Gemeinsame Heuristiken
# --------------------------------------------------------------------------

SUM_ROW_MARKERS = ("summe funktionen", "summe", "gesamtsumme", "total")


def is_sum_label(text: str) -> bool:
    norm = normalize(text)
    return any(norm.startswith(m) for m in SUM_ROW_MARKERS)


def assign_sections(column_numbers: list[str], section_numbers: list[str]) -> list[str]:
    """Ordnet jeder Spaltennummer ihren Abschnitt zu.

    Robust ueber den Neustart der Spaltennummerierung: die Folge
    1,2,3,4,5,1,2,3,4,5,1,... zerfaellt bei jeder 1 in einen neuen Abschnitt.
    Die so entstehenden Bloecke werden der Reihe nach den Abschnittsnummern
    zugeordnet. Passt die Anzahl nicht, wird leer zurueckgegeben (dann greift
    die x-basierte Zuordnung bzw. das Label-Matching).
    """
    blocks: list[list[int]] = []
    for index, number in enumerate(column_numbers):
        if number.strip() in ("1", "1.0") or not blocks:
            blocks.append([index])
        else:
            blocks[-1].append(index)
    if len(blocks) != len(section_numbers):
        return []
    result = [""] * len(column_numbers)
    for block, section in zip(blocks, section_numbers):
        for index in block:
            result[index] = normalize_address(section)
    return result


def nearest_section(x: float, sections: list[tuple[float, str]]) -> str:
    """Faellt die Blockzuordnung aus, wird der raeumlich naechste Abschnitt genommen."""
    if not sections:
        return ""
    return min(sections, key=lambda item: abs(item[0] - x))[1]


def sections_by_position(column_numbers: list[str], column_xs: list[float],
                         sections: list[tuple[float, str]]) -> list[str]:
    """Ordnet Spaltenbloecke ueber ihre Lage den Abschnittsnummern zu.

    Die Spaltennummerierung startet in jedem Abschnitt neu bei 1; daraus
    ergeben sich Bloecke. Jeder Block bekommt die Abschnittsnummer, deren
    Beschriftung am naechsten an seiner Mitte steht - jede Abschnittsnummer
    wird nur einmal vergeben. Das funktioniert sowohl bei zentrierter
    Beschriftung (PDF-Druck) als auch bei linksbuendiger (Excel), und es
    stoert nicht, wenn es mehr Abschnittsnummern als Bloecke gibt (z. B. die
    Bemerkungsspalte ohne eigene Spaltennummer).
    """
    if not sections or not column_numbers:
        return []
    blocks: list[list[int]] = []
    for index, number in enumerate(column_numbers):
        if number.strip() in ("1", "1.0") or not blocks:
            blocks.append([index])
        else:
            blocks[-1].append(index)

    pairs: list[tuple[float, int, int]] = []
    for block_index, block in enumerate(blocks):
        center = sum(column_xs[i] for i in block) / len(block)
        left = column_xs[block[0]]
        for section_index, (x, _) in enumerate(sections):
            # Abstand zur Blockmitte ODER zum linken Blockrand - je nachdem,
            # wo die Vorlage die Abschnittsnummer setzt.
            distance = min(abs(x - center), abs(x - left))
            pairs.append((distance, block_index, section_index))

    pairs.sort()
    used_blocks: set[int] = set()
    used_sections: set[int] = set()
    assignment: dict[int, str] = {}
    for _, block_index, section_index in pairs:
        if block_index in used_blocks or section_index in used_sections:
            continue
        used_blocks.add(block_index)
        used_sections.add(section_index)
        assignment[block_index] = sections[section_index][1]

    result = [""] * len(column_numbers)
    for block_index, block in enumerate(blocks):
        for index in block:
            result[index] = assignment.get(block_index, "")
    return result


def drop_after_footer(rows: list) -> list:
    """Schneidet ab der ersten Zeile des Fussbereichs alles weg.

    Rueckfallebene fuer Seiten ohne gezeichneten Tabellenrahmen. Summen- und
    Uebertragszeilen bleiben erhalten - sie werden nicht gezaehlt, dienen aber
    der Kontrolle. Nachlaufende Leerzeilen werden verworfen.
    """
    from ..models import RowKind

    result = []
    for row in rows:
        if getattr(row, "kind", RowKind.DATEN) in (RowKind.FUSSBEREICH, RowKind.KOPF):
            break
        result.append(row)
    while result and getattr(result[-1], "kind", None) is RowKind.LEER:
        result.pop()
    return result
