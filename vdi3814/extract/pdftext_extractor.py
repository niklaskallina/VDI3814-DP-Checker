"""Extraktion aus PDFs mit Textebene (aus Excel/CAD gedruckte Funktionslisten).

Die meisten GA-Funktionslisten liegen als PDF-Druck einer Excel-Vorlage vor.
Dann sind alle Zellwerte als Text mit exakten Koordinaten vorhanden - eine
Bilderkennung waere hier unnoetig ungenau. Dieses Backend rekonstruiert das
Tabellenraster aus den Kopfzeilen "Abschnitt" und "Spalte".

Nur wenn diese Kopfzeilen fehlen oder das PDF keine Textebene hat, uebernimmt
das Vision-Backend.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from pathlib import Path

from ..models import ColumnHeader, DataPointRow, Cell, DocumentMetadata, Footnote
from ..textutil import extract_footnote_markers, normalize, normalize_address, parse_count
from .base import (ExtractionMode, RawTable, is_sum_label, nearest_section,
                   sections_by_position, truncate_after_sum)

log = logging.getLogger(__name__)

Word = tuple[float, float, float, float, str, int, int, int]

# Toleranzen in PDF-Punkten
ROW_TOLERANCE = 4.0        # Zeilenhoehe-Streuung beim y-Clustern
LABEL_ROW_TOLERANCE = 5.0  # Suchband um die Kopfzeilen "Abschnitt"/"Spalte"

_INT_RE = re.compile(r"^\d{1,3}$")
_ADDR_RE = re.compile(r"^\d{1,2}(\.\d{1,2})*\.?$")


def _is_rotated(word: Word) -> bool:
    """Senkrecht gesetzter Text: Wortbox ist hoeher als breit."""
    return (word[3] - word[1]) > (word[2] - word[0]) * 1.2


def _cx(word: Word) -> float:
    return (word[0] + word[2]) / 2.0


def _cy(word: Word) -> float:
    return (word[1] + word[3]) / 2.0


def pdf_has_text_layer(path: str | Path, min_chars: int = 200) -> bool:
    """True, wenn mindestens eine Seite eine brauchbare Textebene besitzt."""
    import pymupdf

    with pymupdf.open(path) as doc:
        for page in doc:
            if len(page.get_text().strip()) >= min_chars:
                return True
    return False


@dataclass
class _HeaderGrid:
    columns: list[ColumnHeader]
    body_top: float          # y, ab der die Datenzeilen beginnen
    label_left: float        # x-Grenze: links davon steht die Datenpunktbenennung
    note_left: float | None  # x, ab der die Bemerkungsspalte beginnt
    labels_top: float = 0.0  # y, ab der die Spaltenueberschriften beginnen


def _find_label_word(words: list[Word], *labels: str) -> Word | None:
    wanted = {normalize(label) for label in labels}
    for word in words:
        if normalize(word[4]) in wanted:
            return word
    return None


def _band(words: list[Word], anchor: Word, tolerance: float = LABEL_ROW_TOLERANCE) -> list[Word]:
    center = _cy(anchor)
    return sorted(
        (w for w in words if abs(_cy(w) - center) <= tolerance),
        key=lambda w: w[0],
    )


def _rotated_labels(words: list[Word], columns: list[ColumnHeader], header_bottom: float,
                    pitch: float, used: set[int]) -> float:
    """Rekonstruiert die senkrecht gedruckten Spaltenueberschriften.

    Bei 90 Grad gedrehtem Text liefert PyMuPDF pro Wort eine schmale, hohe Box
    an der x-Position der Spalte; die Leserichtung ist von unten nach oben.
    Oberhalb der Ueberschriften steht haeufig der Fussnotenblock - er wird
    dadurch abgetrennt, dass nur der zusammenhaengende Wortstapel direkt ueber
    der Kopfzeile uebernommen wird (Abbruch bei einer groesseren Luecke).

    Rueckgabe: die oberste y-Koordinate der erkannten Ueberschriften. Alles
    darueber gehoert nicht mehr zur Tabelle.
    """
    tolerance = max(3.0, pitch * 0.45)
    top = header_bottom
    for column in columns:
        if column.x_center is None:
            continue
        # Ein Wort gehoert zur Spaltenueberschrift, wenn seine Box in die
        # Spaltenbreite passt (senkrechter Text) - der Aspekt-Test allein
        # scheitert an kurzen Fussnotenzeichen wie "1)".
        parts = [
            (index, w) for index, w in enumerate(words)
            if w[3] <= header_bottom
            and abs(_cx(w) - column.x_center) <= tolerance
            and (w[2] - w[0]) <= pitch * 1.15
        ]
        if not parts:
            continue
        parts.sort(key=lambda item: -item[1][1])
        stack = [parts[0]]
        for index, word in parts[1:]:
            if stack[-1][1][1] - word[3] > 8.0:   # Luecke -> Fussnotenblock beginnt
                break
            stack.append((index, word))
        text = " ".join(w[4] for _, w in stack).strip()
        column.label = text
        column.footnote_markers = extract_footnote_markers(text)
        used.update(index for index, _ in stack)
        top = min(top, min(w[1] for _, w in stack))
    return top


def _build_grid(words: list[Word], page_width: float, used: set[int]) -> _HeaderGrid | None:
    section_anchor = _find_label_word(words, "Abschnitt")
    column_anchor = _find_label_word(words, "Spalte")
    if section_anchor is None or column_anchor is None:
        return None

    section_band = [w for w in _band(words, section_anchor) if _cx(w) > _cx(section_anchor)]
    column_band = [w for w in _band(words, column_anchor) if _cx(w) > _cx(column_anchor)]

    sections = [(_cx(w), normalize_address(w[4])) for w in section_band if _ADDR_RE.match(w[4].strip())]
    column_words = [w for w in column_band if _INT_RE.match(w[4].strip())]
    if len(column_words) < 4:
        return None

    numbers = [w[4].strip() for w in column_words]
    section_ids = sections_by_position(numbers, [_cx(w) for w in column_words], sections)

    columns: list[ColumnHeader] = []
    for index, word in enumerate(column_words):
        section = section_ids[index] if section_ids else nearest_section(_cx(word), sections)
        columns.append(
            ColumnHeader(
                index=index,
                section=section,
                column_no=word[4].strip(),
                x_center=_cx(word),
            )
        )

    xs = [c.x_center for c in columns if c.x_center is not None]
    pitch = (max(xs) - min(xs)) / max(1, len(xs) - 1)
    header_bottom = min(_cy(section_anchor), _cy(column_anchor)) - LABEL_ROW_TOLERANCE
    labels_top = _rotated_labels(words, columns, header_bottom, pitch, used)

    note_anchor = _find_label_word(words, "Bemerkung", "Bemerkungen", "Bemerkungen und Referenzierungen")
    note_left = None
    if note_anchor is not None and note_anchor[0] > max(xs):
        note_left = note_anchor[0] - pitch
        columns.append(
            ColumnHeader(
                index=len(columns),
                label=note_anchor[4],
                section=nearest_section(_cx(note_anchor), sections),
                x_center=_cx(note_anchor),
            )
        )

    return _HeaderGrid(
        columns=columns,
        body_top=max(_cy(section_anchor), _cy(column_anchor)) + LABEL_ROW_TOLERANCE,
        label_left=min(xs) - pitch * 0.6,
        note_left=note_left,
        labels_top=labels_top,
    )


def _cluster_rows(words: list[Word], tolerance: float = ROW_TOLERANCE) -> list[list[Word]]:
    """Gruppiert Woerter zu Tabellenzeilen ueber ihre y-Mitte."""
    ordered = sorted(words, key=_cy)
    clusters: list[list[Word]] = []
    for word in ordered:
        if clusters and abs(_cy(word) - _cy(clusters[-1][-1])) <= tolerance:
            clusters[-1].append(word)
        else:
            clusters.append([word])
    return [sorted(cluster, key=lambda w: w[0]) for cluster in clusters]


def _extract_metadata(words: list[Word], grid: _HeaderGrid, skip: set[int]) -> DocumentMetadata:
    """Liest Kopf-/Fussbereich ueber Beschriftungen wie 'Gewerk:' oder 'Anlage:'."""
    labels = {
        "gewerk": "gewerk",
        "anlage": "anlage",
        "projekt": "projekt",
        "auftraggeber projekt": "auftraggeber",
        "auftraggeber": "auftraggeber",
        "planersteller": "planersteller",
        "planstand": "planstand",
        "blatt nr": "blatt_nr",
        "von": "blatt_von",   # nur gueltig, wenn "Blatt Nr." schon gefunden wurde
        "informationsschwerpunkt": "informationsschwerpunkt",
        "datenkommunikationsprotokoll": "protokoll",
        "ausgabedatum": "datum",
        "datum": "datum",
    }
    # Woerter, die im Fussbereich als weitere Beschriftung dienen und deshalb
    # nie ein Wert sein koennen.
    stop_words = {"name", "geprueft", "inhalt", "index", "rev", "blatt", "datei",
                  "asp", "seite", "stand"}
    metadata = DocumentMetadata()
    rows = _cluster_rows(
        [w for index, w in enumerate(words) if index not in skip and not _is_rotated(w)],
        tolerance=3.0,
    )
    for row in rows:
        for index, word in enumerate(row):
            key = normalize(word[4].rstrip(":"))
            field_name = labels.get(key)
            if not field_name or getattr(metadata, field_name):
                continue
            if field_name == "blatt_von" and not metadata.blatt_nr:
                continue
            # Wert = folgende Woerter derselben Zeile, bis zum naechsten Label
            values: list[str] = []
            previous_right = word[2]
            for follower in row[index + 1:]:
                follower_key = normalize(follower[4].rstrip(":"))
                if follower_key in labels or follower_key in stop_words:
                    break
                if follower[0] - previous_right > 40.0:
                    break                      # groessere Luecke: anderer Block
                if word[0] < grid.label_left <= follower[0]:
                    break                      # ab hier beginnt die Tabelle
                values.append(follower[4])
                previous_right = follower[2]
                if len(" ".join(values)) > 60:
                    break
            if values:
                setattr(metadata, field_name, " ".join(values).strip())
    return metadata


def _extract_footnotes(words: list[Word], header_top: float, skip: set[int]) -> list[Footnote]:
    """Liest die nummerierten Fussnoten oberhalb des Tabellenkopfs."""
    candidates = [w for index, w in enumerate(words)
                  if w[3] <= header_top and index not in skip]
    rows = _cluster_rows(candidates, tolerance=3.0)
    footnotes: list[Footnote] = []
    for row in rows:
        text = " ".join(w[4] for w in row).strip()
        for match in re.finditer(r"(?<![\d.,])(\d{1,2})\)\s*", text):
            marker = match.group(1)
            start = match.end()
            next_match = re.search(r"(?<![\d.,])\d{1,2}\)", text[start:])
            body = text[start:start + next_match.start()] if next_match else text[start:]
            body = body.strip()
            if body:
                footnotes.append(Footnote(marker=marker, text=body))
    # Mehrfachfunde je Marker zusammenfuehren (Fussnoten laufen ueber mehrere Zeilen)
    merged: dict[str, Footnote] = {}
    for footnote in footnotes:
        if footnote.marker in merged:
            merged[footnote.marker].text += " " + footnote.text
        else:
            merged[footnote.marker] = footnote
    return [merged[key] for key in sorted(merged, key=lambda m: int(m))]


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
        page_height = page.rect.height

    if len(words) < 20:
        return None

    label_word_ids: set[int] = set()
    grid = _build_grid(words, page_width, label_word_ids)
    if grid is None:
        return None

    xs = [c.x_center for c in grid.columns if c.x_center is not None]
    pitch = (max(xs) - min(xs)) / max(1, len(xs) - 1)
    tolerance = max(3.0, pitch * 0.5)
    right_limit = grid.note_left if grid.note_left else max(xs) + pitch

    body_words = [w for w in words if _cy(w) > grid.body_top]
    rows: list[DataPointRow] = []
    for cluster in _cluster_rows(body_words):
        label_words = [w for w in cluster if _cx(w) < grid.label_left]
        label = " ".join(w[4] for w in label_words).strip()
        # Fuehrende Zeilennummer abtrennen
        row_no = ""
        parts = label.split(" ", 1)
        if parts and _INT_RE.match(parts[0]):
            row_no, label = parts[0], (parts[1] if len(parts) > 1 else "")

        cells: list[Cell] = []
        for column in grid.columns:
            if column.x_center is None or column.x_center > right_limit:
                continue
            hits = [w for w in cluster if abs(_cx(w) - column.x_center) <= tolerance]
            if not hits:
                continue
            raw = " ".join(w[4] for w in hits).strip()
            count = parse_count(raw)
            cells.append(Cell(column_index=column.index, raw_value=raw, count=count,
                              note="" if count is not None else raw))

        remark = ""
        if grid.note_left is not None:
            note_words = [w for w in cluster
                          if grid.note_left <= _cx(w) and _cx(w) < page_width - 25]
            remark = " ".join(w[4] for w in note_words).strip()
            # Rechts wiederholte Zeilennummer nicht als Bemerkung uebernehmen
            if remark == row_no:
                remark = ""

        sum_row = is_sum_label(label) or is_sum_label(" ".join(w[4] for w in cluster[:2]))
        if not label and not cells and not remark:
            continue
        rows.append(
            DataPointRow(
                row_no=row_no,
                klartext=label,
                remark=remark,
                cells=cells,
                page_index=page_index,
                is_sum_row=sum_row,
                confidence=1.0,
            )
        )

    rows = truncate_after_sum(rows)

    table = RawTable(
        page_index=page_index,
        mode=ExtractionMode.PDF_TEXT,
        columns=grid.columns,
        rows=rows,
        footnotes=_extract_footnotes(words, header_top=grid.labels_top, skip=label_word_ids),
        metadata=_extract_metadata(words, grid, label_word_ids),
        texts=[" ".join(w[4] for w in words)],
    )
    return table
