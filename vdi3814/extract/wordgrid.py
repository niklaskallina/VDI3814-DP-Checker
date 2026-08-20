"""Rekonstruktion einer Tabelle aus Woertern mit Koordinaten.

Diese Schicht ist bewusst quellenunabhaengig: sie bekommt nur eine Liste von
Woertern mit Position (x0, y0, x1, y1, Text) und rekonstruiert daraus das
Tabellenraster. Dieselbe Logik bedient damit

* die Textebene digitaler PDFs (Koordinaten in PDF-Punkten) und
* die OCR-Ausgabe gescannter Seiten (Koordinaten in Pixeln).

Es gibt keine festen Positionen: Spaltenmitten stammen aus der Kopfzeile
"Spalte" des jeweiligen Dokuments, die Oberkante des Datenbereichs aus den
gefundenen Kopfzeilen, alle Toleranzen aus dem gemessenen Spalten- und
Zeilenabstand derselben Seite.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass

from ..models import Cell, ColumnHeader, DataPointRow, DocumentMetadata, Footnote, RowKind
from ..textutil import extract_footnote_markers, normalize, normalize_address, parse_count
from .base import ExtractionMode, RawTable, nearest_section, sections_by_position, drop_after_footer
from .rowfilter import classify_row

log = logging.getLogger(__name__)

Word = tuple[float, float, float, float, str, int, int, int]

# Startwerte fuer die erste Messung. Alle spaeteren Toleranzen werden aus der
# tatsaechlich gemessenen Geometrie der jeweiligen Seite abgeleitet
# (Spaltenabstand und Zeilenabstand) - es gibt keine festen Positionen.
ROW_TOLERANCE = 4.0        # nur fuer den ersten Clusterdurchlauf
LABEL_ROW_TOLERANCE = 5.0  # Suchband um die Kopfzeilen "Abschnitt"/"Spalte"

_INT_RE = re.compile(r"^\d{1,3}$")
_ADDR_RE = re.compile(r"^\d{1,2}(\.\d{1,2})*\.?$")


def _is_rotated(word: Word) -> bool:
    """Senkrecht gesetzter Text: Wortbox ist deutlich hoeher als breit.

    Die Mindesthoehe verhindert Fehltreffer bei kurzen waagerechten Woertern
    wie "01" oder "B", deren Box zufaellig hochkant wirkt.
    """
    height = word[3] - word[1]
    return height > (word[2] - word[0]) * 1.6 and height > 12.0


def _cx(word: Word) -> float:
    return (word[0] + word[2]) / 2.0


def _cy(word: Word) -> float:
    return (word[1] + word[3]) / 2.0


Word = tuple[float, float, float, float, str, int, int, int]

# Startwerte fuer die erste Messung; alle weiteren Toleranzen werden aus der
# gemessenen Geometrie der Seite abgeleitet.
ROW_TOLERANCE = 4.0
LABEL_ROW_TOLERANCE = 5.0

_INT_RE = re.compile(r"^\d{1,3}$")
_ADDR_RE = re.compile(r"^\d{1,2}(\.\d{1,2})*\.?$")


@dataclass
class _HeaderGrid:
    columns: list[ColumnHeader]
    body_top: float          # y, ab der die Datenzeilen beginnen
    label_left: float        # x-Grenze: links davon steht die Datenpunktbenennung
    note_left: float | None  # x, ab der die Bemerkungsspalte beginnt
    labels_top: float = 0.0  # y, ab der die Spaltenueberschriften beginnen



def _is_rotated(word: Word) -> bool:
    """Senkrecht gesetzter Text: Wortbox ist deutlich hoeher als breit.

    Die Mindesthoehe verhindert Fehltreffer bei kurzen waagerechten Woertern
    wie "01" oder "B", deren Box zufaellig hochkant wirkt.
    """
    height = word[3] - word[1]
    return height > (word[2] - word[0]) * 1.6 and height > 12.0



def _cx(word: Word) -> float:
    return (word[0] + word[2]) / 2.0



def _cy(word: Word) -> float:
    return (word[1] + word[3]) / 2.0



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



def _row_pitch(clusters: list[list[Word]]) -> float:
    """Misst den Zeilenabstand der Seite (Median der Abstaende der Zeilenmitten).

    Damit sind alle weiteren Toleranzen von der Schriftgroesse und Aufloesung
    des Dokuments unabhaengig.
    """
    centers = sorted(_cy(cluster[0]) for cluster in clusters)
    gaps = [b - a for a, b in zip(centers, centers[1:]) if 1.0 < b - a < 200.0]
    if not gaps:
        return 0.0
    gaps.sort()
    return gaps[len(gaps) // 2]



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

    def is_label_row(row: list[Word]) -> bool:
        """Fussbereichs-Beschriftungszeile ('Ausgabedatum | Name | Planersteller | ...').

        Dort stehen die Werte UNTER der Beschriftung, nicht daneben.
        """
        hits = sum(1 for w in row if normalize(w[4].rstrip(":")) in labels)
        return hits >= 2

    rows = _cluster_rows(
        [w for index, w in enumerate(words) if index not in skip and not _is_rotated(w)],
        tolerance=3.0,
    )
    for row_index, row in enumerate(rows):
        for index, word in enumerate(row):
            # Beschriftungen aus zwei Woertern ("Blatt Nr.") mit beruecksichtigen
            key = normalize(word[4].rstrip(":"))
            skip_next = 0
            if key not in labels and index + 1 < len(row):
                combined = normalize(f"{word[4]} {row[index + 1][4]}".rstrip(":"))
                if combined in labels:
                    key, skip_next = combined, 1
            field_name = labels.get(key)
            if not field_name or getattr(metadata, field_name):
                continue
            if field_name == "blatt_von" and not metadata.blatt_nr:
                continue
            # Wert = folgende Woerter derselben Zeile, bis zum naechsten Label
            values: list[str] = []
            previous_right = row[index + skip_next][2]
            for follower in row[index + 1 + skip_next:]:
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
            if not values and is_label_row(row):
                # Wert in einer der naechsten Zeilen an derselben x-Position suchen
                right_border = min(
                    (w[0] for w in row[index + 1 + skip_next:]
                     if normalize(w[4].rstrip(":")) in labels),
                    default=float("inf"),
                )
                for follower_row in rows[row_index + 1: row_index + 4]:
                    below = [w for w in follower_row
                             if word[0] - 6.0 <= w[0] < right_border
                             and normalize(w[4].rstrip(":")) not in labels
                             and normalize(w[4].rstrip(":")) not in stop_words]
                    if not below:
                        continue
                    candidate = [w[4] for w in below]
                    # Eine blosse Zahl unter einer Textbeschriftung ist meist ein
                    # Revisionsindex o. Ae. - nur bei Blattangaben ist sie gewollt.
                    if (len(candidate) == 1 and candidate[0].strip().isdigit()
                            and field_name not in ("blatt_nr", "blatt_von")):
                        continue
                    values = candidate
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




def extract_table_from_words(words: list[Word], page_width: float, page_index: int,
                             mode: ExtractionMode,
                             table_bottom_finder=None) -> RawTable | None:
    """Rekonstruiert die Funktionsliste aus Woertern mit Koordinaten.

    table_bottom_finder: optionale Funktion
        (body_top, x_links, x_rechts, spaltenabstand, zeilenabstand) -> y | None
    Liefert die Unterkante des Datenbereichs, sofern die Quelle sie bestimmen
    kann (bei PDFs aus den gezeichneten Spaltentrennern, bei Scans aus den
    erkannten Tabellenlinien). Ohne sie greifen die Schlagwortregeln.
    """
    if len(words) < 20:
        return None

    label_word_ids: set[int] = set()
    grid = _build_grid(words, page_width, label_word_ids)
    if grid is None:
        return None

    xs = [c.x_center for c in grid.columns if c.x_center is not None]
    pitch = (max(xs) - min(xs)) / max(1, len(xs) - 1)
    # Enge Toleranz: ein Wert zaehlt nur, wenn er wirklich in der Spalte steht.
    tolerance = max(pitch * 0.12, pitch * 0.45)
    right_limit = grid.note_left if grid.note_left else max(xs) + pitch

    body_words = [w for w in words if _cy(w) > grid.body_top]
    row_pitch = _row_pitch(_cluster_rows(body_words))
    row_tolerance = max(pitch * 0.2, row_pitch * 0.35) if row_pitch else ROW_TOLERANCE

    bottom = None
    if table_bottom_finder is not None:
        bottom = table_bottom_finder(grid.body_top, min(xs), max(xs), pitch, row_pitch)

    warnings: list[str] = []
    rows: list[DataPointRow] = []
    footer_reached = False
    previous_number: int | None = None

    for cluster in _cluster_rows(body_words, tolerance=row_tolerance):
        cluster_y = _cy(cluster[0])
        below_table = bottom is not None and cluster_y > bottom + row_tolerance

        label_words = [w for w in cluster if _cx(w) < grid.label_left]
        label = " ".join(w[4] for w in label_words).strip()
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
                          if grid.note_left <= _cx(w) < page_width - pitch * 2]
            remark = " ".join(w[4] for w in note_words).strip()
            if remark == row_no:
                remark = ""

        if not label and not cells and not remark and not row_no:
            continue

        has_values = any(cell.count is not None for cell in cells)
        kind, reason = classify_row(label, has_values, remark)

        if below_table:
            # Unterhalb der Tabelle ist nur die Summenzeile interessant (Kontrolle).
            if kind not in (RowKind.SUMME, RowKind.UEBERTRAG):
                continue
        elif footer_reached:
            continue
        elif kind is RowKind.FUSSBEREICH and bottom is None:
            footer_reached = True
            continue

        if kind is RowKind.DATEN and row_no.isdigit():
            number = int(row_no)
            if previous_number is not None and number <= previous_number:
                warnings.append(
                    f"Zeilennummerierung springt zurueck ({previous_number} -> {number}); "
                    f"Zeile '{label[:40]}' bitte in der Vorschau pruefen"
                )
            previous_number = number

        rows.append(
            DataPointRow(
                row_no=row_no,
                klartext=label,
                remark=remark,
                cells=cells,
                page_index=page_index,
                kind=kind,
                exclusion_reason=reason,
                confidence=1.0,
            )
        )

    if bottom is None:
        rows = drop_after_footer(rows)

    return RawTable(
        page_index=page_index,
        mode=mode,
        columns=grid.columns,
        rows=rows,
        footnotes=_extract_footnotes(words, header_top=grid.labels_top, skip=label_word_ids),
        metadata=_extract_metadata(words, grid, label_word_ids),
        texts=[" ".join(w[4] for w in words)],
        warnings=warnings,
    )
