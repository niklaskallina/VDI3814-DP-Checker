"""Extraktion aus Excel-Vorlagen (.xlsx/.xlsm/.xls).

Viele GA-Funktionslisten entstehen in der VDI-Excel-Vorlage. Liegt die
Originaldatei vor, ist das die genaueste Quelle ueberhaupt - jede Zelle wird
1:1 gelesen, es gibt kein Erkennungsrisiko. Auch hier ist der Anker das
Kopfzeilenpaar "Abschnitt" / "Spalte".
"""

from __future__ import annotations

import logging
from pathlib import Path

from ..models import Cell, ColumnHeader, DataPointRow, DocumentMetadata, Footnote, RowKind
from ..textutil import extract_footnote_markers, normalize, normalize_address, parse_count
from .base import (ExtractionMode, RawTable, assign_sections, drop_after_footer,
                   sections_by_position)
from .rowfilter import classify_row, markiere_unbeschriftete_summenzeilen

log = logging.getLogger(__name__)

EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xls"}

METADATA_LABELS = {
    "gewerk": "gewerk",
    "anlage": "anlage",
    "projekt": "projekt",
    "auftraggeber projekt": "auftraggeber",
    "auftraggeber": "auftraggeber",
    "planersteller": "planersteller",
    "planstand": "planstand",
    "blatt nr": "blatt_nr",
    "informationsschwerpunkt": "informationsschwerpunkt",
    "isp": "informationsschwerpunkt",
    "automationsschwerpunkt": "automationsschwerpunkt",
    "asp": "automationsschwerpunkt",
    "datenkommunikationsprotokoll": "protokoll",
    "datenkommunikationsprotokoll ae asp": "protokoll",
    "ausgabedatum": "datum",
    "datum": "datum",
}


def is_excel(path: str | Path) -> bool:
    return Path(path).suffix.lower() in EXCEL_SUFFIXES


def _cell_text(value) -> str:
    """Zellwert als Text; ganzzahlige Floats ohne '.0' (Excel liefert 1.0 statt 1)."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


class _Raster(list):
    """Textraster eines Blatts [zeile][spalte] plus das, was der Text nicht zeigt.

    ``kopien``      Zellen, die nur durch das Ausrollen einer verbundenen Zelle
                    einen Wert tragen. Im Tabellenkopf ist das gewollt
                    (Gruppenueberschriften), im Datenbereich wuerde derselbe
                    Wert sonst mehrfach gezaehlt.
    ``ausgeblendet`` Zeilen (0-basiert), die in Excel ausgeblendet sind.
    ``durchgestrichen`` Zeilen, deren Inhalt ueberwiegend durchgestrichen ist.
    """

    def __init__(self, zeilen=()):
        super().__init__(zeilen)
        self.kopien: set[tuple[int, int]] = set()
        self.ausgeblendet: set[int] = set()
        self.durchgestrichen: set[int] = set()


def _ausrollen(grid: "_Raster", top: int, left: int, bottom: int, right: int) -> None:
    """Rollt eine verbundene Zelle (0-basiert, Ende exklusiv) ueber ihren Bereich aus."""
    if top >= len(grid) or left >= len(grid[top]):
        return
    value = grid[top][left]
    if not value:
        return
    for r in range(top, min(bottom, len(grid))):
        for c in range(left, min(right, len(grid[r]))):
            if not grid[r][c]:
                grid[r][c] = value
                grid.kopien.add((r, c))


def read_grid(path: str | Path) -> list["_Raster"]:
    """Liest alle Blaetter als Textraster [blatt][zeile][spalte].

    Verbundene Zellen werden auf alle beteiligten Zellen ausgerollt, damit
    Gruppenueberschriften ueber jeder zugehoerigen Spalte stehen. Welche Zellen
    dabei nur Kopien sind, haelt das Raster fest (siehe _Raster).
    """
    path = Path(path)
    if path.suffix.lower() == ".xls":
        import xlrd  # nur fuer das alte BIFF-Format noetig

        # formatting_info=True liefert bei .xls zusaetzlich die verbundenen
        # Zellen - noetig, damit Gruppenueberschriften ueber jeder Spalte stehen.
        try:
            book = xlrd.open_workbook(path, formatting_info=True)
        except NotImplementedError:      # pragma: no cover
            book = xlrd.open_workbook(path)
        sheets = []
        for sheet in book.sheets():
            grid = _Raster(
                [_cell_text(sheet.cell_value(r, c)) for c in range(sheet.ncols)]
                for r in range(sheet.nrows)
            )
            for r0, r1, c0, c1 in getattr(sheet, "merged_cells", []):
                _ausrollen(grid, r0, c0, r1, c1)
            for r, info in (getattr(sheet, "rowinfo_map", {}) or {}).items():
                if getattr(info, "hidden", False):
                    grid.ausgeblendet.add(r)
            sheets.append(grid)
        return sheets

    import openpyxl

    book = openpyxl.load_workbook(path, data_only=True, read_only=False)
    sheets = []
    for sheet in book.worksheets:
        grid = _Raster()
        for r in range(1, sheet.max_row + 1):
            zeile, belegt, gestrichen = [], 0, 0
            for c in range(1, sheet.max_column + 1):
                zelle = sheet.cell(r, c)
                text = _cell_text(zelle.value)
                zeile.append(text)
                if text:
                    belegt += 1
                    if zelle.font is not None and zelle.font.strike:
                        gestrichen += 1
            grid.append(zeile)
            if belegt and gestrichen * 2 > belegt:
                grid.durchgestrichen.add(r - 1)
            dimension = sheet.row_dimensions.get(r)
            if dimension is not None and dimension.hidden:
                grid.ausgeblendet.add(r - 1)
        for merged in sheet.merged_cells.ranges:
            _ausrollen(grid, merged.min_row - 1, merged.min_col - 1,
                       merged.max_row, merged.max_col)
        sheets.append(grid)
    return sheets


def _find_label(grid: list[list[str]], *labels: str) -> tuple[int, int] | None:
    wanted = {normalize(label) for label in labels}
    for r, row in enumerate(grid):
        for c, value in enumerate(row):
            if normalize(value) in wanted:
                return r, c
    return None


def _label_row_above(grid: list[list[str]], section_row: int, first_col: int, last_col: int) -> int | None:
    """Sucht oberhalb der Abschnitt-Zeile die Zeile mit den Spaltenbezeichnungen.

    Unterscheidungsmerkmal gegenueber den Gruppenzeilen: Spaltenbezeichnungen
    sind je Spalte verschieden, Gruppenueberschriften sind verbunden und
    wiederholen sich nach dem Ausrollen ueber viele Spalten. Gezaehlt wird
    deshalb die Anzahl UNTERSCHIEDLICHER Texte.
    """
    best, best_count = None, 0
    for r in range(max(0, section_row - 8), section_row):
        values = {
            grid[r][c] for c in range(first_col, min(last_col + 1, len(grid[r])))
            if len(grid[r][c]) > 3
        }
        if len(values) >= best_count:      # bei Gleichstand die untere Zeile
            best, best_count = r, len(values)
    return best if best_count >= 3 else None


def _ffill(grid: list[list[str]], row_index: int, start: int, end: int) -> list[str]:
    """Zieht Gruppenueberschriften nach rechts durch.

    In der alten .xls-Vorlage sind Gruppenzeilen nicht verbunden, der Text
    laeuft optisch nur ueber die Nachbarzellen. Fuer die Zuordnung wird der
    zuletzt gesetzte Wert bis zur naechsten Ueberschrift fortgeschrieben.
    """
    row = list(grid[row_index]) if 0 <= row_index < len(grid) else []
    while len(row) <= end:
        row.append("")
    current = ""
    for c in range(start, end + 1):
        if row[c]:
            current = row[c]
        else:
            row[c] = current
    return row


def _value_at(grid: list[list[str]], r: int, c: int) -> str:
    if 0 <= r < len(grid) and 0 <= c < len(grid[r]):
        return grid[r][c]
    return ""


def _is_label_row(row: list[str]) -> bool:
    """Beschriftungszeile im Fussbereich (z. B. 'Datum | Inhalt | Name | Index').

    Dort stehen die Werte UNTER der Beschriftung, sonst rechts daneben.
    """
    hits = sum(1 for value in row if normalize(value.rstrip(":")) in METADATA_LABELS)
    return hits >= 2


def _extract_metadata(grid: list[list[str]]) -> DocumentMetadata:
    metadata = DocumentMetadata()
    for r, row in enumerate(grid):
        for c, value in enumerate(row):
            key = normalize(value.rstrip(":"))
            field_name = METADATA_LABELS.get(key)
            if not field_name or getattr(metadata, field_name):
                continue
            candidates = [_value_at(grid, r, c + 1), _value_at(grid, r, c + 2)]
            if _is_label_row(row):
                candidates.insert(0, _value_at(grid, r + 1, c))
            for candidate in candidates:
                if not candidate or candidate.endswith(":"):
                    continue          # das ist die naechste Beschriftung, kein Wert
                if normalize(candidate.rstrip(":")) in METADATA_LABELS:
                    continue
                setattr(metadata, field_name, candidate)
                break
    return metadata


def _extract_footnotes(grid: list[list[str]], header_row: int) -> list[Footnote]:
    """Fussnoten stehen als '1)' + Text in den Zeilen oberhalb des Tabellenkopfs.

    Die Fussnoten sind in mehreren Spaltenbloecken nebeneinander gesetzt.
    Fortsetzungszeilen werden deshalb ueber ihre SPALTE der richtigen Fussnote
    zugeordnet - sonst wandert der Text der rechten Spalte an die linke Fussnote.
    """
    footnotes: dict[str, Footnote] = {}
    anchors: list[tuple[int, int, str]] = []      # (Zeile, Textspalte, Marker)

    for r in range(0, max(0, header_row)):
        for c, value in enumerate(grid[r]):
            text = value.strip()
            if not (0 < len(text) <= 3 and text.endswith(")") and text[:-1].isdigit()):
                continue
            marker = text[:-1]
            body_col = c + 1
            body = " ".join(
                part for part in (_value_at(grid, r, c + 1), _value_at(grid, r, c + 2)) if part
            ).strip()
            footnotes.setdefault(marker, Footnote(marker=marker, text=""))
            if body:
                footnotes[marker].text = (footnotes[marker].text + " " + body).strip()
            anchors.append((r, body_col, marker))

    for r in range(0, max(0, header_row)):
        for c, value in enumerate(grid[r]):
            text = value.strip()
            if len(text) <= 10 or (len(text) <= 3 and text.endswith(")")):
                continue
            if any(row == r and abs(col - c) <= 2 for row, col, _ in anchors):
                continue                       # das ist die Ursprungszeile selbst
            candidates = [
                (r - row, marker) for row, col, marker in anchors
                if 0 < r - row <= 5 and abs(col - c) <= 2
            ]
            if not candidates:
                continue
            marker = min(candidates)[1]
            existing = footnotes[marker]
            if text not in existing.text:
                existing.text = (existing.text + " " + text).strip()

    return [footnotes[k] for k in sorted(footnotes, key=lambda m: int(m))]


_KOPF_BESCHRIFTUNGEN = {"abschnitt", "spalte", "abschnitt nr", "spalte nr"}


def _ist_wiederholter_kopf(left_texts: list[str], cells: list[Cell],
                           numbers: list[str]) -> bool:
    """Wiederholung der Nummernzeilen "Abschnitt"/"Spalte" mitten im Blatt."""
    if any(normalize(text) in _KOPF_BESCHRIFTUNGEN for text in left_texts):
        return True
    # Ohne Beschriftung: die Werte sind genau die Spaltennummern des Kopfs
    werte = {cell.column_index: cell.raw_value.strip() for cell in cells}
    return (len(werte) >= 4 and len(werte) == len(numbers)
            and all(werte.get(i) == nummer.strip() for i, nummer in enumerate(numbers)))


def _extract_sheet(grid: list[list[str]], sheet_index: int) -> RawTable | None:
    section_pos = _find_label(grid, "Abschnitt")
    column_pos = _find_label(grid, "Spalte")
    if section_pos is None or column_pos is None:
        return None
    section_row, label_col = section_pos
    column_row, _ = column_pos

    number_row = grid[column_row]
    column_indices = [
        c for c, value in enumerate(number_row)
        if c > label_col and value.strip().isdigit()
    ]
    if len(column_indices) < 4:
        return None

    sections = [(c, normalize_address(grid[section_row][c]))
                for c in range(label_col + 1, len(grid[section_row]))
                if normalize_address(_value_at(grid, section_row, c))]

    numbers = [number_row[c] for c in column_indices]
    section_ids = assign_sections(numbers, [s for _, s in sections])
    if not section_ids:
        # Anzahl der Abschnittsnummern passt nicht zu den Bloecken (verbundene
        # Zellen, Bemerkungsabschnitt ohne Spaltennummer). Dann ueber die Lage:
        # links- wie mittig gesetzte Abschnittsnummern treffen ihren Block.
        section_ids = sections_by_position(
            numbers, [float(c) for c in column_indices],
            [(float(pos), sec) for pos, sec in sections])
    if not section_ids or not all(section_ids):
        # Letzte Rueckfallebene: der naechste Abschnitt links der Spalte.
        vorhanden = section_ids or [""] * len(column_indices)
        section_ids = []
        for c, bisher in zip(column_indices, vorhanden):
            applicable = [sec for pos, sec in sections if pos <= c]
            section_ids.append(bisher or (applicable[-1] if applicable else ""))

    first_col, last_col = column_indices[0], column_indices[-1]
    label_row = _label_row_above(grid, section_row, first_col, last_col)
    # Gruppen-/Untergruppenzeile stehen unmittelbar ueber der Bezeichnungszeile.
    anchor = label_row if label_row is not None else section_row
    group_rows = [r for r in (anchor - 2, anchor - 1) if r >= 0]

    filled_groups = [_ffill(grid, r, first_col, last_col) for r in group_rows]

    columns: list[ColumnHeader] = []
    for index, c in enumerate(column_indices):
        label = _value_at(grid, label_row, c) if label_row is not None else ""
        group, subgroup = "", ""
        if filled_groups:
            texts = [row[c] if c < len(row) else "" for row in filled_groups]
            texts = [t for t in texts if t and t != label]
            if texts:
                group = texts[0]
                subgroup = texts[1] if len(texts) > 1 else ""
        columns.append(
            ColumnHeader(
                index=index,
                group=group,
                subgroup=subgroup,
                label=label.strip(),
                section=section_ids[index],
                column_no=number_row[c].strip(),
                footnote_markers=extract_footnote_markers(label),
                x_center=float(c),
            )
        )

    column_of_index = {c: index for index, c in enumerate(column_indices)}
    note_cols = {c for c in column_indices
                 if normalize_address(section_ids[column_of_index[c]]) in {"4", "9"}}

    kopien = getattr(grid, "kopien", set())
    ausgeblendet = getattr(grid, "ausgeblendet", set())
    durchgestrichen = getattr(grid, "durchgestrichen", set())

    def wert(r: int, c: int) -> str:
        # Im Datenbereich zaehlt nur, was wirklich in der Zelle steht - eine
        # verbundene Zelle ueber zwei Zeilen ist EIN Wert, nicht zwei.
        return "" if (r, c) in kopien else _value_at(grid, r, c)

    rows: list[DataPointRow] = []
    for r in range(column_row + 1, len(grid)):
        left_texts = [wert(r, c) for c in range(0, first_col) if wert(r, c)]
        cells: list[Cell] = []
        for c in column_indices:
            raw = wert(r, c)
            if not raw:
                continue
            count = parse_count(raw)
            if c in note_cols:
                count = None
            cells.append(Cell(column_index=column_of_index[c], raw_value=raw,
                              count=count, note="" if count is not None else raw))
        if not left_texts and not cells:
            continue

        row_no = ""
        texts = list(left_texts)
        if texts and texts[0].isdigit():
            row_no = texts.pop(0)
        qualifier = wert(r, label_col)
        if qualifier in texts:
            texts.remove(qualifier)
        remark = " ".join(wert(r, c) for c in sorted(note_cols)).strip()
        klartext = " ".join(texts).strip()

        has_values = any(cell.count is not None for cell in cells)
        # Nur die Beschriftung bewerten - die Zeilennummer allein macht
        # aus einer leeren Zeile keinen Datenpunkt.
        kind, reason = classify_row(klartext, has_values, remark)
        if kind is RowKind.DATEN and _ist_wiederholter_kopf(left_texts, cells, numbers):
            # Wiederholte "Abschnitt"/"Spalte"-Zeile (Liste ueber mehrere
            # Druckseiten). Die Beschriftung steht in der Qualifikationsspalte
            # und waere oben schon entfernt - ohne diese Pruefung zaehlten die
            # Spaltennummern 1, 2, 3 ... als Funktionen eines Datenpunkts.
            kind, reason = RowKind.KOPF, "Wiederholte Kopfzeile"
        if kind is RowKind.FUSSBEREICH:
            # Unterhalb der Tabelle beginnt der Fussbereich - ab hier nichts mehr
            break
        if kind is RowKind.DATEN and (r in ausgeblendet or r in durchgestrichen):
            kind = RowKind.ENTFALLEN
            reason = ("In Excel ausgeblendete Zeile - entfallener Datenpunkt, wird nicht gezaehlt"
                      if r in ausgeblendet else
                      "Durchgestrichene Zeile - entfallener Datenpunkt, wird nicht gezaehlt")
        rows.append(
            DataPointRow(
                row_no=row_no,
                klartext=klartext,
                qualifier=qualifier,
                remark=remark,
                cells=cells,
                page_index=sheet_index,
                kind=kind,
                exclusion_reason=reason,
                confidence=1.0,
            )
        )

    rows = markiere_unbeschriftete_summenzeilen(drop_after_footer(rows))
    texts = [" ".join(value for value in row if value) for row in grid[:max(section_row + 2, 12)]]
    return RawTable(
        page_index=sheet_index,
        mode=ExtractionMode.EXCEL,
        columns=columns,
        rows=rows,
        footnotes=_extract_footnotes(grid, header_row=min(group_rows) if group_rows else section_row),
        metadata=_extract_metadata(grid),
        texts=texts + [" ".join(c.label for c in columns)],
    )


def extract_excel(path: str | Path) -> list[RawTable]:
    """Liefert je Arbeitsblatt eine RawTable (Blaetter ohne Raster entfallen)."""
    tables: list[RawTable] = []
    for index, grid in enumerate(read_grid(path)):
        try:
            table = _extract_sheet(grid, index)
        except Exception as exc:  # pragma: no cover - defensiv
            log.warning("Blatt %d in %s nicht lesbar: %s", index, path, exc)
            continue
        if table is not None:
            tables.append(table)
    return tables
