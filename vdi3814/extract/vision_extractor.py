"""Extraktion aus Bildern und gescannten PDFs ueber ein lokales Vision-Modell.

Ablauf je Seite:
1. Klassifikation      - Funktionsliste oder Regelschema?
2. Kopfbereich         - Spaltenstruktur (Abschnitt/Spalte + Bezeichnungen) und Fussnoten
3. Zeilenstreifen      - die Seite wird in horizontale Streifen zerlegt; jedem
                         Streifen wird der Tabellenkopf vorangestellt, damit das
                         Modell die Spaltenzuordnung sieht
4. Fussbereich         - Projektmetadaten

Das Banding ist der Kern der Robustheit: eine komplette Funktionsliste mit
50+ Spalten und 30 Zeilen ist fuer ein lokales Modell in einem Zug nicht
zuverlaessig lesbar, ein Streifen mit 8-12 Zeilen dagegen schon.
"""

from __future__ import annotations

import logging

from PIL import Image

from ..config import SETTINGS
from ..ingest.preprocess import crop_band, crop_footer, crop_header, prepare_for_vision, stack_with_header
from ..models import (Cell, ColumnHeader, DataPointRow, DocumentMetadata, Footnote,
                      PageClassification, PageKind, RowKind)
from ..textutil import extract_footnote_markers, normalize, parse_count
from ..vision import prompts
from ..vision.base import VisionError
from .base import ExtractionMode, RawTable, drop_after_footer
from .rowfilter import classify_row

log = logging.getLogger(__name__)

_KIND_MAP = {
    "funktionsliste": PageKind.FUNKTIONSLISTE,
    "regelschema": PageKind.SCHEMA,
    "schema": PageKind.SCHEMA,
    "sonstiges": PageKind.SONSTIGES,
}


def classify_page(backend, image: Image.Image) -> PageClassification:
    """Bestimmt, ob die Seite eine Funktionsliste oder ein Regelschema ist."""
    try:
        answer = backend.ask_json(prepare_for_vision(image), prompts.CLASSIFY, purpose="classify")
    except VisionError as exc:
        return PageClassification(kind=PageKind.FEHLER, confidence=0.0, reason=str(exc))
    kind = _KIND_MAP.get(normalize(answer.get("typ", "")).replace(" ", ""), PageKind.SONSTIGES)
    try:
        confidence = float(answer.get("konfidenz", 0.0))
    except (TypeError, ValueError):
        confidence = 0.0
    return PageClassification(kind=kind, confidence=confidence, reason=str(answer.get("begruendung", "")))


def _columns_from_answer(answer: dict) -> tuple[list[ColumnHeader], list[Footnote]]:
    columns: list[ColumnHeader] = []
    for index, raw in enumerate(answer.get("spalten", []) or []):
        label = str(raw.get("bezeichnung", "")).strip()
        markers = [str(m) for m in raw.get("fussnoten", []) or []] or extract_footnote_markers(label)
        columns.append(
            ColumnHeader(
                index=int(raw.get("index", index)),
                group=str(raw.get("gruppe", "")).strip(),
                subgroup=str(raw.get("untergruppe", "")).strip(),
                label=label,
                section=str(raw.get("abschnitt", "")).strip(),
                column_no=str(raw.get("spalte_nr", "")).strip(),
                footnote_markers=markers,
            )
        )
    footnotes = [
        Footnote(marker=str(f.get("marker", "")).strip(), text=str(f.get("text", "")).strip())
        for f in answer.get("fussnoten", []) or []
        if str(f.get("marker", "")).strip()
    ]
    return columns, footnotes


def _rows_from_answer(answer: dict, page_index: int, column_count: int) -> list[DataPointRow]:
    rows: list[DataPointRow] = []
    for raw in answer.get("zeilen", []) or []:
        cells: list[Cell] = []
        for key, value in (raw.get("werte") or {}).items():
            try:
                column_index = int(str(key).strip())
            except ValueError:
                continue
            if not 0 <= column_index < max(column_count, 1):
                continue
            text = str(value).strip()
            if not text:
                continue
            count = parse_count(text)
            cells.append(Cell(column_index=column_index, raw_value=text, count=count,
                              note="" if count is not None else text))
        klartext = str(raw.get("klartext", "")).strip()
        if not klartext and not cells and not raw.get("bas"):
            continue
        has_values = any(cell.count is not None for cell in cells)
        kind, reason = classify_row(klartext, has_values, str(raw.get("bemerkung", "")))
        if bool(raw.get("summenzeile")) and kind is RowKind.DATEN:
            kind, reason = RowKind.SUMME, "Vom Modell als Summenzeile gemeldet"
        rows.append(
            DataPointRow(
                row_no=str(raw.get("zeile_nr", "")).strip(),
                bas=str(raw.get("bas", "")).strip(),
                klartext=klartext,
                qualifier=str(raw.get("qualifikation", "")).strip(),
                remark=str(raw.get("bemerkung", "")).strip(),
                cells=cells,
                page_index=page_index,
                kind=kind,
                exclusion_reason=reason,
                confidence=0.6,          # Modellwerte gelten als pruefbeduerftig
            )
        )
    return rows


def _merge_rows(existing: list[DataPointRow], new: list[DataPointRow]) -> list[DataPointRow]:
    """Fuehrt die Streifenergebnisse zusammen (Ueberlappung erzeugt Dubletten)."""
    def key(row: DataPointRow) -> str:
        return f"{row.row_no}|{normalize(row.klartext)}|{normalize(row.bas)}"

    seen = {key(row): row for row in existing}
    for row in new:
        row_key = key(row)
        current = seen.get(row_key)
        if current is None:
            existing.append(row)
            seen[row_key] = row
        elif len(row.cells) > len(current.cells):
            # Der Streifen mit mehr gelesenen Werten gewinnt
            current.cells = row.cells
            current.remark = current.remark or row.remark
    return existing


def _metadata_from_answer(answer: dict) -> DocumentMetadata:
    metadata = DocumentMetadata()
    for field_name in vars(metadata):
        value = answer.get(field_name)
        if isinstance(value, str) and value.strip():
            setattr(metadata, field_name, value.strip())
    return metadata


def extract_vision(backend, image: Image.Image, page_index: int,
                   settings=SETTINGS) -> RawTable:
    """Liest eine Bildseite ueber das lokale Vision-Modell aus."""
    prepared = prepare_for_vision(image, settings)
    warnings: list[str] = []

    header_image = crop_header(prepared, settings.header_fraction)
    try:
        header_answer = backend.ask_json(header_image, prompts.HEADER, purpose="header")
    except VisionError as exc:
        return RawTable(page_index=page_index, mode=ExtractionMode.VISION,
                        warnings=[f"Kopfbereich nicht lesbar: {exc}"])
    columns, footnotes = _columns_from_answer(header_answer)
    if not columns:
        return RawTable(page_index=page_index, mode=ExtractionMode.VISION,
                        warnings=["Es konnten keine Spaltenueberschriften erkannt werden."])

    rows: list[DataPointRow] = []
    for band_index in range(settings.band_count):
        band = crop_band(prepared, band_index, settings.band_count,
                         overlap=settings.band_overlap, body_top=settings.header_fraction)
        stacked = stack_with_header(header_image, band)
        try:
            answer = backend.ask_json(stacked, prompts.ROWS, purpose="rows")
        except VisionError as exc:
            warnings.append(f"Streifen {band_index + 1}/{settings.band_count} nicht lesbar: {exc}")
            continue
        _merge_rows(rows, _rows_from_answer(answer, page_index, len(columns)))

    try:
        metadata_answer = backend.ask_json(crop_footer(prepared, settings.footer_fraction),
                                           prompts.METADATA, purpose="metadata")
        metadata = _metadata_from_answer(metadata_answer)
    except VisionError as exc:
        metadata = DocumentMetadata()
        warnings.append(f"Fussbereich nicht lesbar: {exc}")

    texts = [" ".join(filter(None, [c.group, c.subgroup, c.label])) for c in columns]
    texts.append(" ".join(f"{key} {value}" for key, value in vars(metadata).items() if value))

    return RawTable(
        page_index=page_index,
        mode=ExtractionMode.VISION,
        columns=columns,
        rows=drop_after_footer(rows),
        footnotes=footnotes,
        metadata=metadata,
        texts=texts,
        warnings=warnings,
    )
