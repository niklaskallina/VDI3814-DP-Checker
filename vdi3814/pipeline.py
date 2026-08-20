"""Import-Pipeline: Datei -> DocumentResult.

Waehlt je Datei/Seite automatisch das genaueste verfuegbare Verfahren:

    .xlsx/.xls          -> Excel-Extraktion (exakt)
    PDF mit Textebene   -> Geometrie der Textebene (exakt)
    Scan / PNG / JPEG   -> lokales Vision-Language-Modell (+ OCR)

Erst danach wird die VDI-Fassung bestimmt und die Spalten dem Profil
zugeordnet. Seiten ohne Tabellenraster (Regelschemata) werden protokolliert
und uebersprungen.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Callable, Iterable

from .config import SETTINGS
from .extract import extract_excel, extract_pdf_text, is_excel, pdf_has_text_layer
from .extract.pdftext_extractor import classify_page_text
from .extract.base import ExtractionMode, RawTable
from .extract.vision_extractor import classify_page, extract_vision
from .ingest.loader import SUPPORTED_SUFFIXES, file_hash, load_pages
from .models import (
    Cell,
    ColumnHeader,
    DataPointRow,
    DocumentMetadata,
    DocumentResult,
    Footnote,
    PageClassification,
    PageKind,
    PageResult,
    RowKind,
    SumCheck,
)
from .profiles_loader import Profile, assign_footnotes, detect_profile, enrich_columns, match_columns
from .vision import ocr

log = logging.getLogger(__name__)

ProgressCallback = Callable[[str], None]


# --------------------------------------------------------------------------
# Zusammenfuehren der Seiten-/Blattergebnisse
# --------------------------------------------------------------------------

def _column_identity(column: ColumnHeader) -> str:
    """Stabiler Schluessel, um dieselbe Spalte ueber mehrere Seiten zu erkennen."""
    if column.column_key:
        return f"key:{column.column_key}"
    if column.address:
        return f"adr:{column.address}"
    return f"lbl:{column.display().lower()}"


def _merge_tables(tables: list[RawTable], profile: Profile | None) -> tuple[
        list[ColumnHeader], list[DataPointRow], list[Footnote], DocumentMetadata]:
    columns: list[ColumnHeader] = []
    index_of: dict[str, int] = {}
    rows: list[DataPointRow] = []
    footnotes: dict[str, Footnote] = {}
    metadata = DocumentMetadata()

    for table in tables:
        if profile is not None:
            match_columns(table.columns, profile)
            enrich_columns(table.columns, profile)
        mapping: dict[int, int] = {}
        for column in table.columns:
            identity = _column_identity(column)
            if identity not in index_of:
                merged = ColumnHeader(
                    index=len(columns),
                    group=column.group,
                    subgroup=column.subgroup,
                    label=column.label,
                    abbrev=column.abbrev,
                    section=column.section,
                    column_no=column.column_no,
                    footnote_markers=list(column.footnote_markers),
                    column_key=column.column_key,
                    match_score=column.match_score,
                    is_note_column=column.is_note_column,
                )
                index_of[identity] = merged.index
                columns.append(merged)
            else:
                merged = columns[index_of[identity]]
                if not merged.label and column.label:
                    merged.label = column.label
            mapping[column.index] = index_of[identity]

        for row in table.rows:
            row.cells = [
                Cell(column_index=mapping.get(cell.column_index, cell.column_index),
                     raw_value=cell.raw_value, count=cell.count, note=cell.note,
                     bbox=cell.bbox)      # Fundstelle mitnehmen, sie ist der Nachweis
                for cell in row.cells
                if cell.column_index in mapping
            ]
            rows.append(row)

        for footnote in table.footnotes:
            if footnote.marker not in footnotes:
                footnotes[footnote.marker] = footnote
        metadata.merge(table.metadata)

    # Kanonische Fussnotentexte des Profils ergaenzen bzw. bevorzugen
    if profile is not None:
        for marker, text in profile.footnotes:
            if text:
                footnotes[marker] = Footnote(marker=marker, text=text)

    ordered = [footnotes[key] for key in sorted(footnotes, key=lambda m: (len(m), m))]
    return columns, rows, ordered, metadata


def _validate_sums(result: DocumentResult) -> list[SumCheck]:
    """Vergleicht die eigene Aufsummierung mit der Summenzeile des Dokuments."""
    reported: dict[int, float] = {}
    for row in result.rows:
        if row.kind is not RowKind.SUMME:
            continue        # Uebertraege NICHT als Kontrollsumme verwenden
        for cell in row.cells:
            if cell.count is not None:
                reported[cell.column_index] = reported.get(cell.column_index, 0.0) + cell.count

    computed = result.totals()
    checks: list[SumCheck] = []
    for index in sorted(set(computed) | set(reported)):
        own = computed.get(index, 0.0)
        given = reported.get(index)
        checks.append(SumCheck(column_index=index, reported=given, computed=own,
                               matches=given is None or abs(given - own) < 1e-6))
    return checks


# --------------------------------------------------------------------------
# Seitenverarbeitung
# --------------------------------------------------------------------------

def _process_pdf(path: Path, backend, settings, progress: ProgressCallback) -> tuple[list[RawTable], list[PageResult]]:
    tables: list[RawTable] = []
    pages: list[PageResult] = []
    text_layer = pdf_has_text_layer(path)

    import pymupdf

    with pymupdf.open(path) as doc:
        page_count = doc.page_count
        seitengroessen = [(page.rect.width, page.rect.height) for page in doc]

    for page_index in range(page_count):
        table = None
        if text_layer:
            try:
                table = extract_pdf_text(path, page_index)
            except Exception as exc:                      # pragma: no cover - defensiv
                log.warning("Textebene von %s S.%d nicht lesbar: %s", path.name, page_index, exc)
        if table is not None:
            progress(f"{path.name} Seite {page_index + 1}: Tabelle ueber PDF-Textebene gelesen")
            tables.append(table)
            breite, hoehe = seitengroessen[page_index]
            pages.append(PageResult(
                page_index=page_index,
                classification=PageClassification(PageKind.FUNKTIONSLISTE, 1.0,
                                                  "Abschnitt-/Spalte-Raster in der Textebene"),
                width=breite, height=hoehe,
            ))
            continue

        # Kein Tabellenraster gefunden: erst ohne Modell klassifizieren.
        # Ein sicher erkanntes Regelschema muss gar nicht erst gerendert werden.
        if text_layer:
            kind_name, confidence, reason = classify_page_text(path, page_index)
            if kind_name in ("regelschema", "sonstiges") and confidence >= 0.7:
                progress(f"{path.name} Seite {page_index + 1}: uebersprungen ({kind_name})")
                pages.append(PageResult(
                    page_index=page_index,
                    classification=PageClassification(
                        PageKind.SCHEMA if kind_name == "regelschema" else PageKind.SONSTIGES,
                        confidence, reason),
                ))
                continue

        # Sonst Bild rendern und das Vision-Modell fragen
        page_image = None
        for candidate in load_pages(path, settings.render_dpi):
            if candidate.page_index == page_index:
                page_image = candidate.image
                break
        if page_image is None:                            # pragma: no cover
            pages.append(PageResult(page_index=page_index,
                                    classification=PageClassification(PageKind.FEHLER, 0.0,
                                                                      "Seite konnte nicht gerendert werden")))
            continue
        table, page = _process_image(page_image, page_index, backend, settings, progress,
                                     label=f"{path.name} Seite {page_index + 1}")
        page.width, page.height = seitengroessen[page_index]
        pages.append(page)
        if table is not None:
            tables.append(table)
    return tables, pages


def _process_image(image, page_index: int, backend, settings, progress: ProgressCallback,
                   label: str) -> tuple[RawTable | None, PageResult]:
    if backend is None:
        return None, PageResult(page_index=page_index,
                                classification=PageClassification(
                                    PageKind.FEHLER, 0.0,
                                    "Kein Vision-Modell verfuegbar - Bildseite konnte nicht ausgewertet werden"))

    classification = classify_page(backend, image)
    if not classification.is_table:
        progress(f"{label}: uebersprungen ({classification.kind.value})")
        return None, PageResult(page_index=page_index, classification=classification)

    progress(f"{label}: Tabelle wird per Vision-Modell gelesen")
    table = extract_vision(backend, image, page_index, settings)
    if not table.columns:
        classification = PageClassification(PageKind.FEHLER, classification.confidence,
                                            "; ".join(table.warnings) or "Keine Spalten erkannt")
        return None, PageResult(page_index=page_index, classification=classification)
    return table, PageResult(page_index=page_index, classification=classification)


# --------------------------------------------------------------------------
# Oeffentliche Schnittstelle
# --------------------------------------------------------------------------

def process_file(path: str | Path, backend=None, settings=SETTINGS,
                 progress: ProgressCallback | None = None) -> DocumentResult:
    """Verarbeitet eine Datei vollstaendig und liefert das Extraktionsergebnis."""
    path = Path(path)
    progress = progress or (lambda message: None)
    result = DocumentResult(source_path=str(path), file_name=path.name)

    if path.suffix.lower() not in SUPPORTED_SUFFIXES | {".xlsx", ".xlsm", ".xltx", ".xls"}:
        result.warnings.append(f"Nicht unterstuetzter Dateityp: {path.suffix}")
        return result

    try:
        result.file_hash = file_hash(path)
    except OSError as exc:
        result.warnings.append(f"Datei nicht lesbar: {exc}")
        return result

    tables: list[RawTable] = []
    try:
        if is_excel(path):
            tables = extract_excel(path)
            result.backend = ExtractionMode.EXCEL.value
            for table in tables:
                result.pages.append(PageResult(
                    page_index=table.page_index,
                    classification=PageClassification(PageKind.FUNKTIONSLISTE, 1.0, "Excel-Arbeitsblatt"),
                ))
            progress(f"{path.name}: {len(tables)} Arbeitsblatt/-blaetter gelesen")
        elif path.suffix.lower() == ".pdf":
            tables, pages = _process_pdf(path, backend, settings, progress)
            result.pages.extend(pages)
            modes = {table.mode.value for table in tables}
            result.backend = "+".join(sorted(modes)) if modes else ExtractionMode.VISION.value
        else:
            for page in load_pages(path):
                table, page_result = _process_image(page.image, page.page_index, backend, settings,
                                                    progress, label=page.file_name)
                result.pages.append(page_result)
                if table is not None:
                    tables.append(table)
            result.backend = ExtractionMode.VISION.value
    except Exception as exc:
        log.exception("Verarbeitung von %s fehlgeschlagen", path)
        result.warnings.append(f"Verarbeitung fehlgeschlagen: {exc}")
        return result

    for table in tables:
        result.warnings.extend(table.warnings)

    if not tables:
        result.warnings.append("Keine auswertbare Funktionsliste gefunden "
                               "(nur Schemata/Textseiten oder Erkennung fehlgeschlagen).")
        return result

    # --- VDI-Fassung bestimmen ---
    texts: list[str] = []
    for table in tables:
        texts.extend(table.texts)
        texts.extend(" ".join(filter(None, [c.group, c.subgroup, c.label])) for c in table.columns)
    profile, score, all_scores = detect_profile(texts)
    if profile is None and backend is not None:
        # Rueckfallebene: OCR-Volltext der ersten Bildseite mit einbeziehen
        texts.append(_ocr_hint(path, settings))
        profile, score, all_scores = detect_profile(texts)

    if profile is None:
        result.warnings.append(
            "VDI-Fassung konnte nicht sicher bestimmt werden "
            f"(Bewertung: {', '.join(f'{k}={v:.1f}' for k, v in all_scores.items())}). "
            "Bitte in der Vorschau manuell zuordnen."
        )
    else:
        result.profile_id = profile.id
        result.fassung = profile.fassung
        result.profile_score = round(score, 2)

    columns, rows, footnotes, metadata = _merge_tables(tables, profile)
    result.columns = columns
    result.rows = rows
    result.footnotes = footnotes
    result.metadata = metadata
    assign_footnotes(result.columns, result.footnotes)
    result.sum_checks = _validate_sums(result)

    unmatched = [c for c in result.columns if not c.column_key]
    if profile is not None and unmatched:
        result.warnings.append(
            f"{len(unmatched)} Spalte(n) konnten keinem Profileintrag zugeordnet werden: "
            + ", ".join(c.display() for c in unmatched[:5])
            + (" ..." if len(unmatched) > 5 else "")
        )
    for check in result.sum_checks:
        if not check.matches:
            column = result.column_by_index(check.column_index)
            result.warnings.append(
                f"Summenabweichung in Spalte '{column.display() if column else check.column_index}': "
                f"Dokument {check.reported:g}, eigene Summe {check.computed:g}"
            )
    return result


def _ocr_hint(path: Path, settings) -> str:
    """OCR-Volltext der ersten Seite - hilft bei der Fassungserkennung."""
    if not settings.use_ocr or not ocr.available():
        return ""
    try:
        for page in load_pages(path, settings.render_dpi):
            return ocr.page_text(page.image)
    except Exception:                                     # pragma: no cover
        return ""
    return ""


def process_files(paths: Iterable[str | Path], backend=None, settings=SETTINGS,
                  progress: ProgressCallback | None = None) -> list[DocumentResult]:
    """Batch-Verarbeitung beliebig vieler Dateien."""
    results = []
    for path in paths:
        results.append(process_file(path, backend=backend, settings=settings, progress=progress))
    return results
