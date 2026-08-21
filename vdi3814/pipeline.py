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
from .extract.ocr_extractor import available as ocr_available
from .extract.ocr_extractor import classify_image, extract_ocr
from .extract.pdftext_extractor import classify_page_text
from .extract.base import ExtractionMode, RawTable
from .extract.vision_extractor import classify_page, extract_vision
from .ingest.loader import SUPPORTED_SUFFIXES, file_hash, load_pages, render_pdf_page
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
from . import schwerpunkt as schwerpunkte
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


def validate_sums(result: DocumentResult) -> list[SumCheck]:
    """Prueft je Seite: Summenzeile der Seite gegen die Datenzeilen derselben Seite.

    Nur so ist der Vergleich aussagekraeftig. Reine Summenblaetter (Seiten, auf
    denen ausschliesslich Summen je Anlage stehen) werden nicht geprueft - sie
    fassen einen anderen Umfang zusammen und wuerden sonst auf jeder Spalte
    eine scheinbare Abweichung erzeugen.

    Oeffentlich, weil die Oberflaeche die Pruefung nach einer Korrektur (z. B.
    dem Loeschen einer Zeile) neu rechnen muss.
    """
    seiten: dict[int, dict[str, dict[int, float]]] = {}
    for row in result.rows:
        eintrag = seiten.setdefault(row.page_index, {"summe": {}, "eigen": {}, "uebertrag": {}})
        if row.kind is RowKind.SUMME:
            ziel = eintrag["summe"]
        elif row.kind is RowKind.UEBERTRAG:
            # Der Uebertrag wird nicht gezaehlt, gehoert aber zur erwarteten
            # Seitensumme: die Summenzeile enthaelt Uebertrag + Seitenwerte.
            ziel = eintrag["uebertrag"]
        elif row.is_countable:
            ziel = eintrag["eigen"]
        else:
            continue        # Leerzeilen und Fussbereich nie werten
        for cell in row.cells:
            if cell.count is not None:
                ziel[cell.column_index] = ziel.get(cell.column_index, 0.0) + cell.count

    checks: list[SumCheck] = []
    for page_index in sorted(seiten):
        gemeldet = seiten[page_index]["summe"]
        eigen = dict(seiten[page_index]["eigen"])
        uebertrag = seiten[page_index]["uebertrag"]
        if not gemeldet or not eigen:
            # Ohne eigene Datenzeilen gibt es nichts zu pruefen: das ist ein
            # Summenblatt (nur Summen je Anlage, ggf. mit Uebertrag).
            continue
        for index, wert in uebertrag.items():
            eigen[index] = eigen.get(index, 0.0) + wert
        for index in sorted(set(eigen) | set(gemeldet)):
            own = eigen.get(index, 0.0)
            given = gemeldet.get(index)
            checks.append(SumCheck(column_index=index, reported=given, computed=own,
                                   matches=given is None or abs(given - own) < 1e-6,
                                   page_index=page_index))
    return checks


def _summenblatt_seiten(result: DocumentResult) -> set[int]:
    """Seiten, die ausschliesslich Summen bzw. Uebertraege enthalten."""
    seiten: dict[int, list] = {}
    for row in result.rows:
        seiten.setdefault(row.page_index, []).append(row)
    return {
        index for index, rows in seiten.items()
        if any(r.kind in (RowKind.SUMME, RowKind.UEBERTRAG) for r in rows)
        and not any(r.is_countable for r in rows)
    }


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

        # Sonst genau diese eine Seite rendern und auswerten
        try:
            page_image = render_pdf_page(path, page_index, settings.render_dpi)
        except Exception as exc:                          # pragma: no cover - defekte Datei
            log.warning("Seite %d von %s nicht rendbar: %s", page_index, path.name, exc)
            page_image = None
        if page_image is None:
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


_KIND_BY_NAME = {
    "funktionsliste": PageKind.FUNKTIONSLISTE,
    "regelschema": PageKind.SCHEMA,
    "sonstiges": PageKind.SONSTIGES,
}


def _process_image(image, page_index: int, backend, settings, progress: ProgressCallback,
                   label: str) -> tuple[RawTable | None, PageResult]:
    """Wertet eine Bildseite aus - zuerst mit lokalem OCR, ohne KI.

    Reihenfolge:
    1. OCR-Klassifikation: Funktionsliste oder Regelschema?
    2. Zelle-fuer-Zelle-OCR entlang des erkannten Tabellenrasters
    3. nur falls das misslingt und ausdruecklich gewuenscht: lokales Vision-Modell
    """
    if ocr_available():
        art, sicherheit, begruendung = classify_image(image)
        if art == "regelschema" and sicherheit >= 0.7:
            progress(f"{label}: uebersprungen (Regelschema)")
            return None, PageResult(
                page_index=page_index,
                classification=PageClassification(PageKind.SCHEMA, sicherheit, begruendung))

        progress(f"{label}: Seite wird per OCR gelesen (kein Modell noetig)")
        table = extract_ocr(image, page_index)
        if table is not None and table.columns:
            return table, PageResult(
                page_index=page_index,
                classification=PageClassification(PageKind.FUNKTIONSLISTE, 0.9,
                                                  "Tabellenraster im Scan vermessen"))
        if backend is None:
            grund = ("Im Scan wurde kein Abschnitt-/Spalte-Raster gefunden. "
                     f"Klassifikation: {art} ({begruendung}).")
            kind = _KIND_BY_NAME.get(art, PageKind.SONSTIGES)
            if kind is PageKind.FUNKTIONSLISTE:
                kind = PageKind.FEHLER
            progress(f"{label}: nicht auswertbar ({art})")
            return None, PageResult(page_index=page_index,
                                    classification=PageClassification(kind, sicherheit, grund))

    if backend is None:
        return None, PageResult(
            page_index=page_index,
            classification=PageClassification(
                PageKind.FEHLER, 0.0,
                "Bildseite konnte nicht ausgewertet werden: OCR ist nicht eingerichtet "
                "und es wurde kein Vision-Modell aktiviert."))

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
    _ordne_schwerpunkte(result, tables)
    result.sum_checks = validate_sums(result)

    unmatched = [c for c in result.columns if not c.column_key]
    if profile is not None and unmatched:
        result.warnings.append(
            f"{len(unmatched)} Spalte(n) konnten keinem Profileintrag zugeordnet werden: "
            + ", ".join(c.display() for c in unmatched[:5])
            + (" ..." if len(unmatched) > 5 else "")
        )
    # Reine Summenblaetter kenntlich machen: dort steht je Anlage eine Summe,
    # die bewusst nicht mitgezaehlt wird - sonst waere jede Anlage doppelt drin.
    summenblaetter = _summenblatt_seiten(result)
    for page in result.pages:
        if page.page_index in summenblaetter:
            page.classification = PageClassification(
                PageKind.SUMMENBLATT, 1.0,
                "Summenblatt: enthaelt nur Summen je Anlage - wird nicht gezaehlt",
            )
    if summenblaetter:
        result.warnings.append(
            f"{len(summenblaetter)} Summenblatt-Seite(n) erkannt "
            f"(Seite {', '.join(str(i + 1) for i in sorted(summenblaetter))}): "
            "enthalten nur Summen je Anlage und werden nicht mitgezaehlt."
        )

    # Summenabweichungen gebuendelt melden - eine Meldung je Seite statt je Spalte.
    abweichungen = [c for c in result.sum_checks if not c.matches]
    if abweichungen:
        seiten = sorted({c.page_index for c in abweichungen})
        result.warnings.append(
            f"{len(abweichungen)} Summenabweichung(en) auf {len(seiten)} Seite(n) "
            f"(Seite {', '.join(str(i + 1) for i in seiten[:8])}"
            f"{' ...' if len(seiten) > 8 else ''}). "
            "Einzelheiten im Reiter 'Nachweis & Differenzen'."
        )
    return result


def _ordne_schwerpunkte(result: DocumentResult, tables: list[RawTable]) -> None:
    """Ordnet jede Zeile ihrem Automations-/Informationsschwerpunkt zu.

    Ein Plansatz enthaelt haeufig mehrere Schwerpunkte (ASP01, ASP02, ...) -
    je Seite einen oder, auf Uebersichtsblaettern, mehrere untereinander.
    Massgeblich ist deshalb zuerst der Kopf-/Fussbereich der jeweiligen Seite,
    dann die Angabe in der Zeile selbst (siehe schwerpunkt.zuordnen).
    """
    vorgaben = {}
    for table in tables:
        gefunden = schwerpunkte.aus_metadaten(table.metadata)
        if gefunden is not None:
            vorgaben[table.page_index] = gefunden
    schwerpunkte.zuordnen(result.rows, vorgaben,
                          schwerpunkte.aus_metadaten(result.metadata))

    je_seite = schwerpunkte.je_seite(result.rows)
    for page in result.pages:
        vorgabe = vorgaben.get(page.page_index)
        page.schwerpunkt = je_seite.get(page.page_index) or (vorgabe.kennung if vorgabe else "")

    # Mehrere Schwerpunkte je Datei sind der Normalfall und deshalb kein
    # Hinweis wert - gar keiner dagegen schon: dann fehlt der Auswertung die
    # Zuordnung, und der Anwender muss sie nachtragen koennen.
    if not any(row.schwerpunkt for row in result.counted_rows()):
        result.warnings.append(
            "Kein Automations-/Informationsschwerpunkt (ASP/ISP) gefunden. "
            "Die Mengen lassen sich damit keinem Schwerpunkt zuordnen - "
            "Angabe bei Bedarf im Reiter 'Prüfen & korrigieren' nachtragen."
        )


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
