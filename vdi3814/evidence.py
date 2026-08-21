"""Nachweis: woher stammt ein Wert und warum wurde er so bewertet?

Jede Zahl traegt ihre Fundstelle (Seite und Rechteck) mit sich. Damit kann die
Oberflaeche zu jedem Befund den Ausschnitt der Originalseite anzeigen und die
gefundene Stelle markieren - Abweichungen lassen sich so in Sekunden pruefen,
statt die PDF-Datei selbst zu durchsuchen.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path

from PIL import Image, ImageDraw

from . import db

log = logging.getLogger(__name__)

FINDING_COLORS = {
    "abweichung": (198, 40, 40),
    "ausgeschlossen": (239, 140, 0),
    "hinweis": (25, 118, 210),
    "datenpunkt": (46, 125, 50),
}


@dataclass
class Finding:
    """Ein Punkt, den der Anwender pruefen sollte."""

    art: str                    # abweichung | ausgeschlossen | hinweis
    titel: str
    erklaerung: str
    dokument_id: int
    datei: str
    seite: int = 0              # 0-basiert
    bbox: tuple[float, float, float, float] | None = None
    spalte: str = ""
    schwerpunkt: str = ""       # ASP/ISP, zu dem der Befund gehoert
    werte: dict[str, float] = field(default_factory=dict)

    @property
    def anzeige_seite(self) -> int:
        return self.seite + 1


# --------------------------------------------------------------------------
# Befunde sammeln
# --------------------------------------------------------------------------

def _schwerpunkt_der_seite(document: db.Document, page_index: int) -> str:
    """ASP/ISP der Seite - fuer Befunde, die keiner einzelnen Zeile gehoeren."""
    for page in document.pages:
        if page.page_index == page_index and page.schwerpunkt:
            return page.schwerpunkt
    return ""


def findings_for_document(document: db.Document) -> list[Finding]:
    """Alle pruefwuerdigen Punkte eines importierten Dokuments."""
    befunde: list[Finding] = []
    columns = {c.idx: c for c in document.columns}

    # 1. Abweichung zwischen eigener Summe und der Summenzeile - je Seite.
    #    Ein dokumentweiter Vergleich waere falsch: Plansaetze enthalten oft ein
    #    eigenes Summenblatt mit Summen je Anlage. Das fasst einen anderen
    #    Umfang zusammen (teils Anlagen aus anderen Dateien) und wuerde in
    #    jeder Spalte eine scheinbare Abweichung erzeugen.
    seiten: dict[int, dict[str, dict]] = {}
    for row in document.rows:
        eintrag = seiten.setdefault(row.page_index, {"eigen": {}, "gemeldet": {}, "zeile": {}})
        for value in row.values:
            if value.count is None:
                continue
            if row.kind == "daten":
                eintrag["eigen"][value.column_idx] = (
                    eintrag["eigen"].get(value.column_idx, 0.0) + value.count)
            elif row.kind == "summe":
                eintrag["gemeldet"][value.column_idx] = (
                    eintrag["gemeldet"].get(value.column_idx, 0.0) + value.count)
                eintrag["zeile"].setdefault(value.column_idx, row)

    for page_index in sorted(seiten):
        eigen = seiten[page_index]["eigen"]
        gemeldet = seiten[page_index]["gemeldet"]
        if not eigen or not gemeldet:
            continue          # Seite ohne Summenzeile oder reines Summenblatt
        for column_idx, wert in gemeldet.items():
            eigener_wert = eigen.get(column_idx, 0.0)
            if abs(wert - eigener_wert) < 1e-6:
                continue
            column = columns.get(column_idx)
            zeile = seiten[page_index]["zeile"].get(column_idx)
            bbox = db.bbox_from_text(zeile.bbox) if zeile else None
            befunde.append(Finding(
                art="abweichung",
                titel=f"Seite {page_index + 1}, Spalte "
                      f"{column.address if column else column_idx}: "
                      f"Liste {wert:g}, eigene Zaehlung {eigener_wert:g}",
                erklaerung=(
                    "Die Summenzeile dieser Seite weicht von der Summe der Einzelzeilen "
                    "derselben Seite ab. Haeufigste Ursachen: die Liste wurde nach dem "
                    "Ausfuellen geaendert, eine Zeile wurde uebersehen, oder ein Wert "
                    "steht zwischen zwei Spalten. Die Markierung zeigt die Summenzeile."
                ),
                dokument_id=document.id,
                datei=document.file_name,
                seite=zeile.page_index if zeile else page_index,
                bbox=bbox,
                spalte=column.label if column else "",
                schwerpunkt=_schwerpunkt_der_seite(document, page_index),
                werte={"Liste": wert, "eigene Zaehlung": eigener_wert},
            ))

    # 2. Nicht gezaehlte Zeilen mit Werten - hier lohnt der Blick am meisten.
    #    Eine Summenzeile, die zur eigenen Zaehlung passt, ist dagegen normal
    #    und wird nicht gemeldet (sonst geht der Blick fuers Wesentliche verloren).
    for row in document.rows:
        if row.kind in ("daten", "leer"):
            continue
        if row.kind == "summe":
            # Eine Summenzeile, die zur eigenen Zaehlung derselben Seite passt,
            # ist normal und wird nicht gemeldet.
            seite = seiten.get(row.page_index, {"eigen": {}, "gemeldet": {}})
            eigen_seite, gemeldet_seite = seite["eigen"], seite["gemeldet"]
            stimmt = eigen_seite and all(
                abs(gemeldet_seite.get(idx, 0.0) - eigen_seite.get(idx, 0.0)) < 1e-6
                for idx in gemeldet_seite
            )
            if stimmt:
                continue
        werte = [v for v in row.values if v.count is not None]
        if not werte:
            continue
        befunde.append(Finding(
            art="ausgeschlossen",
            titel=f"Zeile {row.row_no or '-'}: {row.klartext or '(ohne Beschriftung)'} "
                  f"({row.kind})",
            erklaerung=row.exclusion_reason or "Zeile wurde nicht als Datenpunkt gewertet.",
            dokument_id=document.id,
            datei=document.file_name,
            seite=row.page_index,
            bbox=db.bbox_from_text(row.bbox),
            schwerpunkt=row.schwerpunkt or _schwerpunkt_der_seite(document, row.page_index),
            werte={"Summe der Zeile": sum(v.count or 0.0 for v in werte)},
        ))

    # 3. Allgemeine Hinweise aus dem Import
    for hinweis in document.warnings:
        befunde.append(Finding(
            art="hinweis",
            titel=hinweis[:90],
            erklaerung=hinweis,
            dokument_id=document.id,
            datei=document.file_name,
        ))
    return befunde


def all_findings(session) -> list[Finding]:
    from sqlalchemy import select

    befunde: list[Finding] = []
    for document in session.scalars(select(db.Document)):
        befunde.extend(findings_for_document(document))
    reihenfolge = {"abweichung": 0, "ausgeschlossen": 1, "hinweis": 2}
    return sorted(befunde, key=lambda f: (reihenfolge.get(f.art, 9), f.datei))


# --------------------------------------------------------------------------
# Seitenausschnitt mit Markierung
# --------------------------------------------------------------------------

def source_file(document: db.Document) -> Path | None:
    """Die im Projekt abgelegte Kopie der Originaldatei."""
    for kandidat in (document.stored_path, document.source_path):
        if kandidat and Path(kandidat).is_file():
            return Path(kandidat)
    return None


def render_page(document: db.Document, page_index: int,
                highlights: list[tuple[tuple[float, float, float, float], str]] | None = None,
                dpi: int = 130, zoom_to: tuple[float, float, float, float] | None = None,
                margin: float = 60.0) -> Image.Image | None:
    """Rendert eine Seite der Quelldatei und markiert die Fundstellen.

    highlights: Liste aus (Rechteck in Quellkoordinaten, Art)
    zoom_to:    optionaler Ausschnitt, um den herum zugeschnitten wird
    """
    pfad = source_file(document)
    if pfad is None:
        return None

    seite = None
    if pfad.suffix.lower() == ".pdf":
        import pymupdf

        try:
            with pymupdf.open(pfad) as doc:
                if page_index >= doc.page_count:
                    return None
                page = doc[page_index]
                skalierung = dpi / 72.0
                pix = page.get_pixmap(matrix=pymupdf.Matrix(skalierung, skalierung), alpha=False)
                seite = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        except Exception as exc:            # pragma: no cover - defekte Datei
            log.warning("Seite konnte nicht gerendert werden: %s", exc)
            return None
    elif pfad.suffix.lower() in {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"}:
        seite = Image.open(pfad).convert("RGB")
        skalierung = 1.0
    else:
        return None                          # Excel-Quellen haben kein Seitenbild

    zeichnung = ImageDraw.Draw(seite, "RGBA")
    for bbox, art in highlights or []:
        if not bbox:
            continue
        farbe = FINDING_COLORS.get(art, FINDING_COLORS["datenpunkt"])
        rechteck = [bbox[0] * skalierung - 2, bbox[1] * skalierung - 2,
                    bbox[2] * skalierung + 2, bbox[3] * skalierung + 2]
        zeichnung.rectangle(rechteck, outline=farbe, width=3)
        zeichnung.rectangle(rechteck, fill=(*farbe, 40))

    if zoom_to:
        x0 = max(0, int((zoom_to[0] - margin) * skalierung))
        y0 = max(0, int((zoom_to[1] - margin * 0.6) * skalierung))
        x1 = min(seite.width, int((zoom_to[2] + margin) * skalierung))
        y1 = min(seite.height, int((zoom_to[3] + margin * 0.6) * skalierung))
        if x1 > x0 and y1 > y0:
            seite = seite.crop((x0, y0, x1, y1))
    return seite


def render_finding(document: db.Document, finding: Finding, ganze_seite: bool = False):
    """Seitenbild zu einem Befund - markiert und auf die Fundstelle gezoomt."""
    if finding.bbox is None:
        return render_page(document, finding.seite)
    return render_page(
        document, finding.seite,
        highlights=[(finding.bbox, finding.art)],
        zoom_to=None if ganze_seite else finding.bbox,
    )
