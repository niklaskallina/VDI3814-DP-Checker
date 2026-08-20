"""Erzeugt den Beispiel-/Testdatensatz.

Erstellt zwei Funktionslisten-PDFs (alte und neue Fassung) im Layout der
VDI-Vorlagen - inkl. Kopfzeilen "Abschnitt"/"Spalte", Fussnoten, Datenzeilen,
Summenzeile und Fussbereich - sowie ein Regelschema als zusaetzliche Seite,
damit sich das Ueberspringen von Schemaseiten demonstrieren laesst.

Aufruf:  python sample_data/make_samples.py
"""

from __future__ import annotations

import sys
from pathlib import Path

import pymupdf

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from vdi3814.profiles_loader import get_profile   # noqa: E402

OUT_DIR = Path(__file__).resolve().parent

FONT = "helv"
CELL_WIDTH = 11.0
LABEL_WIDTH = 150.0
LEFT = 30.0
ROW_HEIGHT = 14.0


def _text(page, x, y, value, size=6.0, rotate=0):
    page.insert_text((x, y), str(value), fontname=FONT, fontsize=size, rotate=rotate)


def _draw_list(page, profile, rows, meta, footnotes, title):
    columns = list(profile.columns)
    note_column = profile.note_columns[0] if profile.note_columns else None

    # --- Fussnoten oben ---
    y = 24.0
    for marker, text in footnotes:
        _text(page, LEFT, y, f"{marker}) {text}", size=5.5)
        y += 8.0

    header_bottom = 210.0
    grid_left = LEFT + LABEL_WIDTH

    # --- Titel/Kopf ---
    _text(page, LEFT, y + 10, title, size=9)
    _text(page, LEFT, y + 22, f" Gewerk: {meta['gewerk']}", size=6.5)
    _text(page, LEFT, y + 32, f" Anlage: {meta['anlage']}", size=6.5)

    # --- Senkrechte Spaltenueberschriften ---
    for index, column in enumerate(columns):
        x = grid_left + index * CELL_WIDTH + 3
        _text(page, x, header_bottom - 6, column.label, size=5.0, rotate=90)

    # --- Gruppenzeile ---
    seen: dict[str, float] = {}
    for index, column in enumerate(columns):
        seen.setdefault(column.group, grid_left + index * CELL_WIDTH)
    for group, x in seen.items():
        _text(page, x, 96.0, group, size=6.0)

    # --- Kopfzeilen "Abschnitt" und "Spalte" ---
    section_row_y = header_bottom + 8
    column_row_y = header_bottom + 20
    _text(page, LEFT + 80, section_row_y, "Abschnitt", size=6.0)
    _text(page, LEFT + 80, column_row_y, "Spalte", size=6.0)
    _text(page, LEFT, column_row_y, " Benennung", size=6.0)

    section_starts: dict[str, list[float]] = {}
    for index, column in enumerate(columns):
        section = column.address.rsplit(".", 1)[0]
        number = column.address.rsplit(".", 1)[1]
        x = grid_left + index * CELL_WIDTH
        section_starts.setdefault(section, []).append(x)
        _text(page, x + 2, column_row_y, number, size=6.0)
    for section, positions in section_starts.items():
        centre = (positions[0] + positions[-1]) / 2
        _text(page, centre, section_row_y, section, size=6.0)

    if note_column is not None:
        note_x = grid_left + len(columns) * CELL_WIDTH + 12
        _text(page, note_x, column_row_y, "Bemerkung", size=6.0)
    else:                                     # pragma: no cover
        note_x = None

    # --- Datenzeilen ---
    y = column_row_y + ROW_HEIGHT
    totals: dict[str, float] = {}
    for row_no, (label, values, remark) in enumerate(rows, start=1):
        _text(page, LEFT, y, str(row_no), size=6.0)
        _text(page, LEFT + 12, y, label, size=6.5)
        for address, value in values.items():
            index = next(i for i, c in enumerate(columns) if c.address == address)
            _text(page, grid_left + index * CELL_WIDTH + 3, y, value, size=6.0)
            totals[address] = totals.get(address, 0.0) + float(value)
        if remark and note_x:
            _text(page, note_x, y, remark, size=5.5)
        y += ROW_HEIGHT

    # --- Summenzeile ---
    y += 6
    _text(page, LEFT + 12, y, "Summe Funktionen", size=6.5)
    for address, total in totals.items():
        index = next(i for i, c in enumerate(columns) if c.address == address)
        _text(page, grid_left + index * CELL_WIDTH + 3, y, f"{total:g}", size=6.0)

    # --- Fussbereich ---
    y += 24
    _text(page, LEFT, y, f" Auftraggeber/Projekt: {meta['projekt']}", size=6.0)
    _text(page, LEFT + 260, y, f" Planersteller: {meta['planersteller']}", size=6.0)
    _text(page, LEFT + 480, y, f" Planstand: {meta['planstand']}", size=6.0)
    _text(page, LEFT, y + 10, f" Blatt Nr. {meta['blatt_nr']}", size=6.0)
    _text(page, LEFT + 90, y + 10, f" von {meta['blatt_von']}", size=6.0)
    if meta.get("protokoll"):
        _text(page, LEFT + 260, y + 10, f" Datenkommunikationsprotokoll: {meta['protokoll']}", size=6.0)


def _draw_schema(page):
    """Ein einfaches Regelschema - darf vom Tool NICHT ausgewertet werden."""
    _text(page, 60, 50, "Regelschema Lueftungsanlage RLT 01", size=12)
    _text(page, 60, 70, "Prinzipschaltbild - keine Funktionsliste", size=8)
    shape = page.new_shape()
    shape.draw_rect(pymupdf.Rect(80, 110, 700, 300))
    for offset in range(0, 620, 40):
        shape.draw_line(pymupdf.Point(80 + offset, 110), pymupdf.Point(80 + offset, 300))
    for offset in range(0, 200, 20):
        shape.draw_line(pymupdf.Point(80, 110 + offset), pymupdf.Point(700, 110 + offset))
    shape.draw_circle(pymupdf.Point(200, 205), 30)
    shape.draw_circle(pymupdf.Point(400, 205), 30)
    shape.finish(width=0.6)
    shape.commit()


def build_alt(path: Path) -> None:
    profile = get_profile("vdi3814_alt")
    rows = [
        ("Aussentemperatur", {"1.5": "1", "8.1": "1"}, ""),
        ("RLT 01 Zuluftventilator", {"1.1": "2", "1.3": "2", "3.3": "1", "3.5": "1",
                                     "4.2": "1", "8.1": "1", "8.2": "1"}, "Dauerbefehl 2 BA"),
        ("RLT 01 Zulufttemperatur", {"1.5": "1", "3.1": "1", "5.2": "1", "8.1": "1"}, ""),
        ("RLT 01 Heizventil", {"1.2": "1", "5.4": "1", "8.1": "1"}, ""),
        ("RLT 01 Frostschutzthermostat", {"1.3": "1", "4.5": "1", "3.6": "1", "8.3": "1"}, "Sicherheitskette"),
        ("RLT 01 Aussenluftklappe", {"1.1": "2", "1.3": "1", "8.1": "1"}, "3-Punkt = 2 x 2-Punkt"),
        ("RLT 01 Filterueberwachung", {"1.3": "1", "3.6": "1", "8.4": "1"}, ""),
        ("RLT 01 Zeitprogramm", {"6.4": "1", "7.3": "1"}, ""),
    ]
    meta = {
        "gewerk": "MSR / Gebaeudeautomation",
        "anlage": "RLT 01 Lueftung Buerotrakt",
        "projekt": "Neubau Verwaltungsgebaeude Musterstadt",
        "planersteller": "Musterplanung GmbH",
        "planstand": "2024-05-14",
        "blatt_nr": "1",
        "blatt_von": "1",
        "protokoll": "",
    }
    document = pymupdf.open()
    page = document.new_page(width=842, height=595)
    _draw_list(page, get_profile("vdi3814_alt"), rows, meta,
               [(marker, text) for marker, text in profile.footnotes[:4]],
               "GA-Funktionsliste VDI 3814 Blatt 1 (alte Fassung)")
    _draw_schema(document.new_page(width=842, height=595))
    document.save(path)
    document.close()


def build_neu(path: Path) -> None:
    profile = get_profile("vdi3814_neu")
    rows = [
        ("Automationseinrichtung ASE 1 *ASE01*DEV", {"0.1": "1", "3.1.1": "1", "3.1.2": "4"},
         "1.3=3.1.1=DEV"),
        ("Feiertagskalender *ASE0x*CAL", {"1.3.2": "1", "3.1.1": "1"}, ""),
        ("Meldungsklasse Alarm *ASE01*NC3", {"1.3.3": "1", "3.1.1": "1"}, ""),
        ("Zulufttemperatur *L01*ZUL*TEMP", {"1.1.1": "1", "2.2.1": "1", "2.3.1": "1", "3.1.2": "1"}, ""),
        ("Zuluftventilator *L01*ZUL*VENT", {"1.1.4": "1", "1.1.2": "2", "2.2.6": "1",
                                            "2.2.4": "1", "3.1.2": "1"}, ""),
        ("Heizventil *L01*HEIZ*VENT", {"1.1.3": "1", "2.3.3": "1", "3.1.1": "1"}, ""),
        ("Sonnenschutz Fassade Sued", {"1.2.2": "1", "2.6.1": "1", "2.6.4": "1", "3.1.2": "1"},
         "Darstellungsvariante 1"),
        ("Beleuchtung Grossraumbuero", {"1.2.1": "1", "2.5.4": "1", "3.1.3": "1"}, ""),
    ]
    meta = {
        "gewerk": "GA",
        "anlage": "Basis BACnet-AE / Beispiel 1",
        "projekt": "Muster-Campus, Bauteil B",
        "planersteller": "Musterplanung GmbH",
        "planstand": "2024-06-01",
        "blatt_nr": "1",
        "blatt_von": "2",
        "protokoll": "BACnet/IP",
    }
    document = pymupdf.open()
    page = document.new_page(width=1190, height=595)
    _draw_list(page, profile, rows, meta, [], "GA-Funktionsliste VDI 3814 (neue Fassung)")
    document.save(path)
    document.close()


def main() -> None:
    build_alt(OUT_DIR / "beispiel_alte_fassung.pdf")
    build_neu(OUT_DIR / "beispiel_neue_fassung.pdf")
    print(f"Beispieldateien geschrieben nach {OUT_DIR}")


if __name__ == "__main__":
    main()
