"""Extraktion aus PDF-Textebene und Excel - gegen die erzeugten Beispieldateien."""

import openpyxl
import pytest

from vdi3814.extract.base import sections_by_position, truncate_after_sum
from vdi3814.extract.excel_extractor import extract_excel
from vdi3814.extract.pdftext_extractor import classify_page_text, extract_pdf_text
from vdi3814.models import DataPointRow
from vdi3814.pipeline import process_file


def test_alte_fassung_wird_vollstaendig_gelesen(samples):
    result = process_file(samples["alt"])
    assert result.fassung == "alt"
    assert result.profile_id == "vdi3814_alt"
    assert len(result.data_rows()) == 8
    assert all(column.column_key for column in result.columns)
    # Die eigene Summe muss mit der Zeile "Summe Funktionen" uebereinstimmen
    assert all(check.matches for check in result.sum_checks)
    assert result.grand_total() == 31


def test_neue_fassung_wird_vollstaendig_gelesen(samples):
    result = process_file(samples["neu"])
    assert result.fassung == "neu"
    assert len(result.data_rows()) == 8
    assert result.metadata.protokoll == "BACnet/IP"
    assert result.metadata.blatt_nr == "1"
    keys = {c.column_key for c in result.columns}
    assert "N_2_6_1" in keys and "N_3_1_2" in keys


def test_regelschema_wird_uebersprungen_ohne_modell(samples):
    result = process_file(samples["alt"])
    skipped = result.skipped_pages()
    assert [p.page_index for p in skipped] == [1]
    assert skipped[0].classification.kind.value == "schema"
    # uebersprungen heisst protokolliert, nicht verloren
    assert skipped[0].classification.reason


def test_klassifikation_ohne_modell(samples):
    kind, confidence, _ = classify_page_text(samples["alt"], 1)
    assert kind == "regelschema" and confidence >= 0.7


def test_spaltenadressen_der_alten_fassung(samples):
    table = extract_pdf_text(samples["alt"], 0)
    adressen = [c.address for c in table.columns if c.address]
    assert adressen[:5] == ["1.1", "1.2", "1.3", "1.4", "1.5"]
    assert "6.13" in adressen and "8.4" in adressen


def test_excel_vorlage_wird_gelesen(tmp_path):
    """Minimalvorlage im Aufbau der VDI-Excel-Datei.

    Spaltenaufteilung wie im Original: A = Zeile Nr., B = Datenpunkt,
    C = Beschriftungsspalte ("Abschnitt"/"Spalte"), ab D die Funktionsspalten.
    """
    path = tmp_path / "vorlage.xlsx"
    book = openpyxl.Workbook()
    sheet = book.active
    sheet["A1"] = " Gewerk:"
    sheet["B1"] = "Heizung"
    sheet["D2"] = "1. Ein-/Ausgabefunktionen"
    sheet["D3"] = "Physikalische"
    sheet["D4"] = "Analoge Eingabe (AI)"
    sheet["E4"] = "Binäre Eingabe (BI)"
    sheet["F4"] = "Analoge Ausgabe (AO)"
    sheet["G4"] = "Binäre Ausgabe (BO)"
    sheet["H4"] = "Grenzwertüberwachung"
    sheet["C5"], sheet["D5"], sheet["H5"] = "Abschnitt ", "1.1.", "2.2."
    sheet["C6"] = "Spalte "
    sheet["D6"], sheet["E6"], sheet["F6"], sheet["G6"], sheet["H6"] = "1", "2", "3", "4", "1"
    sheet["A7"], sheet["B7"], sheet["D7"] = "1", "Vorlauftemperatur", 2
    sheet["A8"], sheet["B8"], sheet["E8"], sheet["H8"] = "2", "Pumpe", 1, 1
    sheet["B9"] = "Summe Funktionen"
    sheet["D9"], sheet["E9"], sheet["H9"] = 2, 1, 1
    book.save(path)

    tables = extract_excel(path)
    assert len(tables) == 1
    table = tables[0]
    assert [c.address for c in table.columns] == ["1.1.1", "1.1.2", "1.1.3", "1.1.4", "2.2.1"]
    assert table.columns[0].label == "Analoge Eingabe (AI)"
    assert table.columns[0].subgroup == "Physikalische"
    data = [r for r in table.rows if not r.is_sum_row]
    assert [r.klartext for r in data] == ["Vorlauftemperatur", "Pumpe"]
    assert data[0].cells[0].count == 2.0
    assert any(r.is_sum_row for r in table.rows)
    assert table.metadata.gewerk == "Heizung"


def test_excel_pipeline_erkennt_neue_fassung(tmp_path):
    path = tmp_path / "neu.xlsx"
    book = openpyxl.Workbook()
    sheet = book.active
    sheet["D2"] = "2. Anwendungsfunktionen"
    sheet["D3"] = "Sonnenschutz"
    sheet["D4"] = "Sonnenschutz stellen"
    sheet["E4"] = "Sonnen-/Dämmerungsautomatik"
    sheet["F4"] = "Thermoautomatik"
    sheet["G4"] = "Witterungsschutz"
    sheet["C5"], sheet["D5"] = "Abschnitt ", "2.6."
    sheet["C6"] = "Spalte "
    sheet["D6"], sheet["E6"], sheet["F6"], sheet["G6"] = "1", "2", "3", "4"
    sheet["A7"], sheet["B7"], sheet["D7"] = "1", "Jalousie Sued", 4
    book.save(path)

    result = process_file(path)
    assert result.fassung == "neu"
    assert {c.column_key for c in result.columns} == {"N_2_6_1", "N_2_6_2", "N_2_6_3", "N_2_6_4"}
    assert result.grand_total() == 4
    assert result.data_rows()[0].klartext == "Jalousie Sued"


def test_abschnittszuordnung_ueber_position():
    # Spaltennummern 1..5 | 1..3, Abschnittsbeschriftung mittig darueber
    numbers = ["1", "2", "3", "4", "5", "1", "2", "3"]
    xs = [10, 20, 30, 40, 50, 70, 80, 90]
    sections = [(30.0, "1"), (80.0, "2"), (150.0, "9")]
    assert sections_by_position(numbers, xs, sections) == ["1"] * 5 + ["2"] * 3


def test_fussbereich_wird_abgeschnitten():
    rows = [
        DataPointRow(klartext="Pumpe"),
        DataPointRow(klartext="Summe Funktionen", is_sum_row=True),
        DataPointRow(klartext="Ausgabedatum"),
        DataPointRow(klartext="Rev. 1"),
    ]
    assert [r.klartext for r in truncate_after_sum(rows)] == ["Pumpe", "Summe Funktionen"]


def test_unbekannter_dateityp_stuerzt_nicht_ab(tmp_path):
    path = tmp_path / "notiz.txt"
    path.write_text("kein Plan", encoding="utf-8")
    result = process_file(path)
    assert result.columns == []
    assert any("Nicht unterstuetzter Dateityp" in w for w in result.warnings)


def test_bildseite_ohne_modell_meldet_verstaendlich(tmp_path):
    from PIL import Image

    path = tmp_path / "scan.png"
    Image.new("RGB", (600, 400), "white").save(path)
    result = process_file(path, backend=None)
    assert result.columns == []
    assert any("Vision-Modell" in w or "Funktionsliste" in w for w in result.warnings)
    assert result.pages[0].classification.kind.value == "fehler"


def test_fussbereich_mit_werten_unter_der_beschriftung(tmp_path):
    """Im Fussbereich stehen Projekt/Planersteller unter, nicht neben der Beschriftung."""
    import pymupdf

    path = tmp_path / "fuss.pdf"
    document = pymupdf.open()
    page = document.new_page(width=842, height=595)
    page.insert_text((40, 400), "Ausgabedatum", fontsize=7)
    page.insert_text((160, 400), "Planersteller", fontsize=7)
    page.insert_text((320, 400), "Projekt", fontsize=7)
    page.insert_text((40, 412), "2024-11-02", fontsize=7)
    page.insert_text((160, 412), "IPG Ingenieur-Planung", fontsize=7)
    page.insert_text((320, 412), "Mehrzweckhalle Tralau", fontsize=7)
    document.save(path)
    document.close()

    from vdi3814.extract.pdftext_extractor import _cluster_rows, _extract_metadata, _HeaderGrid

    with pymupdf.open(path) as opened:
        words = opened[0].get_text("words")
    grid = _HeaderGrid(columns=[], body_top=0.0, label_left=800.0, note_left=None)
    metadata = _extract_metadata(words, grid, skip=set())

    assert metadata.planersteller == "IPG Ingenieur-Planung"
    assert metadata.projekt == "Mehrzweckhalle Tralau"
    assert metadata.datum == "2024-11-02"
