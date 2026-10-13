"""Genauigkeit der Datenpunktliste: was zaehlt, was nicht - und nichts doppelt.

Jeder Test hier haelt einen Fehler fest, der in echten Plansaetzen Mengen
verfaelscht hat: Datenpunkte, die wie Summen- oder Fusszeilen klangen und
verschluckt wurden, eine Liste, die nach einer wiederholten Kopfzeile oder
einem "Freigabe"-Datenpunkt abgeschnitten wurde, Werte, die doppelt in zwei
Spalten fielen, und Summenpruefungen, die den Uebertrag vergassen.
"""

import openpyxl
import pytest
from openpyxl.styles import Font
from sqlalchemy.orm import Session

from vdi3814 import aggregate, db, evidence
from vdi3814.extract.base import drop_after_footer
from vdi3814.extract.rowfilter import classify_row
from vdi3814.models import DataPointRow, RowKind
from vdi3814.pipeline import process_file
from vdi3814.textutil import parse_count

# Typische Bezeichnungen aus GA-Funktionslisten, die frueher als Summen-,
# Uebertrags- oder Fusszeile galten und deshalb nicht gezaehlt wurden.
ECHTE_DATENPUNKTE = [
    "Freigabe Kessel",
    "Freigabe Anlage",
    "Revisionsschalter",
    "Revisionsschalter Ventilator",
    "Summenstörung",
    "Summenstoermeldung Lüftung",
    "Sammel-/Summenstörung",
    "Gesamtstörung",
    "Gesamt Störung",
    "Totalausfall Pumpe",
    "Temp. Wärmeübertrager",
    "Übertrager Vorlauf",
    "Störung Übertragung",
    "Fehler Datenübertragung",
    "Index Sollwert",
    "Erstellt Datum",
    "Bearbeitung Zeitprogramm",
    "Anmerkung zur Pumpe",
]


@pytest.mark.parametrize("beschriftung", ECHTE_DATENPUNKTE)
@pytest.mark.parametrize("mit_werten", [True, False])
def test_echte_datenpunkte_bleiben_datenpunkte(beschriftung, mit_werten):
    art, _ = classify_row(beschriftung, mit_werten)
    assert art is RowKind.DATEN, beschriftung


@pytest.mark.parametrize("beschriftung, erwartet", [
    ("Summe", RowKind.SUMME),
    ("Summe:", RowKind.SUMME),
    ("Summe Funktionen", RowKind.SUMME),
    ("Summe ASP01", RowKind.SUMME),
    ("Zwischensumme Heizung", RowKind.SUMME),
    ("Gesamtsumme", RowKind.SUMME),
    ("Gesamt", RowKind.SUMME),
    ("Gesamt Funktionen", RowKind.SUMME),
    ("Total ASP02", RowKind.SUMME),
    ("Anzahl Funktionen", RowKind.SUMME),
    ("Übertrag", RowKind.UEBERTRAG),
    ("Übertrag von Blatt 2", RowKind.UEBERTRAG),
    ("Summe Übertrag", RowKind.UEBERTRAG),
    ("Planersteller", RowKind.FUSSBEREICH),
    ("Blatt-Nr.: 3", RowKind.FUSSBEREICH),
    ("Ausgabedatum", RowKind.FUSSBEREICH),
])
def test_summen_uebertraege_und_fussbereich_werden_weiter_erkannt(beschriftung, erwartet):
    assert classify_row(beschriftung, True)[0] is erwartet


@pytest.mark.parametrize("beschriftung", ["Freigabe:", "Freigabe", "Rev.", "Index",
                                          "Bearbeiter: Müller", "Anmerkung:"])
def test_felder_des_schriftkopfs_ohne_werte_sind_fussbereich(beschriftung):
    assert classify_row(beschriftung, False)[0] is RowKind.FUSSBEREICH
    # Mit Werten in den Funktionsspalten ist es dagegen ein Datenpunkt
    assert classify_row(beschriftung, True)[0] is RowKind.DATEN


def test_wiederholte_kopfzeile_schneidet_die_liste_nicht_ab():
    zeilen = [
        DataPointRow(klartext="Pumpe 1"),
        DataPointRow(klartext="Spalte", kind=RowKind.KOPF),
        DataPointRow(klartext="Pumpe 2"),
        DataPointRow(klartext="Planersteller", kind=RowKind.FUSSBEREICH),
        DataPointRow(klartext="Muster GmbH"),
    ]
    assert [z.klartext for z in drop_after_footer(zeilen)] == ["Pumpe 1", "Pumpe 2"]


@pytest.mark.parametrize("text, erwartet", [
    ("1", 1.0), ("12", 12.0), ("x", 1.0), ("X", 1.0),
    ("1.205", 1205.0), ("12.480", 12480.0), ("1.000.000", 1_000_000.0),
    ("1,5", 1.5), ("3 ", 3.0), ("", None), ("entfällt", None),
])
def test_zahlen_werden_richtig_gelesen(text, erwartet):
    assert parse_count(text) == erwartet


# --------------------------------------------------------------------------
# Excel: der genaueste Weg - hier darf gar nichts verloren gehen
# --------------------------------------------------------------------------

def _kopf(sheet, startzeile: int = 2) -> int:
    """Schreibt den VDI-Kopf (Abschnitt 1.1, vier Spalten) und liefert die erste Datenzeile."""
    z = startzeile
    sheet.cell(z, 4, "1. Ein-/Ausgabefunktionen")
    sheet.cell(z + 1, 4, "Physikalische")
    for spalte, text in zip(range(4, 8), ("Analoge Eingabe (AI)", "Binäre Eingabe (BI)",
                                          "Analoge Ausgabe (AO)", "Binäre Ausgabe (BO)")):
        sheet.cell(z + 2, spalte, text)
    sheet.cell(z + 3, 3, "Abschnitt ")
    sheet.cell(z + 3, 4, "1.1.")
    sheet.cell(z + 4, 3, "Spalte ")
    for spalte, nummer in zip(range(4, 8), ("1", "2", "3", "4")):
        sheet.cell(z + 4, spalte, nummer)
    return z + 5


def test_excel_liste_zaehlt_jeden_datenpunkt_genau_einmal(tmp_path):
    pfad = tmp_path / "liste.xlsx"
    buch = openpyxl.Workbook()
    blatt = buch.active
    z = _kopf(blatt)
    daten = [
        ("Vorlauftemperatur", {4: 1}),
        ("Freigabe Kessel", {7: 1}),
        ("Revisionsschalter", {5: 1}),
        ("Summenstörung", {5: 1}),
        ("Temp. Wärmeübertrager", {4: 1}),
    ]
    for nummer, (text, werte) in enumerate(daten, start=1):
        blatt.cell(z, 1, str(nummer))
        blatt.cell(z, 2, text)
        for spalte, wert in werte.items():
            blatt.cell(z, spalte, wert)
        z += 1
    # Wiederholte Nummernzeile mitten in der Liste (Druck ueber mehrere Seiten)
    blatt.cell(z, 3, "Spalte ")
    for spalte, nummer in zip(range(4, 8), ("1", "2", "3", "4")):
        blatt.cell(z, spalte, nummer)
    z += 1
    blatt.cell(z, 1, "6"), blatt.cell(z, 2, "Gesamtstörung"), blatt.cell(z, 5, 1)
    z += 1
    # Ausgeblendete und durchgestrichene Zeile: entfallene Datenpunkte
    blatt.cell(z, 1, "7"), blatt.cell(z, 2, "Pumpe alt"), blatt.cell(z, 5, 2)
    blatt.row_dimensions[z].hidden = True
    z += 1
    blatt.cell(z, 1, "8"), blatt.cell(z, 2, "Ventil entfällt"), blatt.cell(z, 6, 1)
    for spalte in (1, 2, 6):
        blatt.cell(z, spalte).font = Font(strike=True)
    z += 1
    # Ein ueber zwei Zeilen verbundener Wert ist EIN Wert
    blatt.cell(z, 1, "9"), blatt.cell(z, 2, "Klappe Zuluft"), blatt.cell(z, 7, 2)
    blatt.merge_cells(start_row=z, start_column=7, end_row=z + 1, end_column=7)
    blatt.cell(z + 1, 1, "10"), blatt.cell(z + 1, 2, "Klappe Abluft"), blatt.cell(z + 1, 6, 1)
    z += 2
    blatt.cell(z, 2, "Summe Funktionen")
    for spalte, wert in {4: 2, 5: 3, 6: 1, 7: 3}.items():
        blatt.cell(z, spalte, wert)
    blatt.cell(z + 2, 2, "Planersteller"), blatt.cell(z + 2, 4, "Muster GmbH")
    buch.save(pfad)

    ergebnis = process_file(pfad)

    gezaehlt = [r.klartext for r in ergebnis.counted_rows()]
    assert gezaehlt == ["Vorlauftemperatur", "Freigabe Kessel", "Revisionsschalter",
                        "Summenstörung", "Temp. Wärmeübertrager", "Gesamtstörung",
                        "Klappe Zuluft", "Klappe Abluft"]
    assert ergebnis.grand_total() == 9      # 1+1+1+1+1+1 + 2 + 1
    entfallen = {r.klartext: r.exclusion_reason for r in ergebnis.rows
                 if r.kind is RowKind.ENTFALLEN}
    assert set(entfallen) == {"Pumpe alt", "Ventil entfällt"}
    assert all(entfallen.values()), "entfallene Zeilen brauchen eine Begruendung"
    # Die Summenzeile der Liste geht damit exakt auf
    assert ergebnis.sum_checks and all(c.matches for c in ergebnis.sum_checks)


def test_uebertrag_gehoert_zur_erwarteten_seitensumme(tmp_path):
    """Zwei Blaetter: das zweite beginnt mit dem Uebertrag des ersten.

    Die Summenzeile des zweiten Blatts enthaelt Uebertrag + eigene Zeilen.
    Weder der Nachweis noch der Pruefbericht duerfen das als Abweichung melden.
    """
    pfad = tmp_path / "zwei_blaetter.xlsx"
    buch = openpyxl.Workbook()
    erstes = buch.active
    z = _kopf(erstes)
    erstes.cell(z, 1, "1"), erstes.cell(z, 2, "Pumpe 1"), erstes.cell(z, 5, 2)
    erstes.cell(z + 1, 1, "2"), erstes.cell(z + 1, 2, "Pumpe 2"), erstes.cell(z + 1, 5, 1)
    erstes.cell(z + 2, 2, "Summe Funktionen"), erstes.cell(z + 2, 5, 3)
    zweites = buch.create_sheet("Blatt 2")
    z = _kopf(zweites)
    zweites.cell(z, 2, "Übertrag von Blatt 1"), zweites.cell(z, 5, 3)
    zweites.cell(z + 1, 1, "3"), zweites.cell(z + 1, 2, "Pumpe 3"), zweites.cell(z + 1, 5, 4)
    zweites.cell(z + 2, 2, "Summe Funktionen"), zweites.cell(z + 2, 5, 7)
    buch.save(pfad)

    ergebnis = process_file(pfad)
    assert ergebnis.grand_total() == 7
    assert all(c.matches for c in ergebnis.sum_checks)

    engine = db.make_engine(tmp_path / "t.sqlite3")
    with Session(engine) as session:
        db.save_document(session, ergebnis)
        befunde = evidence.all_findings(session)
        assert not [b for b in befunde if b.art == "abweichung"], \
            [b.titel for b in befunde if b.art == "abweichung"]
        bericht = aggregate.pruefbericht(session)
    abgleich = bericht[bericht["art"] == "Summenabgleich"]
    assert set(abgleich["seite"]) == {1, 2}, "verglichen wird je Seite"
    assert (abgleich["bewertung"] == "stimmt überein").all(), abgleich.to_dict("records")


# --------------------------------------------------------------------------
# Wortraster (PDF-Textebene und OCR): ein Wert faellt in genau eine Spalte
# --------------------------------------------------------------------------

def test_wert_faellt_bei_ungleichen_spalten_nur_in_eine_spalte():
    from vdi3814.extract.base import ExtractionMode
    from vdi3814.extract.wordgrid import _HeaderGrid, extract_table_from_words
    from vdi3814.models import ColumnHeader

    # Spalten 100 | 110 | 160 | 170: mittlerer Abstand 23,3 - die Toleranz
    # (knapp 10,5) reicht von 106 aus an beide linken Spalten heran.
    mitten = [100.0, 110.0, 160.0, 170.0]
    spalten = [ColumnHeader(index=i, label=f"S{i}", section="1.1", column_no=str(i + 1),
                            x_center=x) for i, x in enumerate(mitten)]
    raster = _HeaderGrid(columns=spalten, body_top=50.0, label_left=80.0, note_left=200.0)

    def wort(x, y, text, nr):
        return (x - 2, y - 4, x + 2, y + 4, text, 0, 0, nr)

    woerter = [
        wort(20, 60, "1", 1), wort(40, 60, "Pumpe", 2), wort(106, 60, "1", 3),
        wort(20, 75, "2", 4), wort(40, 75, "Ventil", 5), wort(160, 75, "2", 6),
    ]
    tabelle = extract_table_from_words(woerter, 300.0, 0, ExtractionMode.PDF_TEXT,
                                       grid=raster, table_bottom=90.0)
    pumpe = next(r for r in tabelle.rows if r.klartext == "Pumpe")
    assert [(c.column_index, c.count) for c in pumpe.cells] == [(1, 1.0)], \
        "der Wert darf nur in der naechstgelegenen Spalte stehen"
    assert sum(c.count or 0 for r in tabelle.rows for c in r.cells) == 3.0


# --------------------------------------------------------------------------
# Scans: ein Kreuz zaehlt wie eine 1
# --------------------------------------------------------------------------

def test_kreuz_im_scan_wird_als_eins_gelesen():
    from vdi3814.extract.ocr_extractor import _lies_ziffer, available

    if not available():
        pytest.skip("Tesseract ist nicht installiert")
    import numpy as np
    from PIL import Image, ImageDraw, ImageFont

    try:
        schrift = ImageFont.truetype("DejaVuSans.ttf", 22)
    except OSError:
        pytest.skip("Schrift DejaVuSans nicht vorhanden")
    ergebnisse = {}
    for zeichen in ("x", "X", "2", "4"):
        bild = Image.new("L", (34, 34), 255)
        zeichnung = ImageDraw.Draw(bild)
        zeichnung.rectangle([0, 0, 33, 33], outline=0)
        zeichnung.text((9, 4), zeichen, fill=0, font=schrift)
        ergebnisse[zeichen] = _lies_ziffer(np.asarray(bild)[3:-2, 3:-2])
    assert ergebnisse == {"x": "x", "X": "x", "2": "2", "4": "4"}
    assert parse_count(ergebnisse["x"]) == 1.0


def test_excel_mittig_gesetzte_abschnittsnummern_treffen_ihren_block(tmp_path):
    """Abschnittsnummer mittig ueber dem Block, nicht verbunden, dazu ein
    Bemerkungsabschnitt ohne Spaltennummer - die Anzahl passt dann nicht zu
    den Bloecken. Frueher bekam die erste Spalte keinen Abschnitt und die
    Werte landeten in einer falschen bzw. unbekannten Spalte."""
    pfad = tmp_path / "mittig.xlsx"
    buch = openpyxl.Workbook()
    blatt = buch.active
    blatt["D2"] = "1. Ein-/Ausgabefunktionen"
    for spalte, text in zip("DEFGHI", ("Analoge Eingabe (AI)", "Binäre Eingabe (BI)",
                                       "Analoge Ausgabe (AO)", "Binäre Ausgabe (BO)",
                                       "Analogwert (AV)", "Binärwert(e) (BV/MV)")):
        blatt[f"{spalte}4"] = text
    blatt["C5"], blatt["E5"], blatt["H5"], blatt["K5"] = "Abschnitt", "1.1.", "1.2.", "9."
    blatt["C6"] = "Spalte"
    for spalte, nummer in zip("DEFGHI", ("1", "2", "3", "4", "1", "2")):
        blatt[f"{spalte}6"] = nummer
    blatt["K4"] = "Bemerkungen"
    blatt["A7"], blatt["B7"], blatt["D7"], blatt["H7"] = "1", "Raumtemperatur", 1, 2
    buch.save(pfad)

    ergebnis = process_file(pfad)
    # Der Kopf traegt kaum Text - die Fassung folgt dann aus den Adressen
    assert ergebnis.fassung == "neu"
    werte = {c.column_key: cell.count for c in ergebnis.columns
             for cell in ergebnis.counted_rows()[0].cells if cell.column_index == c.index}
    assert werte == {"N_1_1_1": 1.0, "N_1_2_1": 2.0}


@pytest.mark.parametrize("name", ["alt", "neu"])
def test_spaltenueberschriften_stimmen_mit_der_vorlage_ueberein(samples, name):
    """Die senkrechten Ueberschriften muessen bei ihrer Spalte bleiben.

    Bei zweistelligen Spaltennummern (6.10 ... 6.13, 2.2.10 ... 2.2.12)
    rutschte die Ueberschrift frueher in die Nachbarspalte - in Kosten-
    schaetzung und Export stand dann unter der richtigen Spalte der falsche
    Name.
    """
    from vdi3814.profiles_loader import get_profile
    from vdi3814.textutil import similarity

    ergebnis = process_file(samples[name])
    vorlage = get_profile(ergebnis.profile_id)
    verrutscht = []
    for spalte in ergebnis.columns:
        eintrag = vorlage.column(spalte.column_key) if spalte.column_key else None
        if eintrag is None:
            continue
        passung = max(similarity(spalte.label, k) for k in eintrag.candidates() if k)
        if passung < 0.7:
            verrutscht.append((spalte.address, spalte.label, eintrag.label))
    assert not verrutscht, verrutscht
    assert not any("Spaltenueberschrift" in w for w in ergebnis.warnings)


def test_korrekturtabelle_verliert_keine_anmerkungen(samples):
    """Schon das Anzeigen der Korrekturtabelle hat Freitext und Originalwerte
    ueberschrieben: aus "x" wurde "1.0", Anmerkungen in Funktionsspalten und
    Zellen der Bemerkungsspalten verschwanden."""
    from vdi3814.models import Cell
    from vdi3814.ui.table import frame_to_result, result_to_frame

    ergebnis = process_file(samples["alt"])
    zeile = ergebnis.data_rows()[0]
    spalte = next(c for c in ergebnis.columns if not c.is_note_column
                  and zeile.value(c.index) is None)
    zeile.cells.append(Cell(column_index=spalte.index, raw_value="s. Bem.", count=None,
                            note="s. Bem."))
    gezaehlt = next(c for c in zeile.cells if c.count is not None)
    gezaehlt.raw_value = "x" if gezaehlt.count == 1 else gezaehlt.raw_value
    vorher = ergebnis.grand_total()

    frame_to_result(result_to_frame(ergebnis), ergebnis)      # unveraendert zurueck

    assert ergebnis.grand_total() == vorher
    notizen = [c.note for c in zeile.cells if c.count is None]
    assert "s. Bem." in notizen
    assert any(c.raw_value == gezaehlt.raw_value for c in zeile.cells if c.count is not None)


def test_korrekturtabelle_uebernimmt_geaenderte_werte(samples):
    from vdi3814.ui.table import column_title, frame_to_result, result_to_frame

    ergebnis = process_file(samples["alt"])
    tabelle = result_to_frame(ergebnis)
    spalte = next(c for c in ergebnis.columns if not c.is_note_column)
    vorher = ergebnis.grand_total()
    alter_wert = tabelle.loc[0, column_title(spalte)]
    tabelle.loc[0, column_title(spalte)] = 5.0
    frame_to_result(tabelle, ergebnis)
    alt = 0.0 if alter_wert is None or alter_wert != alter_wert else float(alter_wert)
    assert ergebnis.grand_total() == pytest.approx(vorher - alt + 5.0)
    assert ergebnis.data_rows()[0].value(spalte.index) == 5.0


def test_spalte_erscheint_nur_einmal_auch_bei_abweichender_beschriftung(tmp_path, samples):
    """Zwei Dateien, dieselbe Spalte, unterschiedlich gelesene Ueberschrift.

    Frueher entstanden daraus zwei Zeilen in Spaltensummen und Kosten-
    schaetzung - im Excel-Export verwiesen beide auf dieselbe Summe, Menge
    und Kosten zaehlten doppelt.
    """
    import copy

    erste = process_file(samples["alt"])
    zweite = copy.deepcopy(erste)
    zweite.file_hash = "anders" + erste.file_hash[6:]
    for spalte in zweite.columns:
        spalte.label = spalte.label.upper() + " (gelesen)"
        spalte.group = "Gruppe falsch gelesen"

    engine = db.make_engine(tmp_path / "t.sqlite3")
    with Session(engine) as session:
        db.save_document(session, erste)
        db.save_document(session, zweite)
        roh = aggregate.raw_dataframe(session)
    uebersicht = aggregate.column_summary(roh)

    assert not uebersicht["spalte_key"].duplicated().any(), \
        uebersicht[uebersicht["spalte_key"].duplicated(keep=False)]
    assert uebersicht["menge"].sum() == pytest.approx(2 * erste.grand_total())
    from vdi3814.profiles_loader import get_profile

    vorlage = {c.group for c in get_profile("vdi3814_alt").columns}
    assert set(aggregate.group_summary(roh)["gruppe"]) <= vorlage


def _zeile(text, werte):
    from vdi3814.models import Cell

    return DataPointRow(klartext=text, cells=[Cell(column_index=i, raw_value=str(w), count=w)
                                              for i, w in werte.items()])


def test_letzter_datenpunkt_ohne_lesbaren_namen_bleibt_gezaehlt():
    """Im Scan ist der Name der letzten Zeile unlesbar - sie ist trotzdem ein
    Datenpunkt und keine Summenzeile."""
    from vdi3814.extract.rowfilter import markiere_unbeschriftete_summenzeilen

    zeilen = [_zeile("Pumpe", {0: 1, 1: 1}), _zeile("Ventil", {2: 1, 3: 1}),
              _zeile("Klappe", {4: 1, 5: 1}), _zeile("", {1: 1})]
    markiere_unbeschriftete_summenzeilen(zeilen)
    assert zeilen[-1].kind is RowKind.DATEN


def test_unbeschriftete_summenzeile_wird_weiter_erkannt():
    from vdi3814.extract.rowfilter import markiere_unbeschriftete_summenzeilen

    zeilen = [_zeile("Pumpe", {0: 1, 1: 1}), _zeile("Ventil", {0: 1, 2: 1}),
              _zeile("Klappe", {1: 2, 2: 1}), _zeile("", {0: 2, 1: 3, 2: 3})]
    markiere_unbeschriftete_summenzeilen(zeilen)
    assert zeilen[-1].kind is RowKind.SUMME
    # auch, wenn ein Wert falsch gelesen wurde (3 statt 2)
    zeilen = [_zeile("Pumpe", {0: 1, 1: 1}), _zeile("Ventil", {0: 1, 2: 1}),
              _zeile("Klappe", {1: 2, 2: 1}), _zeile("", {0: 3, 1: 3, 2: 2})]
    markiere_unbeschriftete_summenzeilen(zeilen)
    assert zeilen[-1].kind is RowKind.SUMME


def test_doppelt_importierte_liste_wird_gemeldet(tmp_path, samples):
    """Derselbe Plansatz als zweite Datei (anderer Hash) - die Mengen zaehlten
    sonst unbemerkt doppelt."""
    import copy

    erste = process_file(samples["alt"])
    kopie = copy.deepcopy(erste)
    kopie.file_hash = "kopie" + erste.file_hash[5:]
    kopie.file_name = "Plansatz_als_Excel.xlsx"
    engine = db.make_engine(tmp_path / "t.sqlite3")
    with Session(engine) as session:
        db.save_document(session, erste)
        db.save_document(session, kopie)
        db.save_document(session, process_file(samples["neu"]))
        roh = aggregate.raw_dataframe(session)
    dubletten = aggregate.moegliche_dubletten(roh)
    assert len(dubletten) == 1
    assert set(dubletten.iloc[0][["datei_a", "datei_b"]]) == {
        erste.file_name, "Plansatz_als_Excel.xlsx"}
    assert dubletten.iloc[0]["anteil"] == 1.0
