"""Abnahme des Excel-Exports.

Der Export ist das Arbeitsergebnis: dort werden Mengen abgelesen und
Einheitspreise eingetragen. Diese Tests pruefen deshalb nicht nur, dass
Formeln dastehen, sondern rechnen sie nach und vergleichen mit der Datenbank.
"""

import openpyxl
import pytest
from openpyxl.utils import get_column_letter
from sqlalchemy.orm import Session

from vdi3814 import aggregate, db, export_excel
from vdi3814.pipeline import process_file
from vdi3814.profiles_loader import get_profile

PREISE = {"A_1_1": 45.0, "A_1_5": 80.0, "A_8_1": 12.5}


@pytest.fixture(scope="module")
def export(tmp_path_factory, samples):
    """Importiert beide Beispiellisten und schreibt einen echten Export."""
    verzeichnis = tmp_path_factory.mktemp("export")
    engine = db.make_engine(verzeichnis / "test.sqlite3")
    with Session(engine) as session:
        for name in ("alt", "neu"):
            db.save_document(session, process_file(samples[name]))
        db.set_unit_prices(session, PREISE)
        roh = aggregate.raw_dataframe(session)
        ziel = export_excel.export_workbook(
            verzeichnis / "Auswertung.xlsx",
            summary=aggregate.column_summary(roh),
            raw=roh,
            documents=aggregate.documents_frame(session),
            footnotes=aggregate.footnotes_frame(session),
            prices=db.get_unit_prices(session),
            projects=aggregate.pivot_projects(roh),
            layouts={p: aggregate.vdi_layout_frame(session, p)
                     for p in aggregate.profiles_in_use(session)},
            matrizen={p: aggregate.datei_spalten_matrix(session, p)
                      for p in aggregate.profiles_in_use(session)},
            pruefung=aggregate.pruefbericht(session),
        )
        mengen = roh.groupby("spalte_key")["wert"].sum().to_dict()
    return ziel, mengen


def test_uebersicht_ist_ein_vdi_blatt_mit_einer_zeile_je_datei(export):
    ziel, _ = export
    buch = openpyxl.load_workbook(ziel)
    blaetter = [name for name in buch.sheetnames if name.startswith("Übersicht")]
    assert blaetter, "es muss eine Uebersicht im VDI-Aufbau geben"

    blatt = buch[blaetter[0]]
    # Kopfaufbau wie in der Vorlage
    assert blatt.cell(1, 1).value == "Datei"
    assert blatt.cell(4, 6).value == "Abschnitt"
    assert blatt.cell(5, 6).value == "Spalte"
    # Nummernzeilen ergeben zusammen die Spaltenadresse
    assert f"{blatt.cell(4, 7).value}.{blatt.cell(5, 7).value}" == "1.1"

    # Je Datei genau eine Zeile, darunter Summe / Einheitspreis / Kosten
    beschriftungen = [blatt.cell(z, 1).value for z in range(6, blatt.max_row + 1)]
    assert "Summe Funktionen" in beschriftungen
    assert any(str(b).startswith("Einheitspreis") for b in beschriftungen)
    assert any(str(b).startswith("Kosten") for b in beschriftungen)
    dateien = [b for b in beschriftungen if str(b).endswith(".pdf")]
    assert len(dateien) == 1, "je Fassung eine Beispieldatei"


def test_mengen_und_kosten_sind_formeln(export):
    """Ohne Formeln kann in Excel nicht weitergerechnet werden."""
    ziel, _ = export
    buch = openpyxl.load_workbook(ziel)
    blatt = buch[[n for n in buch.sheetnames if n.startswith("Übersicht")][0]]
    summenzeile = blatt.max_row - 2
    kostenzeile = blatt.max_row

    assert str(blatt.cell(summenzeile, 7).value).startswith("=SUM(")
    assert str(blatt.cell(kostenzeile, 7).value).startswith("=G")
    # Der Einheitspreis muss eine schlichte Zahl bleiben - er wird eingetippt
    assert isinstance(blatt.cell(summenzeile + 1, 7).value, (int, float))

    kosten = buch["Kostenschätzung"]
    assert str(kosten.cell(2, 6).value).startswith("='Übersicht"), \
        "die Menge muss auf die Uebersicht verweisen, nicht kopiert sein"
    assert kosten.cell(2, 8).value == "=F2*G2"


def test_formeln_rechnen_richtig(export):
    """Formeln werden ausgewertet und gegen die Datenbank geprueft."""
    formulas = pytest.importorskip("formulas", reason="Formelrechner nicht installiert")
    import logging
    import warnings

    warnings.filterwarnings("ignore")
    logging.disable(logging.WARNING)

    ziel, mengen = export
    modell = formulas.ExcelModel().loads(str(ziel)).finish()
    loesung = modell.calculate()
    werte = {}
    for schluessel, wert in loesung.items():
        teile = schluessel.split("!")
        if len(teile) < 2:
            continue
        blattname = teile[-2].split("]")[-1].strip("'").upper()
        try:
            werte[(blattname, teile[-1].strip("'"))] = wert.value[0, 0]
        except Exception:
            continue

    buch = openpyxl.load_workbook(ziel)
    name = [n for n in buch.sheetnames if n.startswith("Übersicht")][0]
    blatt = buch[name]
    summenzeile = blatt.max_row - 2
    kostenzeile = blatt.max_row
    profil = get_profile("vdi3814_alt" if "alt" in name or len(buch.sheetnames) else "vdi3814_alt")

    geprueft = 0
    for spalte in range(7, blatt.max_column):
        adresse = f"{blatt.cell(4, spalte).value}.{blatt.cell(5, spalte).value}"
        eintrag = profil.column_by_address(adresse)
        if eintrag is None:
            continue
        buchstabe = get_column_letter(spalte)
        summe = werte.get((name.upper(), f"{buchstabe}{summenzeile}"))
        if summe is None:
            continue
        erwartet = mengen.get(eintrag.key, 0.0)
        assert float(summe) == pytest.approx(erwartet), f"Spalte {adresse}"

        kosten = werte.get((name.upper(), f"{buchstabe}{kostenzeile}"))
        assert float(kosten) == pytest.approx(erwartet * PREISE.get(eintrag.key, 0.0)), \
            f"Kosten in Spalte {adresse}"
        geprueft += 1
    assert geprueft >= 10, "es muessen mehrere Spalten geprueft worden sein"


def test_zeilensumme_je_datei_stimmt(export):
    """Die Zeilensumme einer Datei muss deren Gesamtmenge sein."""
    ziel, _ = export
    buch = openpyxl.load_workbook(ziel)
    name = [n for n in buch.sheetnames if n.startswith("Übersicht")][0]
    blatt = buch[name]
    # Ganz rechts liegt die ausgeblendete Schluesselspalte, davor "Funktionen gesamt".
    gesamt_spalte = blatt.max_column - 1
    letzte = get_column_letter(gesamt_spalte)

    werte = _rechne(ziel)

    zeile = 6
    while str(blatt.cell(zeile, 1).value or "").endswith(".pdf"):
        einzeln = sum(_zahl(werte, name, spalte, zeile)
                      for spalte in range(7, gesamt_spalte))
        gesamt = werte.get((name.upper(), f"{letzte}{zeile}"))
        assert float(gesamt) == pytest.approx(einzeln)
        zeile += 1
    assert zeile > 6, "es muss mindestens eine Dateizeile geben"


def test_uebersicht_haengt_an_der_ga_funktionsliste(export):
    """Ohne Verknuepfung wuerde ein Loeschen von Zeilen unbemerkt bleiben."""
    ziel, _ = export
    buch = openpyxl.load_workbook(ziel)
    name = [n for n in buch.sheetnames if n.startswith("Übersicht")][0]
    blatt = buch[name]
    liste = [n for n in buch.sheetnames if n.startswith("GA-Funktionsliste")][0]

    # Verknuepft wird ueber einen Schluessel je Liste - Dateinamen sind nicht
    # eindeutig. Die Spalte steht ganz rechts und ist ausgeblendet.
    schluessel = blatt.max_column
    assert blatt.cell(1, schluessel).value == "Schlüssel"
    assert blatt.column_dimensions[get_column_letter(schluessel)].hidden
    assert blatt.cell(6, schluessel).value == buch[liste].cell(6, buch[liste].max_column).value

    # Mengen und Datenpunkte werden gezaehlt, nicht kopiert.
    menge = str(blatt.cell(6, 7).value)
    assert menge.startswith("=SUMIFS("), menge
    assert f"'{liste}'!" in menge
    assert str(blatt.cell(6, 6).value).startswith("=COUNTIFS(")

    # Auch die flache Spaltenliste zieht die Menge aus der Uebersicht.
    assert str(buch["Spaltensummen"].cell(2, 7).value).startswith("='Übersicht")


def test_geloeschte_zeile_wirkt_bis_in_die_kosten(export, tmp_path):
    """Der eigentliche Zweck: in der GA-Funktionsliste darf geloescht werden."""
    ziel, _ = export
    buch = openpyxl.load_workbook(ziel)
    name = [n for n in buch.sheetnames if n.startswith("Übersicht")][0]
    liste = buch[[n for n in buch.sheetnames if n.startswith("GA-Funktionsliste")][0]]

    gesamt_zelle = (f"{get_column_letter(buch[name].max_column - 1)}"
                    f"{buch[name].max_row - 2}")          # Summenzeile, Spalte "Funktionen gesamt"
    vorher = float(_rechne(ziel)[(name.upper(), gesamt_zelle)])

    # Erste Datenzeile entfernen - so wie es in Excel jemand tun wuerde.
    entfallen = sum(wert for spalte in range(7, liste.max_column - 1)
                    if isinstance(wert := liste.cell(6, spalte).value, (int, float)))
    assert entfallen > 0, "die geloeschte Zeile muss Mengen enthalten"
    letzte_datenzeile = liste.max_row - 3
    liste.delete_rows(letzte_datenzeile + 1, 3)   # Summe/Einheitspreis/Kosten der Liste
    liste.delete_rows(6)
    gekuerzt = tmp_path / "gekuerzt.xlsx"
    buch.save(gekuerzt)

    nachher = float(_rechne(gekuerzt)[(name.upper(), gesamt_zelle)])
    assert nachher == pytest.approx(vorher - entfallen)


def _rechne(pfad) -> dict:
    """Wertet alle Formeln der Mappe aus - Schluessel: (BLATT, "A1")."""
    formulas = pytest.importorskip("formulas", reason="Formelrechner nicht installiert")
    import logging
    import warnings

    warnings.filterwarnings("ignore")
    logging.disable(logging.WARNING)

    loesung = formulas.ExcelModel().loads(str(pfad)).finish().calculate()
    werte = {}
    for schluessel, wert in loesung.items():
        teile = schluessel.split("!")
        if len(teile) < 2:
            continue
        try:
            werte[(teile[-2].split("]")[-1].strip("'").upper(), teile[-1].strip("'"))] = wert.value[0, 0]
        except Exception:
            continue
    return werte


def _zahl(werte: dict, blatt: str, spalte: int, zeile: int) -> float:
    try:
        return float(werte[(blatt.upper(), f"{get_column_letter(spalte)}{zeile}")])
    except (KeyError, TypeError, ValueError):
        return 0.0
