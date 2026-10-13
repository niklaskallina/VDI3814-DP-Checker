"""Auswerteseite (Reiter 4): Kennzahlen, Qualitaetsbefunde und Diagramme."""

import copy

import pytest
from sqlalchemy.orm import Session

from vdi3814 import aggregate, db
from vdi3814.pipeline import process_file
from vdi3814.ui import auswertung


@pytest.fixture()
def projekt(tmp_path, samples):
    engine = db.make_engine(tmp_path / "t.sqlite3")
    with Session(engine) as session:
        for name in ("alt", "neu"):
            db.save_document(session, process_file(samples[name]))
    return engine


def test_kennzahlen_entsprechen_den_rohdaten(projekt):
    with Session(projekt) as session:
        roh = aggregate.raw_dataframe(session)
        gezaehlt = sum(len(db.load_document(session, d.id).counted_rows())
                       for d in db.list_documents(session))
    werte = auswertung.kennzahlen(roh, {"A_1_1": 10.0})
    assert werte.funktionen == pytest.approx(roh["wert"].sum())
    assert werte.datenpunkte == gezaehlt
    assert werte.dateien == 2
    menge_a11 = roh.loc[roh["spalte_key"] == "A_1_1", "wert"].sum()
    assert werte.kosten == pytest.approx(menge_a11 * 10.0)
    assert auswertung.kennzahlen(roh).kosten is None, "ohne Preise keine Kostenangabe"


def test_saubere_listen_ergeben_gruene_summenpruefung(projekt):
    with Session(projekt) as session:
        roh = aggregate.raw_dataframe(session)
        pruefung = auswertung.qualitaet(session, roh)
    assert pruefung.seiten_geprueft == 2 and not pruefung.abweichende_seiten
    assert pruefung.befunde()[0].stufe == "gut"
    assert pruefung.gesamtstufe != "kritisch"


def test_doppelte_liste_ist_ein_kritischer_befund(projekt, samples):
    kopie = process_file(samples["alt"])
    kopie = copy.deepcopy(kopie)
    kopie.file_hash = "kopie" + kopie.file_hash[5:]
    kopie.file_name = "zweimal.pdf"
    with Session(projekt) as session:
        db.save_document(session, kopie)
        roh = aggregate.raw_dataframe(session)
        pruefung = auswertung.qualitaet(session, roh)
    assert pruefung.dubletten
    assert pruefung.gesamtstufe == "kritisch"
    assert any("doppelte" in b.titel for b in pruefung.befunde())


@pytest.mark.parametrize("dunkel", [False, True])
def test_diagramme_entstehen_in_beiden_farbschemata(projekt, dunkel):
    with Session(projekt) as session:
        roh = aggregate.raw_dataframe(session)
    zusammenfassung = aggregate.column_summary(roh)
    diagramme = [
        auswertung.diagramm_gruppen(roh, dunkel),
        auswertung.diagramm_schwerpunkte(roh, dunkel),
        auswertung.diagramm_spalten(zusammenfassung, dunkel=dunkel),
        auswertung.diagramm_raster(roh, dunkel),
    ]
    for diagramm in diagramme:
        assert diagramm is not None
        spezifikation = diagramm.to_dict()
        assert spezifikation["config"]["background"] == "transparent"
    # Die Balken summieren sich zur Gesamtmenge
    werte = diagramme[0].to_dict()
    daten = next(iter(werte["datasets"].values()))
    assert sum(z["menge"] for z in daten) == pytest.approx(roh["wert"].sum())


def test_kacheln_und_befunde_sind_sauberes_html():
    werte = auswertung.Kennzahlen(funktionen=12480, datenpunkte=1520, dateien=3,
                                  schwerpunkte=4, kosten=40158.0, spalten_mit_preis=4)
    kacheln = auswertung.kacheln_html(werte, "EUR")
    assert "12.480" in kacheln and "40.158" in kacheln and "1.520" in kacheln
    befunde = auswertung.qualitaet_html(
        [auswertung.Befund("kritisch", "Datei <b>", "Text & mehr")])
    assert "&lt;b&gt;" in befunde and "&amp;" in befunde
    assert "✕" in befunde, "Status nie nur ueber die Farbe"


def test_leere_daten_liefern_keine_diagramme():
    import pandas as pd

    leer = pd.DataFrame(columns=aggregate.RAW_COLUMNS)
    assert auswertung.diagramm_gruppen(leer) is None
    assert auswertung.diagramm_raster(leer) is None
    assert auswertung.kennzahlen(leer).funktionen == 0
