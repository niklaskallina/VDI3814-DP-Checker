"""Die Oberflaeche muss in allen Zustaenden fehlerfrei zeichnen.

Ein Streamlit-Skript laeuft von oben nach unten durch. Ein Aufruf vor der
Definition oder ein Tippfehler faellt deshalb erst zur Laufzeit auf - und nur
dann, wenn der betreffende Zweig auch gezeichnet wird. Genau das pruefen
diese Tests: einmal ohne Daten und einmal mit einem echten Importergebnis.
"""

from pathlib import Path

import pytest

from vdi3814.config import SETTINGS
from vdi3814.pipeline import process_file

APP = str(Path(__file__).resolve().parents[1] / "vdi3814" / "ui" / "app.py")


@pytest.fixture()
def eigene_ablage(tmp_path, monkeypatch):
    """Projekte in ein Testverzeichnis umlenken, nichts anfassen."""
    monkeypatch.setattr(SETTINGS, "data_dir", tmp_path)
    return tmp_path


def _app(timeout: int = 240):
    from streamlit.testing.v1 import AppTest

    return AppTest.from_file(APP, default_timeout=timeout)


def test_oberflaeche_startet_ohne_daten(eigene_ablage):
    at = _app()
    at.run()
    assert not at.exception, [str(e.value) for e in at.exception]


def test_oberflaeche_zeichnet_das_importergebnis(eigene_ablage, samples):
    """Der Zweig, der erst nach einem Import sichtbar wird.

    Genau hier fehlte eine Hilfsfunktion, die weiter unten im Skript stand -
    ohne Import war der Zweig nie gelaufen und der Fehler unentdeckt.
    """
    ergebnis = process_file(samples["alt"])

    at = _app()
    at.session_state["results"] = {"beispiel_alte_fassung.pdf": ergebnis}
    at.run()

    assert not at.exception, [str(e.value) for e in at.exception]

    # Die Ergebnistabelle muss die Summenprüfung im Klartext zeigen
    texte = " ".join(str(d.value.to_dict()) for d in at.dataframe if hasattr(d.value, "to_dict"))
    assert "Seiten exakt" in texte or "abweichend" in texte


def test_oberflaeche_mit_gespeicherten_daten(eigene_ablage, samples):
    """Uebersicht, Kosten und Export zeichnen mit Daten in der Datenbank."""
    from sqlalchemy.orm import Session

    from vdi3814 import db, projects

    projekt = projects.resolve(None)
    engine = db.make_engine(projekt.db_path)
    with Session(engine) as session:
        db.save_document(session, process_file(samples["alt"]))
        db.save_document(session, process_file(samples["neu"]))

    at = _app()
    at.run()
    assert not at.exception, [str(e.value) for e in at.exception]

    beschriftungen = [m.label for m in at.metric]
    assert "Funktionen gesamt" in beschriftungen


def test_oberflaeche_zeigt_abgelegte_kostenschaetzungen(eigene_ablage, samples):
    """Reiter 5 mit einem abgelegten Stand: Liste, Auswahl und Positionen.

    Dieser Zweig wird erst sichtbar, wenn im Projekt schon eine
    Kostenschaetzung liegt - ohne Test liefe er nie mit.
    """
    from sqlalchemy.orm import Session

    from vdi3814 import db, projects

    projekt = projects.resolve(None)
    engine = db.make_engine(projekt.db_path)
    with Session(engine) as session:
        db.save_document(session, process_file(samples["alt"]))
        db.save_cost_estimate(session, "Angebot", [
            {"spalte_adresse": "1.1", "gruppe": "Ein-/Ausgabefunktionen", "spalte": "BI",
             "spalte_key": "A_1_1", "menge": 10.0, "einheitspreis": 45.0},
        ], bemerkung="Erstkalkulation")

    at = _app()
    at.run()
    assert not at.exception, [str(e.value) for e in at.exception]

    texte = " ".join(str(d.value.to_dict()) for d in at.dataframe if hasattr(d.value, "to_dict"))
    assert "Angebot" in texte
    assert "Erstkalkulation" in texte


# --------------------------------------------------------------------------
# Reiter 2: ganze Zeilen loeschen
# --------------------------------------------------------------------------
#
# st.data_editor laesst sich mit AppTest nicht bedienen. Geprueft wird deshalb
# die Umrechnung, die genau dazwischen liegt: angekreuzte Tabellenzeile ->
# Position -> entfernte Datenzeile. Das ist die Stelle, an der ein Versatz
# zwischen Anzeige und Datenmodell die falsche Zeile treffen wuerde.


def test_korrekturtabelle_hat_eine_ankreuzspalte(samples):
    from vdi3814.ui.table import DELETE_FIELD, result_to_frame

    ergebnis = process_file(samples["alt"])
    frame = result_to_frame(ergebnis)

    assert list(frame.columns)[0] == DELETE_FIELD
    assert not frame[DELETE_FIELD].any()
    assert len(frame) == len(ergebnis.data_rows())


def test_angekreuzte_zeilen_werden_zu_positionen(samples):
    from vdi3814.ui.table import DELETE_FIELD, marked_rows, result_to_frame

    ergebnis = process_file(samples["alt"])
    frame = result_to_frame(ergebnis)
    assert marked_rows(frame) == []

    frame.loc[0, DELETE_FIELD] = True
    frame.loc[2, DELETE_FIELD] = True

    assert marked_rows(frame) == [0, 2]


def test_mehrere_zeilen_loeschen_trifft_genau_die_angekreuzten(samples):
    """Der ganze Weg: ankreuzen, Positionen bestimmen, Zeilen entfernen."""
    from vdi3814.ui.table import DELETE_FIELD, marked_rows, result_to_frame

    ergebnis = process_file(samples["alt"])
    vorher = [r.klartext for r in ergebnis.data_rows()]
    assert len(vorher) >= 3

    frame = result_to_frame(ergebnis)
    frame.loc[0, DELETE_FIELD] = True
    frame.loc[2, DELETE_FIELD] = True
    entfernt = ergebnis.remove_data_rows(marked_rows(frame))

    assert entfernt == 2
    assert [r.klartext for r in ergebnis.data_rows()] == [vorher[1]] + vorher[3:]
    assert list(result_to_frame(ergebnis)["Datenpunkt"]) == [vorher[1]] + vorher[3:]


def test_einzelne_zeile_loeschen_senkt_die_summe(samples):
    ergebnis = process_file(samples["alt"])
    gezaehlt = [i for i, r in enumerate(ergebnis.data_rows()) if r.is_countable]
    assert gezaehlt

    summe_vorher = ergebnis.grand_total()
    wert_der_zeile = sum(c.count or 0 for c in ergebnis.data_rows()[gezaehlt[0]].cells)

    assert ergebnis.remove_data_rows([gezaehlt[0]]) == 1
    assert ergebnis.grand_total() == pytest.approx(summe_vorher - wert_der_zeile)


def test_loeschen_rechnet_die_summenpruefung_neu(samples):
    """Nach dem Entfernen darf die Summenpruefung nicht veraltet stehenbleiben."""
    from vdi3814.pipeline import validate_sums

    ergebnis = process_file(samples["alt"])
    gezaehlt = [i for i, r in enumerate(ergebnis.data_rows()) if r.is_countable]
    assert not [c for c in ergebnis.sum_checks if not c.matches]

    ergebnis.remove_data_rows([gezaehlt[0]])
    ergebnis.sum_checks = validate_sums(ergebnis)

    # Die Summenzeile des Dokuments passt jetzt bewusst nicht mehr - genau das
    # soll der Anwender im Reiter "Nachweis & Differenzen" sehen.
    assert [c for c in ergebnis.sum_checks if not c.matches]


def test_loeschen_laesst_summen_und_leerzeilen_stehen(samples):
    """Nur Datenzeilen verschwinden - sonst wuerde sich die Auswahl verschieben."""
    ergebnis = process_file(samples["alt"])
    nicht_gezaehlt = [(r.kind, r.row_no) for r in ergebnis.excluded_rows()]
    assert nicht_gezaehlt

    ergebnis.remove_data_rows([0])

    assert [(r.kind, r.row_no) for r in ergebnis.excluded_rows()] == nicht_gezaehlt


def test_loeschen_ignoriert_positionen_ausserhalb_der_tabelle(samples):
    ergebnis = process_file(samples["alt"])
    anzahl = len(ergebnis.rows)

    assert ergebnis.remove_data_rows([]) == 0
    assert ergebnis.remove_data_rows([999, -1]) == 0
    assert len(ergebnis.rows) == anzahl


def test_geloeschte_zeilen_landen_nicht_in_der_datenbank(eigene_ablage, samples):
    """Nach dem Speichern muss die Auswertung ohne die geloeschten Zeilen rechnen."""
    from sqlalchemy.orm import Session

    from vdi3814 import db, projects

    ergebnis = process_file(samples["alt"])
    gezaehlt = [i for i, r in enumerate(ergebnis.data_rows()) if r.is_countable]
    ergebnis.remove_data_rows(gezaehlt[:1])
    erwartet = len(ergebnis.data_rows())

    projekt = projects.resolve(None)
    engine = db.make_engine(projekt.db_path)
    with Session(engine) as session:
        dokument = db.save_document(session, ergebnis)
        geladen = db.load_document(session, dokument.id)

    assert len(geladen.data_rows()) == erwartet
    assert geladen.grand_total() == ergebnis.grand_total()


def test_korrekturtabelle_traegt_den_schwerpunkt_mit(samples):
    """Die ASP/ISP-Spalte muss den Umweg ueber die Tabelle unbeschadet ueberstehen.

    Anzeige und Ruecklesen liegen in vdi3814.ui.table - genau dort trafen die
    Ankreuzspalte und die Schwerpunktspalte aufeinander.
    """
    from vdi3814.ui.table import ROW_FIELDS, frame_to_result, result_to_frame

    ergebnis = process_file(samples["asp"])
    assert "Schwerpunkt" in ROW_FIELDS

    frame = result_to_frame(ergebnis)
    assert [z for z in frame["Schwerpunkt"] if z], "kein Schwerpunkt in der Tabelle"

    frame.loc[0, "Schwerpunkt"] = "ASP07 Kaeltetechnik"
    frame_to_result(frame, ergebnis)
    erste = ergebnis.data_rows()[0]

    assert erste.schwerpunkt == "ASP07"
    assert erste.schwerpunkt_text == "Kaeltetechnik"
    assert result_to_frame(ergebnis).loc[0, "Schwerpunkt"].startswith("ASP07")
