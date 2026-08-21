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
