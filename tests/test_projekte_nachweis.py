"""Projekte, Loeschen von Datensaetzen und der Nachweis im Originaldokument."""

import pytest
from sqlalchemy.orm import Session

from vdi3814 import db, evidence, projects
from vdi3814.config import SETTINGS
from vdi3814.pipeline import process_file


@pytest.fixture()
def projektordner(tmp_path, monkeypatch):
    monkeypatch.setattr(SETTINGS, "data_dir", tmp_path)
    return tmp_path


def test_projekte_sind_getrennt(projektordner, samples):
    eins = projects.create("Projekt A")
    zwei = projects.create("Projekt B")
    assert eins.db_path != zwei.db_path

    with Session(db.make_engine(eins.db_path)) as session:
        db.save_document(session, process_file(samples["alt"]))
        assert len(db.list_documents(session)) == 1
    with Session(db.make_engine(zwei.db_path)) as session:
        assert db.list_documents(session) == []       # keine Vermischung

    assert {p.name for p in projects.list_projects()} == {"Projekt A", "Projekt B"}


def test_projekt_leeren_und_loeschen(projektordner, samples):
    projekt = projects.create("Zum Loeschen")
    with Session(db.make_engine(projekt.db_path)) as session:
        db.save_document(session, process_file(samples["alt"]))

    assert projects.clear("Zum Loeschen") is True
    assert not projekt.db_path.exists()
    assert projekt.path.exists()                      # Projekt bleibt bestehen

    assert projects.delete("Zum Loeschen") is True
    assert not projekt.path.exists()


def test_einzelne_datei_loeschen(projektordner, samples):
    projekt = projects.create("Loeschtest")
    with Session(db.make_engine(projekt.db_path)) as session:
        erstes = db.save_document(session, process_file(samples["alt"]))
        db.save_document(session, process_file(samples["neu"]))
        assert len(db.list_documents(session)) == 2

        db.delete_document(session, erstes.id)
        verbleibend = db.list_documents(session)
        assert len(verbleibend) == 1
        assert verbleibend[0].fassung == "neu"
        # Zeilen und Werte der geloeschten Datei sind ebenfalls weg
        assert session.query(db.Row).filter(db.Row.document_id == erstes.id).count() == 0


def test_originaldatei_wird_im_projekt_abgelegt(projektordner, samples):
    projekt = projects.create("Nachweis")
    with Session(db.make_engine(projekt.db_path)) as session:
        dokument = db.save_document(session, process_file(samples["alt"]))
        assert dokument.stored_path
        assert evidence.source_file(dokument) is not None


def test_fundstelle_wird_gespeichert_und_gezeichnet(projektordner, samples):
    projekt = projects.create("Nachweis")
    with Session(db.make_engine(projekt.db_path)) as session:
        dokument = db.save_document(session, process_file(samples["alt"]))
        zeilen = [r for r in dokument.rows if r.kind == "daten" and r.bbox]
        assert zeilen, "Datenzeilen muessen ihre Fundstelle mitfuehren"

        werte = [v for r in zeilen for v in r.values if v.count is not None and v.bbox]
        assert werte, "Auch einzelne Werte muessen eine Fundstelle haben"

        bbox = db.bbox_from_text(zeilen[0].bbox)
        bild = evidence.render_page(dokument, zeilen[0].page_index,
                                    highlights=[(bbox, "datenpunkt")], zoom_to=bbox)
        assert bild is not None and bild.width > 50


def test_befunde_melden_nur_pruefwuerdiges(projektordner, samples):
    projekt = projects.create("Befunde")
    with Session(db.make_engine(projekt.db_path)) as session:
        db.save_document(session, process_file(samples["alt"]))
        befunde = evidence.all_findings(session)
    # Eine stimmige Summenzeile ist kein Befund
    assert not [b for b in befunde if b.art == "abweichung"]
    assert not [b for b in befunde if "Summe Funktionen" in b.titel]


def test_projekt_loeschen_bei_offener_datenbank(projektordner, samples):
    """Unter Windows scheitert das Loeschen, wenn die Datei noch geoeffnet ist."""
    projekt = projects.create("Offen")
    engine = db.make_engine(projekt.db_path)
    with Session(engine) as session:
        db.save_document(session, process_file(samples["alt"]))
    # bewusst KEIN dispose durch den Aufrufer - das muss das Programm erledigen
    assert projects.delete("Offen") is True
    assert not projekt.path.exists()
