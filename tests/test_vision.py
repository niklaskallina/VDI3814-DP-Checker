"""Vision-Pfad mit Fixture-Backend - prueft Prompt-Auswertung und Banding-Merge."""

import json

from PIL import Image

from vdi3814.config import Settings
from vdi3814.extract.vision_extractor import classify_page, extract_vision
from vdi3814.models import PageKind
from vdi3814.vision.fixture_backend import FixtureVisionBackend


def _fixture(tmp_path, data) -> FixtureVisionBackend:
    path = tmp_path / "antworten.json"
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return FixtureVisionBackend(path)


def test_klassifikation_erkennt_schema(tmp_path):
    backend = _fixture(tmp_path, {"classify": {"typ": "regelschema", "konfidenz": 0.9,
                                               "begruendung": "Prinzipschaltbild"}})
    result = classify_page(backend, Image.new("RGB", (400, 300), "white"))
    assert result.kind is PageKind.SCHEMA
    assert result.is_table is False


def test_vision_extraktion_mit_banding(tmp_path):
    """Zwei Streifen liefern ueberlappende Zeilen - Dubletten muessen verschmelzen."""
    header = {
        "spalten": [
            {"index": 0, "gruppe": "1. Ein-/Ausgabefunktionen", "untergruppe": "Physikalische",
             "bezeichnung": "Analoge Eingabe (AI)", "abschnitt": "1.1.", "spalte_nr": "1"},
            {"index": 1, "gruppe": "2. Anwendungsfunktionen", "untergruppe": "Regelung",
             "bezeichnung": "unleserlich", "abschnitt": "2.3.", "spalte_nr": "1"},
        ],
        "fussnoten": [{"marker": "1", "text": "pro Eingangs-Benutzeradresse"}],
    }
    band_one = {"zeilen": [
        {"zeile_nr": "1", "klartext": "Vorlauftemperatur", "bas": "*T01", "werte": {"0": "1"}},
        {"zeile_nr": "2", "klartext": "Regler", "werte": {}},
    ]}
    band_two = {"zeilen": [
        # dieselbe Zeile 2, hier aber mit gelesenem Wert -> muss gewinnen
        {"zeile_nr": "2", "klartext": "Regler", "werte": {"1": "3"}},
        {"zeile_nr": "3", "klartext": "Bemerkung frei", "werte": {"0": "Freitext"},
         "bemerkung": "Darstellungsvariante 1"},
        {"klartext": "Summe Funktionen", "werte": {"0": "1", "1": "3"}, "summenzeile": True},
    ]}
    backend = _fixture(tmp_path, {
        "header": header, "rows#0": band_one, "rows#1": band_two,
        "metadata": {"anlage": "RLT 01", "gewerk": "GA", "protokoll": "BACnet/IP"},
    })
    settings = Settings()
    settings.band_count = 2
    table = extract_vision(backend, Image.new("RGB", (900, 1200), "white"), 0, settings)

    assert [c.address for c in table.columns] == ["1.1.1", "2.3.1"]
    data = [r for r in table.rows if not r.is_sum_row]
    assert [r.klartext for r in data] == ["Vorlauftemperatur", "Regler", "Bemerkung frei"]
    assert data[1].value(1) == 3.0                    # Wert aus dem zweiten Streifen
    # Freitext in einer Zaehlspalte wird als Notiz gefuehrt, nicht als Zahl
    assert data[2].cells[0].count is None
    assert data[2].cells[0].note == "Freitext"
    assert any(r.is_sum_row for r in table.rows)
    assert table.metadata.protokoll == "BACnet/IP"
    assert table.footnotes[0].marker == "1"


def test_leere_modellantwort_fuehrt_zu_hinweis(tmp_path):
    backend = _fixture(tmp_path, {})
    table = extract_vision(backend, Image.new("RGB", (600, 800), "white"), 0)
    assert table.columns == []
    assert table.warnings and "Spalten" in table.warnings[0]
