"""Regressionstests fuer die drei Fehlerquellen aus den echten Plansaetzen.

1. Summenzeilen ohne Beschriftung wurden als Datenpunkte gezaehlt
   -> die Seite zaehlte doppelt.
2. Seiten mit /Rotate 90 wurden gar nicht erkannt.
3. Teilschriften ohne gueltige Zeichenzuordnung lieferten keine Zahlen.
"""

import pymupdf
import pytest

from vdi3814.extract.pdftext_extractor import _repariere_zeichen, extract_pdf_text
from vdi3814.extract.rowfilter import markiere_unbeschriftete_summenzeilen
from vdi3814.models import Cell, DataPointRow, RowKind
from vdi3814.pipeline import process_file


def _zeile(label, werte, kind=RowKind.DATEN):
    return DataPointRow(
        klartext=label,
        kind=kind,
        cells=[Cell(column_index=i, raw_value=str(v), count=float(v)) for i, v in werte.items()],
    )


def test_summenzeile_ohne_beschriftung_wird_erkannt():
    rows = [
        _zeile("Vorlauftemperatur", {0: 1, 2: 3}),
        _zeile("Pumpe", {0: 2, 2: 4}),
        _zeile("", {0: 3, 2: 7}),          # Summe der beiden Zeilen darueber
    ]
    markiere_unbeschriftete_summenzeilen(rows)
    assert [r.kind for r in rows] == [RowKind.DATEN, RowKind.DATEN, RowKind.SUMME]
    assert "Summe" in rows[2].exclusion_reason
    assert sum(c.count for r in rows if r.is_countable for c in r.cells) == 10


def test_echte_datenzeile_wird_nicht_faelschlich_zur_summe():
    """Nur eine Uebereinstimmung in ALLEN belegten Spalten gilt als Summe."""
    rows = [
        _zeile("Ventil", {0: 1, 2: 3}),
        _zeile("Pumpe", {0: 2, 2: 4}),
        _zeile("Klappe", {0: 3, 2: 9}),    # 3 passt, 9 passt nicht -> Datenpunkt
    ]
    markiere_unbeschriftete_summenzeilen(rows)
    assert all(r.kind is RowKind.DATEN for r in rows)


def test_einzelne_zeile_wird_nie_zur_summe():
    rows = [_zeile("Fuehler", {0: 1, 1: 1})]
    markiere_unbeschriftete_summenzeilen(rows)
    assert rows[0].kind is RowKind.DATEN


def test_nach_summenzeile_beginnt_ein_neuer_block():
    rows = [
        _zeile("A", {0: 1, 1: 1}),
        _zeile("B", {0: 1, 1: 1}),
        _zeile("", {0: 2, 1: 2}),          # Summe Block 1
        _zeile("C", {0: 5, 1: 5}),
        _zeile("D", {0: 5, 1: 5}),
        _zeile("", {0: 10, 1: 10}),        # Summe Block 2
    ]
    markiere_unbeschriftete_summenzeilen(rows)
    arten = [r.kind for r in rows]
    assert arten[2] is RowKind.SUMME and arten[5] is RowKind.SUMME
    assert sum(c.count for r in rows if r.is_countable for c in r.cells) == 24


def _wie_plansatz_gedreht(quelle, ziel):
    """Baut eine Datei nach, wie sie CAD-Plansaetze liefern.

    Hochformat-Seite, Inhalt quer hineingezeichnet, dazu /Rotate 90 - angezeigt
    wird also Querformat. PyMuPDF liefert Text und Zeichnungen dann im
    ungedrehten System, weshalb solche Dateien ohne Normalisierung gar nicht
    erkannt wurden.
    """
    with pymupdf.open(quelle) as src, pymupdf.open() as out:
        for index in range(src.page_count):
            masse = src[index].rect
            seite = out.new_page(width=masse.height, height=masse.width)
            seite.show_pdf_page(seite.rect, src, index, rotate=90)
            seite.set_rotation(90)
        out.save(ziel)
    return ziel


def test_gedrehte_seite_wird_wie_die_ungedrehte_gelesen(samples, tmp_path):
    """Ein Plansatz mit /Rotate 90 muss dasselbe Ergebnis liefern."""
    original = process_file(samples["alt"])
    gedreht = _wie_plansatz_gedreht(samples["alt"], tmp_path / "gedreht.pdf")

    with pymupdf.open(gedreht) as pruefung:
        assert pruefung[0].rotation == 90, "Testdatei ist nicht gedreht"

    ergebnis = process_file(gedreht)
    assert ergebnis.fassung == original.fassung
    assert len(ergebnis.counted_rows()) == len(original.counted_rows())
    assert ergebnis.grand_total() == original.grand_total()
    assert [r.klartext for r in ergebnis.counted_rows()] == \
           [r.klartext for r in original.counted_rows()]


@pytest.mark.parametrize("roh,erwartet", [
    ("\x14", "1"),
    ("\x13\x13\x14", "001"),
    ("\x15\x1c", "29"),
    ("Pumpe", "Pumpe"),          # unveraendert, keine Steuerzeichen
    ("", ""),
])
def test_defekte_zeichenzuordnung_wird_zurueckgerechnet(roh, erwartet):
    assert _repariere_zeichen(roh) == erwartet


def test_seitensummen_stimmen_im_beispiel(samples):
    """Die Zeile 'Summe Funktionen' muss zur eigenen Zaehlung passen."""
    ergebnis = process_file(samples["alt"])
    assert ergebnis.sum_checks, "es muss mindestens eine Seite geprueft werden"
    assert all(pruefung.matches for pruefung in ergebnis.sum_checks)


def test_summenblatt_wird_nicht_gezaehlt(tmp_path):
    """Seiten mit ausschliesslich Anlagensummen zaehlen nicht mit."""
    from vdi3814.extract.rowfilter import classify_row

    kind, grund = classify_row("Summe 253002101O2432002_ RLT02", has_values=True)
    assert kind is RowKind.SUMME and grund
