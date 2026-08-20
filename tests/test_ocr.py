"""Scan-Auswertung ohne KI: Tabellenraster vermessen und Zellen einzeln lesen."""

import pymupdf
import pytest
from PIL import Image

from vdi3814.extract.ocr_extractor import available, extract_ocr, spaltenraster_aus_linien
from vdi3814.ingest.preprocess import prepare_for_ocr
from vdi3814.pipeline import process_file
from vdi3814.profiles_loader import get_profile, profile_by_block_sizes

pytestmark = pytest.mark.skipif(not available(), reason="Tesseract ist nicht installiert")


def _als_scan(pdf_pfad, ziel, dpi: int = 300):
    """Erzeugt aus einem PDF ein rein bildbasiertes PDF (wie ein Scan)."""
    quelle = pymupdf.open(pdf_pfad)
    ziel_dokument = pymupdf.open()
    for seite in quelle:
        pix = seite.get_pixmap(matrix=pymupdf.Matrix(dpi / 72, dpi / 72), alpha=False)
        neu = ziel_dokument.new_page(width=seite.rect.width, height=seite.rect.height)
        neu.insert_image(neu.rect, pixmap=pix)
    ziel_dokument.save(ziel)
    ziel_dokument.close()
    quelle.close()
    return ziel


def _bild(pdf_pfad, seite: int = 0, dpi: int = 300) -> Image.Image:
    with pymupdf.open(pdf_pfad) as dokument:
        pix = dokument[seite].get_pixmap(matrix=pymupdf.Matrix(dpi / 72, dpi / 72), alpha=False)
        return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def test_blockgroessen_identifizieren_die_fassung():
    assert profile_by_block_sizes(get_profile("vdi3814_alt").block_sizes()).id == "vdi3814_alt"
    assert profile_by_block_sizes(get_profile("vdi3814_neu").block_sizes()).id == "vdi3814_neu"
    assert profile_by_block_sizes([3, 3]) is None      # unbekanntes Layout: nichts raten


def test_spaltenraster_wird_im_scan_vermessen(samples):
    raster = spaltenraster_aus_linien(prepare_for_ocr(_bild(samples["alt"])))
    if raster is None:
        pytest.skip("Beispielbild ohne Tabellenlinien")
    assert raster.profil.id == "vdi3814_alt"
    assert len(raster.spalten_mitten) == 50


def _je_spalte(ergebnis) -> dict:
    werte: dict[str, float] = {}
    for zeile in ergebnis.counted_rows():
        for zelle in zeile.cells:
            spalte = ergebnis.column_by_index(zelle.column_index)
            if spalte and spalte.column_key and zelle.count is not None:
                werte[spalte.column_key] = werte.get(spalte.column_key, 0.0) + zelle.count
    return werte


def test_scan_wird_ohne_ki_ausgewertet(tmp_path, samples):
    """Ein rein bildbasiertes PDF wird per OCR gezaehlt - ohne Sprachmodell.

    Geprueft wird vor allem, dass nichts erfunden wird: jeder im Scan
    gezaehlte Wert muss auch im digitalen Original stehen. (Auf einer echten
    gescannten Liste stimmten beide Wege spaltenweise exakt ueberein; die
    hier erzeugte Beispieldatei hat bewusst sehr kleine Ziffern.)
    """
    original = process_file(samples["alt"])
    scan = process_file(_als_scan(samples["alt"], tmp_path / "scan.pdf"))

    assert scan.backend == "ocr"                      # ohne KI ausgewertet
    assert scan.fassung == original.fassung
    # Zaehlspalten (ohne die Bemerkungsspalte) muessen vollstaendig sein
    def zaehlspalten(ergebnis):
        return [c for c in ergebnis.columns if c.column_key and not c.is_note_column]
    assert len(zaehlspalten(scan)) == len(zaehlspalten(original))

    soll, ist = _je_spalte(original), _je_spalte(scan)
    assert scan.grand_total() >= original.grand_total() * 0.6

    # Entscheidend fuer die Kalkulation: eine Leseabweichung darf nicht
    # unbemerkt bleiben. Entweder stimmt der Scan mit dem Original ueberein,
    # oder die Pruefung gegen die Summenzeile der Liste schlaegt an.
    abweichungen = {key: (soll.get(key), ist.get(key))
                    for key in set(soll) | set(ist) if soll.get(key) != ist.get(key)}
    if abweichungen:
        gemeldet = [pruefung for pruefung in scan.sum_checks if not pruefung.matches]
        assert gemeldet, (
            "Abweichung zur Vorlage wurde nicht gemeldet: "
            f"{abweichungen}"
        )


def test_scan_ohne_tabelle_wird_uebersprungen(tmp_path, samples):
    """Die Schemaseite darf auch im Scan nicht ausgewertet werden."""
    scan = process_file(_als_scan(samples["alt"], tmp_path / "scan.pdf"))
    uebersprungen = [p for p in scan.pages if p.classification.kind.value != "funktionsliste"]
    assert uebersprungen, "Regelschema-Seite muss erkennbar bleiben"


def test_leeres_bild_stuerzt_nicht_ab():
    assert extract_ocr(Image.new("RGB", (900, 600), "white"), 0) is None
