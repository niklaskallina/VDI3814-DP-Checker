from vdi3814.textutil import (
    extract_footnote_markers,
    is_freetext,
    join_address,
    normalize,
    normalize_address,
    parse_count,
    similarity,
)


def test_normalize_umlaute():
    assert normalize("Überwachen") == "ueberwachen"
    assert normalize("Grenzwert  gleitend") == "grenzwert gleitend"


def test_similarity_toleriert_umbrueche():
    assert similarity("Grenzwert fest", "Grenzwert\nfest") == 1.0
    assert similarity("Motorsteuerung", "Sonnenschutz") < 0.5


def test_parse_count():
    assert parse_count("16") == 16.0
    assert parse_count("x") == 1.0
    assert parse_count("2,5") == 2.5
    assert parse_count("") is None
    assert parse_count("Darstellungsvariante 1 = Anzahl") is None


def test_freitext_wird_nicht_gezaehlt():
    assert is_freetext("Darstellungsvariante 1") is True
    assert is_freetext("3") is False


def test_adressen():
    assert normalize_address("2.2.") == "2.2"
    assert normalize_address("1.0") == "1.0"
    assert normalize_address("abc") == ""
    assert join_address("1.1.", "3") == "1.1.3"
    assert join_address("0.", "1") == "0.1"


def test_fussnotenmarker():
    assert extract_footnote_markers("Stellausgabe 2-Punkt 6)") == ["6"]
    assert extract_footnote_markers("Komplexer Objekttyp 8) 9)") == ["8", "9"]
    assert extract_footnote_markers("Motorsteuerung") == []
