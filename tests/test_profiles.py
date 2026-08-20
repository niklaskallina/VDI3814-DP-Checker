from vdi3814.models import ColumnHeader
from vdi3814.profiles_loader import detect_profile, get_profile, load_profiles, match_columns


def test_beide_fassungen_vorhanden():
    ids = {p.id for p in load_profiles()}
    assert {"vdi3814_alt", "vdi3814_neu"} <= ids


def test_spaltenzahlen_entsprechen_den_vorlagen():
    assert len(get_profile("vdi3814_alt").columns) == 50      # Beiblatt 070-5
    assert len(get_profile("vdi3814_neu").columns) == 56      # Vorlage 2022-07


def test_adresslookup():
    alt = get_profile("vdi3814_alt")
    assert alt.column_by_address("6.13").label == "Tarifabhängiges Schalten"
    neu = get_profile("vdi3814_neu")
    assert neu.column_by_address("2.2.10").label == "Frostschutzsteuerung"
    assert neu.column_by_address("1.1.1").abbrev == "AI"


def test_fassungserkennung():
    alt, score, _ = detect_profile([
        "Ein-/Ausgabefunktionen Verarbeitungsfunktionen Überwachen Steuern Regeln "
        "Rechnen/Optimieren Managementfunktionen Bedienfunktionen"
    ])
    assert alt is not None and alt.fassung == "alt" and score >= alt.min_score

    neu, _, _ = detect_profile([
        "0. Integration 1. Ein-/Ausgabefunktionen 2. Anwendungsfunktionen Sonnenschutz "
        "Beleuchtung 3. B-/A-Funktionen Bemerkungen und Referenzierungen"
    ])
    assert neu is not None and neu.fassung == "neu"


def test_unbekanntes_layout_wird_nicht_erzwungen():
    profile, _, _ = detect_profile(["Materialliste Kabelquerschnitt Klemmenplan"])
    assert profile is None


def test_zuordnung_ueber_adresse_schlaegt_label():
    profile = get_profile("vdi3814_neu")
    headers = [
        # Bezeichnung unlesbar, Adresse aber vorhanden -> muss trotzdem sitzen
        ColumnHeader(index=0, label="unleserlich", section="2.2", column_no="10"),
        ColumnHeader(index=1, label="Kalender (CAL)"),
        ColumnHeader(index=2, label="Bemerkungen und Referenzierungen"),
    ]
    match_columns(headers, profile)
    assert headers[0].column_key == "N_2_2_10"
    assert headers[1].column_key == "N_1_3_2"
    assert headers[2].is_note_column is True


def test_jede_profilspalte_wird_hoechstens_einmal_vergeben():
    profile = get_profile("vdi3814_alt")
    headers = [ColumnHeader(index=i, label="Motorsteuerung") for i in range(3)]
    match_columns(headers, profile)
    keys = [h.column_key for h in headers if h.column_key]
    assert len(keys) == len(set(keys)) == 1
