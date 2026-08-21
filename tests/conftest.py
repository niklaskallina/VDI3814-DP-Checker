import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

SAMPLE_DIR = Path(__file__).resolve().parents[1] / "sample_data"


@pytest.fixture(scope="session")
def samples():
    """Erzeugt die Beispieldateien einmalig, falls sie noch nicht existieren."""
    dateien = {
        "alt": SAMPLE_DIR / "beispiel_alte_fassung.pdf",
        "neu": SAMPLE_DIR / "beispiel_neue_fassung.pdf",
        # Uebersichtsblatt mit einer Zeile je Automationsschwerpunkt
        "asp": SAMPLE_DIR / "beispiel_asp_uebersicht.pdf",
    }
    if not all(pfad.exists() for pfad in dateien.values()):
        sys.path.insert(0, str(SAMPLE_DIR))
        import make_samples

        make_samples.main()
    return dateien
