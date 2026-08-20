import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

SAMPLE_DIR = Path(__file__).resolve().parents[1] / "sample_data"


@pytest.fixture(scope="session")
def samples():
    """Erzeugt die Beispieldateien einmalig, falls sie noch nicht existieren."""
    alt = SAMPLE_DIR / "beispiel_alte_fassung.pdf"
    neu = SAMPLE_DIR / "beispiel_neue_fassung.pdf"
    if not alt.exists() or not neu.exists():
        sys.path.insert(0, str(SAMPLE_DIR))
        import make_samples

        make_samples.main()
    return {"alt": alt, "neu": neu}
