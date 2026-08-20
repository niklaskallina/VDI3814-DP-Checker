"""Vision-Backend fuer Tests und den Offline-Demolauf.

Statt ein Modell zu fragen, werden vorbereitete JSON-Antworten geliefert.
Damit laeuft der komplette Ablauf (Import -> Erkennung -> Vorschau -> Export)
reproduzierbar, auch ohne installiertes Ollama-Modell.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from PIL import Image


class FixtureVisionBackend:
    """Liefert Antworten aus einer JSON-Datei.

    Aufbau der Fixture-Datei:
        {"<zweck>": {...}}  oder  {"<zweck>#<n>": {...}} fuer mehrfache Aufrufe
    """

    def __init__(self, fixture_path: str | Path):
        self.path = Path(fixture_path)
        self.data: dict[str, Any] = json.loads(self.path.read_text(encoding="utf-8"))
        self.name = f"fixture:{self.path.name}"
        self._calls: dict[str, int] = {}

    def available(self) -> bool:
        return True

    def ask_json(self, image: Image.Image, prompt: str, *, purpose: str = "") -> dict[str, Any]:
        count = self._calls.get(purpose, 0)
        self._calls[purpose] = count + 1
        for key in (f"{purpose}#{count}", purpose):
            if key in self.data:
                return self.data[key]
        return {}
