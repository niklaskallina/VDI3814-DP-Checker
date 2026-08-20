"""Schnittstelle der Vision-Backends.

Die Pipeline kennt nur dieses Protokoll. Damit laesst sich das lokale
Ollama-Modell gegen ein anderes lokales Modell (oder gegen die
Test-/Demo-Fixtures) tauschen, ohne den Erkennungskern zu aendern.
"""

from __future__ import annotations

from typing import Any, Protocol, runtime_checkable

from PIL import Image


class VisionError(RuntimeError):
    """Fehler bei der Modellabfrage (Timeout, ungueltiges JSON, Modell fehlt)."""


@runtime_checkable
class VisionBackend(Protocol):
    name: str

    def ask_json(self, image: Image.Image, prompt: str, *, purpose: str = "") -> dict[str, Any]:
        """Schickt ein Bild + Prompt an das Modell und liefert die JSON-Antwort."""
        ...

    def available(self) -> bool:
        """True, wenn das Backend einsatzbereit ist (Modell geladen o. Ae.)."""
        ...
