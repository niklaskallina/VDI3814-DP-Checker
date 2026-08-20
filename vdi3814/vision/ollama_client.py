"""Client fuer ein lokal laufendes Ollama-Modell (localhost, keine Cloud).

Es werden ausschliesslich Requests an SETTINGS.ollama_host gesendet
(Voreinstellung http://127.0.0.1:11434). Es verlaesst kein Projektdatum
das Geraet.
"""

from __future__ import annotations

import base64
import io
import json
import logging
import re
import time
from typing import Any

import requests
from PIL import Image

from ..config import SETTINGS
from .base import VisionError

log = logging.getLogger(__name__)

_JSON_BLOCK = re.compile(r"\{.*\}", re.DOTALL)


class OllamaUnavailable(VisionError):
    """Ollama laeuft nicht oder das Modell ist nicht installiert."""


def _encode(image: Image.Image) -> str:
    buffer = io.BytesIO()
    image.convert("RGB").save(buffer, format="PNG", optimize=True)
    return base64.b64encode(buffer.getvalue()).decode("ascii")


def _parse_json(text: str) -> dict[str, Any]:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    match = _JSON_BLOCK.search(text)
    if match:
        try:
            return json.loads(match.group(0))
        except json.JSONDecodeError as exc:
            raise VisionError(f"Modellantwort ist kein gueltiges JSON: {exc}") from exc
    raise VisionError("Modellantwort enthaelt kein JSON")


class OllamaVisionBackend:
    """Vision-Backend ueber die lokale Ollama-HTTP-API."""

    def __init__(self, host: str | None = None, model: str | None = None, settings=SETTINGS):
        self.settings = settings
        self.host = (host or settings.ollama_host).rstrip("/")
        self.model = model or settings.vision_model
        self.name = f"ollama:{self.model}"
        self._session = requests.Session()
        # Kein Proxy fuer localhost - sonst scheitert der Aufruf in Firmennetzen.
        self._session.trust_env = False

    # ---------------- Verfuegbarkeit ----------------

    def list_models(self) -> list[str]:
        try:
            resp = self._session.get(f"{self.host}/api/tags", timeout=10)
            resp.raise_for_status()
        except requests.RequestException as exc:
            raise OllamaUnavailable(
                f"Ollama unter {self.host} nicht erreichbar. Laeuft 'ollama serve'? ({exc})"
            ) from exc
        return [m.get("name", "") for m in resp.json().get("models", [])]

    def available(self) -> bool:
        try:
            models = self.list_models()
        except OllamaUnavailable:
            return False
        base = self.model.split(":")[0]
        return any(m == self.model or m.split(":")[0] == base for m in models)

    def ensure_ready(self) -> None:
        models = self.list_models()          # wirft OllamaUnavailable
        base = self.model.split(":")[0]
        if not any(m == self.model or m.split(":")[0] == base for m in models):
            raise OllamaUnavailable(
                f"Modell '{self.model}' ist nicht installiert. "
                f"Bitte einmalig ausfuehren:  ollama pull {self.model}\n"
                f"Vorhanden: {', '.join(models) or '(keine)'}"
            )

    # ---------------- Abfrage ----------------

    def ask_json(self, image: Image.Image, prompt: str, *, purpose: str = "") -> dict[str, Any]:
        payload = {
            "model": self.model,
            "prompt": prompt,
            "images": [_encode(image)],
            "stream": False,
            "format": "json",           # erzwingt JSON-Ausgabe
            "options": {
                "temperature": self.settings.vision_temperature,
                "num_ctx": self.settings.vision_num_ctx,
            },
        }
        last_error: Exception | None = None
        for attempt in range(self.settings.vision_retries + 1):
            try:
                response = self._session.post(
                    f"{self.host}/api/generate",
                    json=payload,
                    timeout=self.settings.vision_timeout_s,
                )
                response.raise_for_status()
                text = response.json().get("response", "")
                return _parse_json(text)
            except (requests.RequestException, VisionError) as exc:
                last_error = exc
                log.warning("Vision-Abfrage '%s' fehlgeschlagen (Versuch %d): %s",
                            purpose or "?", attempt + 1, exc)
                time.sleep(min(2 ** attempt, 5))
        raise VisionError(f"Vision-Abfrage '{purpose}' endgueltig fehlgeschlagen: {last_error}")
