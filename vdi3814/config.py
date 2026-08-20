"""Zentrale Konfiguration.

Alle Werte lassen sich ueber Umgebungsvariablen ueberschreiben (Praefix VDI_),
damit auf dem Zielrechner nichts im Code angepasst werden muss.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path


def _env(name: str, default: str) -> str:
    return os.environ.get(f"VDI_{name}", default)


def _env_int(name: str, default: int) -> int:
    try:
        return int(_env(name, str(default)))
    except ValueError:
        return default


def _env_float(name: str, default: float) -> float:
    try:
        return float(_env(name, str(default)))
    except ValueError:
        return default


APP_DIR = Path(__file__).resolve().parent
PROJECT_DIR = APP_DIR.parent
PROFILE_DIR = Path(_env("PROFILE_DIR", str(APP_DIR / "profiles")))


@dataclass
class Settings:
    """Laufzeiteinstellungen des Tools."""

    # --- Datenhaltung ---
    data_dir: Path = field(default_factory=lambda: Path(_env("DATA_DIR", str(PROJECT_DIR / "data"))))
    db_path: Path = field(default_factory=lambda: Path(_env("DB_PATH", str(PROJECT_DIR / "data" / "vdi3814.sqlite3"))))

    # --- Rasterisierung / Bildvorverarbeitung ---
    render_dpi: int = _env_int("RENDER_DPI", 300)
    max_edge_px: int = _env_int("MAX_EDGE_PX", 2200)   # Downscale fuer das VLM
    deskew: bool = _env("DESKEW", "1") == "1"
    autocontrast: bool = _env("AUTOCONTRAST", "1") == "1"

    # --- Lokales Vision-Language-Modell (Ollama, laeuft auf localhost) ---
    ollama_host: str = _env("OLLAMA_HOST", "http://127.0.0.1:11434")
    vision_model: str = _env("VISION_MODEL", "qwen2.5vl:7b")
    vision_timeout_s: int = _env_int("VISION_TIMEOUT", 600)
    vision_retries: int = _env_int("VISION_RETRIES", 2)
    vision_temperature: float = _env_float("VISION_TEMPERATURE", 0.0)
    vision_num_ctx: int = _env_int("VISION_NUM_CTX", 8192)

    # --- Banding: grosse Tabellen werden zeilenweise in Streifen zerlegt ---
    band_count: int = _env_int("BAND_COUNT", 3)
    band_overlap: float = _env_float("BAND_OVERLAP", 0.08)
    header_fraction: float = _env_float("HEADER_FRACTION", 0.30)
    footer_fraction: float = _env_float("FOOTER_FRACTION", 0.20)

    # --- Ergaenzendes lokales OCR ---
    use_ocr: bool = _env("USE_OCR", "1") == "1"
    tesseract_cmd: str = _env("TESSERACT_CMD", "")
    ocr_lang: str = _env("OCR_LANG", "deu")

    # --- Sonstiges ---
    currency: str = _env("CURRENCY", "EUR")

    def ensure_dirs(self) -> None:
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)


SETTINGS = Settings()
