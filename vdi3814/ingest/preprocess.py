"""Bildvorverarbeitung vor der Modellabfrage.

Ziel: schiefe/kontrastarme Scans so aufbereiten, dass das lokale Vision-Modell
und Tesseract die feinen Tabellenraster zuverlaessig lesen - ohne sich auf
feste Pixelkoordinaten zu verlassen.
"""

from __future__ import annotations

import numpy as np
from PIL import Image, ImageOps

from ..config import SETTINGS


def to_grayscale(image: Image.Image) -> Image.Image:
    return image.convert("L") if image.mode != "L" else image


def estimate_skew(image: Image.Image, max_angle: float = 4.0, step: float = 0.25) -> float:
    """Schaetzt den Schiefwinkel ueber die Varianz der Zeilen-Schwarzanteile.

    Bei geradem Text ist die Projektion auf die y-Achse stark strukturiert
    (Textzeilen = Peaks); bei schiefem Text verschmiert sie. Wir suchen den
    Winkel mit maximaler Varianz.
    """
    gray = to_grayscale(image)
    small = gray.copy()
    small.thumbnail((1000, 1000))
    best_angle, best_score = 0.0, -1.0
    angle = -max_angle
    while angle <= max_angle + 1e-9:
        rotated = small.rotate(angle, resample=Image.BILINEAR, fillcolor=255)
        arr = 255 - np.asarray(rotated, dtype=np.float32)
        profile = arr.sum(axis=1)
        score = float(np.var(profile))
        if score > best_score:
            best_score, best_angle = score, angle
        angle += step
    return best_angle


def deskew(image: Image.Image, max_angle: float = 4.0) -> Image.Image:
    angle = estimate_skew(image, max_angle=max_angle)
    if abs(angle) < 0.2:
        return image
    return image.rotate(angle, resample=Image.BICUBIC, fillcolor=(255, 255, 255) if image.mode == "RGB" else 255)


def limit_size(image: Image.Image, max_edge: int) -> Image.Image:
    if max(image.size) <= max_edge:
        return image
    ratio = max_edge / max(image.size)
    new_size = (max(1, int(image.width * ratio)), max(1, int(image.height * ratio)))
    return image.resize(new_size, Image.LANCZOS)


def prepare_for_vision(image: Image.Image, settings=SETTINGS) -> Image.Image:
    """Standard-Aufbereitung: Deskew -> Kontrast -> Groessenbegrenzung."""
    result = image
    if settings.deskew:
        result = deskew(result)
    if settings.autocontrast:
        result = ImageOps.autocontrast(to_grayscale(result), cutoff=1).convert("RGB")
    return limit_size(result, settings.max_edge_px)


def prepare_for_ocr(image: Image.Image, settings=SETTINGS) -> Image.Image:
    """Fuer Tesseract: Graustufen, Kontrast, moderat hochskaliert."""
    gray = ImageOps.autocontrast(to_grayscale(deskew(image) if settings.deskew else image), cutoff=1)
    if gray.width < 1600:
        factor = 1600 / gray.width
        gray = gray.resize((int(gray.width * factor), int(gray.height * factor)), Image.LANCZOS)
    return gray


# --------------------------------------------------------------------------
# Zuschnitte: Kopf-, Fussbereich und horizontale Streifen ("Banding")
# --------------------------------------------------------------------------

def crop_header(image: Image.Image, fraction: float | None = None) -> Image.Image:
    frac = SETTINGS.header_fraction if fraction is None else fraction
    height = max(1, int(image.height * frac))
    return image.crop((0, 0, image.width, height))


def crop_footer(image: Image.Image, fraction: float | None = None) -> Image.Image:
    frac = SETTINGS.footer_fraction if fraction is None else fraction
    height = max(1, int(image.height * frac))
    return image.crop((0, image.height - height, image.width, image.height))


def crop_band(image: Image.Image, band_index: int, band_count: int, overlap: float = 0.08,
              body_top: float = 0.0) -> Image.Image:
    """Schneidet den band_index-ten horizontalen Streifen (mit Ueberlappung).

    body_top: Anteil der Seitenhoehe, der zum Tabellenkopf gehoert und beim
    Banding uebersprungen wird (die Zeilen liegen darunter).
    """
    band_count = max(1, band_count)
    top_offset = image.height * max(0.0, min(0.9, body_top))
    usable = image.height - top_offset
    band_height = usable / band_count
    top = top_offset + band_index * band_height - band_height * overlap
    bottom = top_offset + (band_index + 1) * band_height + band_height * overlap
    top = int(max(top_offset, top))
    bottom = int(min(image.height, bottom))
    return image.crop((0, top, image.width, bottom))


def stack_with_header(header: Image.Image, band: Image.Image) -> Image.Image:
    """Setzt den Tabellenkopf ueber einen Zeilenstreifen.

    Damit sieht das Modell in jedem Teilbild die Spaltenzuordnung - das ist
    der entscheidende Trick, um breite Funktionslisten streifenweise korrekt
    auszulesen.
    """
    width = max(header.width, band.width)
    canvas = Image.new("RGB", (width, header.height + band.height), "white")
    canvas.paste(header.convert("RGB"), (0, 0))
    canvas.paste(band.convert("RGB"), (0, header.height))
    return canvas
