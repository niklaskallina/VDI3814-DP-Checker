"""Textnormalisierung und unscharfer Textvergleich.

Wird gebraucht, um die vom Vision-Modell gelesenen Spaltenueberschriften den
Profil-Spalten zuzuordnen - ohne feste Positionen und tolerant gegenueber
OCR-Fehlern, Umlautschreibweisen und Zeilenumbruechen im Tabellenkopf.
"""

from __future__ import annotations

import re
import unicodedata
from difflib import SequenceMatcher

_UMLAUT_MAP = {
    "ä": "ae", "ö": "oe", "ü": "ue", "ß": "ss",
    "Ä": "ae", "Ö": "oe", "Ü": "ue",
}

_FOOTNOTE_SUP = str.maketrans("¹²³⁴⁵⁶⁷⁸⁹⁰", "1234567890")


def normalize(text: str | None) -> str:
    """Kleinschreibung, Umlaute aufgeloest, Sonderzeichen entfernt."""
    if not text:
        return ""
    text = text.translate(_FOOTNOTE_SUP)
    for src, dst in _UMLAUT_MAP.items():
        text = text.replace(src, dst)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def similarity(a: str, b: str) -> float:
    """Aehnlichkeit 0..1 zweier Texte nach Normalisierung."""
    na, nb = normalize(a), normalize(b)
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    # Teilstring-Treffer bewusst hoch bewerten: Kopfzeilen sind oft abgekuerzt
    # oder umgebrochen ("Grenzwert\nfest" vs. "Grenzwert fest").
    if na in nb or nb in na:
        shorter, longer = (na, nb) if len(na) < len(nb) else (nb, na)
        if len(shorter) >= 4:
            return 0.85 + 0.1 * (len(shorter) / len(longer))
    return SequenceMatcher(None, na, nb).ratio()


def contains_keyword(haystack: str, needle: str) -> bool:
    return normalize(needle) in normalize(haystack)


_NUM_RE = re.compile(r"^\s*(\d+(?:[.,]\d+)?)\s*$")


def parse_count(value) -> float | None:
    """Wandelt einen Zellwert in eine Zahl.

    - Zahl oder Zahl-String -> float
    - Kreuz/Haken ("x", "X", "✓", "*") -> 1.0
    - leer / Freitext -> None
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return 1.0 if value else None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    if text in {"x", "X", "✓", "✔", "*", "•"}:
        return 1.0
    m = _NUM_RE.match(text.replace(" ", " "))
    if m:
        return float(m.group(1).replace(",", "."))
    return None


def is_freetext(value) -> bool:
    """True, wenn der Zellinhalt eine Anmerkung und kein Zaehlwert ist."""
    if value is None:
        return False
    text = str(value).strip()
    if not text:
        return False
    return parse_count(text) is None


def extract_footnote_markers(text: str | None) -> list[str]:
    """Findet Fussnotenverweise in einer Spaltenueberschrift ('Stellen 3)' -> ['3'])."""
    if not text:
        return []
    markers: list[str] = []
    for raw in re.findall(r"[¹²³⁴⁵⁶⁷⁸⁹⁰]", text):
        markers.append(raw.translate(_FOOTNOTE_SUP))
    for raw in re.findall(r"(?<![\d,.])(\d{1,2})\s*\)", text):
        markers.append(raw)
    seen: set[str] = set()
    out: list[str] = []
    for m in markers:
        if m not in seen:
            seen.add(m)
            out.append(m)
    return out


def normalize_address(address: str | None) -> str:
    """Normalisiert eine Spaltenadresse der neuen Fassung.

    '2.2.' / ' 2.2 ' / '2,2' -> '2.2'; '0.' -> '0'
    """
    if not address:
        return ""
    text = str(address).strip().replace(",", ".").replace(" ", "")
    parts = [p for p in text.split(".") if p != ""]
    if not parts or not all(p.isdigit() for p in parts):
        return ""
    return ".".join(str(int(p)) for p in parts)


def join_address(section: str | None, column_no: str | None) -> str:
    """Setzt die Adresse aus den Kopfzeilen 'Abschnitt' und 'Spalte' zusammen.

    ('1.1.', '3') -> '1.1.3'; ('0.', '1') -> '0.1'
    """
    sec = normalize_address(section)
    col = normalize_address(column_no)
    if sec and col:
        return f"{sec}.{col}"
    return sec or col
