"""Laden und Anwenden der Spalten-Profile (alte/neue VDI-3814-Fassung).

Die Profile liegen als YAML in vdi3814/profiles/ und sind frei erweiterbar.
Der Erkennungskern kennt keine festen Spalten - er kennt nur Profile.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path
from typing import Iterable

import yaml

from .config import PROFILE_DIR
from .models import ColumnHeader
from .textutil import normalize, normalize_address, similarity

# Ab dieser Aehnlichkeit gilt eine gelesene Ueberschrift als Profil-Spalte.
MATCH_THRESHOLD = 0.62


@dataclass(frozen=True)
class ProfileColumn:
    key: str
    group: str
    subgroup: str
    label: str
    abbrev: str
    aggregate_group: str
    aliases: tuple[str, ...] = ()
    is_note: bool = False
    address: str = ""      # normative Spaltenadresse, z. B. "2.2.10" (neu) / "6.13" (alt)
    footnotes: tuple[str, ...] = ()   # Fussnotenverweise laut Vorlage

    def candidates(self) -> tuple[str, ...]:
        return (self.label, self.abbrev, *self.aliases)

    def display(self) -> str:
        parts = [p for p in (self.group, self.subgroup, self.label) if p]
        text = " / ".join(parts)
        return f"{self.address} {text}".strip() if self.address else text


@dataclass(frozen=True)
class RowField:
    key: str
    label: str
    aliases: tuple[str, ...] = ()


@dataclass
class Profile:
    id: str
    name: str
    fassung: str
    priority: int = 0
    min_score: float = 3.0
    keywords: tuple[tuple[str, float], ...] = ()
    negative_keywords: tuple[tuple[str, float], ...] = ()
    columns: tuple[ProfileColumn, ...] = ()
    note_columns: tuple[ProfileColumn, ...] = ()
    row_fields: tuple[RowField, ...] = ()
    footnotes: tuple[tuple[str, str], ...] = ()   # (Marker, Text) aus der Vorlage
    source_file: str = ""

    # ---------------- Erkennung ----------------

    def score(self, texts: Iterable[str]) -> float:
        """Punktzahl fuer 'passt dieses Profil zum Dokument?'."""
        blob = normalize(" ".join(t for t in texts if t))
        score = 0.0
        for text, weight in self.keywords:
            if normalize(text) in blob:
                score += weight
        for text, weight in self.negative_keywords:
            if normalize(text) in blob:
                score -= weight
        return score

    # ---------------- Zugriff ----------------

    def all_columns(self) -> tuple[ProfileColumn, ...]:
        return self.columns + self.note_columns

    def column(self, key: str) -> ProfileColumn | None:
        for col in self.all_columns():
            if col.key == key:
                return col
        return None

    def column_by_address(self, address: str) -> ProfileColumn | None:
        wanted = normalize_address(address)
        if not wanted:
            return None
        for col in self.all_columns():
            if col.address and normalize_address(col.address) == wanted:
                return col
        return None


def _as_pairs(items) -> tuple[tuple[str, float], ...]:
    out = []
    for item in items or []:
        if isinstance(item, dict):
            out.append((str(item.get("text", "")), float(item.get("weight", 1.0))))
        else:
            out.append((str(item), 1.0))
    return tuple(out)


def _as_column(raw: dict, is_note: bool) -> ProfileColumn:
    return ProfileColumn(
        key=str(raw["key"]),
        group=str(raw.get("group", "")),
        subgroup=str(raw.get("subgroup", "")),
        label=str(raw.get("label", raw["key"])),
        abbrev=str(raw.get("abbrev", "")),
        aggregate_group=str(raw.get("aggregate_group", "SONSTIGE")),
        aliases=tuple(str(a) for a in raw.get("aliases", []) or []),
        is_note=is_note,
        address=str(raw.get("address", "")),
        footnotes=tuple(str(f) for f in raw.get("footnotes", []) or []),
    )


def load_profile(path: Path) -> Profile:
    raw = yaml.safe_load(path.read_text(encoding="utf-8"))
    detection = raw.get("detection", {}) or {}
    return Profile(
        id=str(raw["id"]),
        name=str(raw.get("name", raw["id"])),
        fassung=str(raw.get("fassung", "unbekannt")),
        priority=int(raw.get("priority", 0)),
        min_score=float(detection.get("min_score", 3.0)),
        keywords=_as_pairs(detection.get("keywords")),
        negative_keywords=_as_pairs(detection.get("negative_keywords")),
        columns=tuple(_as_column(c, False) for c in raw.get("columns", []) or []),
        note_columns=tuple(_as_column(c, True) for c in raw.get("note_columns", []) or []),
        row_fields=tuple(
            RowField(str(f["key"]), str(f.get("label", f["key"])), tuple(f.get("aliases", []) or []))
            for f in raw.get("row_fields", []) or []
        ),
        footnotes=tuple(
            (str(f.get("marker", "")), str(f.get("text", "")))
            for f in raw.get("footnotes", []) or []
        ),
        source_file=str(path),
    )


@lru_cache(maxsize=1)
def load_profiles(profile_dir: str | None = None) -> tuple[Profile, ...]:
    directory = Path(profile_dir) if profile_dir else PROFILE_DIR
    profiles = [load_profile(p) for p in sorted(directory.glob("*.yaml"))]
    return tuple(sorted(profiles, key=lambda p: -p.priority))


def get_profile(profile_id: str) -> Profile | None:
    for profile in load_profiles():
        if profile.id == profile_id:
            return profile
    return None


def detect_profile(texts: Iterable[str]) -> tuple[Profile | None, float, dict[str, float]]:
    """Bestimmt die VDI-Fassung anhand der gelesenen Kopf-/Seitentexte.

    Rueckgabe: (Profil oder None, Score des Gewinners, alle Scores).
    """
    texts = list(texts)
    scores = {p.id: p.score(texts) for p in load_profiles()}
    if not scores:
        return None, 0.0, {}
    best_id = max(scores, key=lambda k: scores[k])
    best = get_profile(best_id)
    if best is None or scores[best_id] < best.min_score:
        return None, scores[best_id], scores
    return best, scores[best_id], scores


def match_columns(headers: list[ColumnHeader], profile: Profile) -> list[ColumnHeader]:
    """Ordnet gelesene Spaltenueberschriften den Profil-Spalten zu.

    Greedy nach bester Aehnlichkeit; jede Profil-Spalte wird hoechstens einmal
    vergeben. Nicht zuordenbare Spalten behalten column_key=None und werden
    spaeter als "unbekannte Spalte" gefuehrt (Daten gehen nicht verloren).
    """
    pairs: list[tuple[float, int, str]] = []
    for header in headers:
        # 1) Adressbasierte Zuordnung (neue Fassung) - schlaegt jeden Textvergleich.
        column = profile.column_by_address(header.address)
        if column is not None:
            pairs.append((1.0 + 0.5, header.index, column.key))
            continue
        header_texts = [t for t in (header.label, header.abbrev) if t]
        combined = " ".join([header.group, header.subgroup, header.label]).strip()
        for column in profile.all_columns():
            best = 0.0
            for cand in column.candidates():
                if not cand:
                    continue
                for text in header_texts or [combined]:
                    best = max(best, similarity(text, cand))
                best = max(best, similarity(combined, cand) * 0.95)
            # Gleiche Gruppe/Untergruppe erhoeht die Trefferwahrscheinlichkeit.
            if column.group and similarity(header.group, column.group) > 0.7:
                best = min(1.0, best + 0.08)
            if column.subgroup and similarity(header.subgroup, column.subgroup) > 0.7:
                best = min(1.0, best + 0.05)
            if best >= MATCH_THRESHOLD:
                pairs.append((best, header.index, column.key))

    pairs.sort(key=lambda p: -p[0])
    used_headers: set[int] = set()
    used_columns: set[str] = set()
    assignment: dict[int, tuple[str, float]] = {}
    for score, header_index, column_key in pairs:
        if header_index in used_headers or column_key in used_columns:
            continue
        used_headers.add(header_index)
        used_columns.add(column_key)
        assignment[header_index] = (column_key, score)

    for header in headers:
        if header.index in assignment:
            key, score = assignment[header.index]
            column = profile.column(key)
            header.column_key = key
            header.match_score = round(min(1.0, score), 3)
            header.is_note_column = bool(column and column.is_note)
        else:
            header.column_key = None
            header.match_score = 0.0
    return headers


def assign_footnotes(headers: list[ColumnHeader], footnotes: list) -> None:
    """Traegt in jede Fussnote ein, auf welche Spalten sie sich bezieht."""
    by_marker: dict[str, list[str]] = {}
    for header in headers:
        for marker in header.footnote_markers:
            by_marker.setdefault(marker, []).append(header.column_key or header.display())
    for footnote in footnotes:
        footnote.referenced_columns = by_marker.get(footnote.marker, [])
