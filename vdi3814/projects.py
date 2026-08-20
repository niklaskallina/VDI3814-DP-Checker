"""Projektverwaltung.

Jedes Projekt ist eine eigene Datenbankdatei mit einem eigenen Ablageordner
fuer die importierten Originaldateien:

    data/projekte/<projekt>/vdi3814.sqlite3
    data/projekte/<projekt>/quellen/<pruefsumme>.pdf

Damit lassen sich Auswertungen sauber trennen ("nicht alles auf einen Haufen"),
ein Projekt vollstaendig loeschen und Ergebnisse als ein Ordner weitergeben.
"""

from __future__ import annotations

import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from .config import SETTINGS

DEFAULT_PROJECT = "Standard"
DB_FILENAME = "vdi3814.sqlite3"
SOURCE_DIRNAME = "quellen"


def projects_root() -> Path:
    return SETTINGS.data_dir / "projekte"


def slug(name: str) -> str:
    """Dateisystemtauglicher Ordnername aus einem Projektnamen."""
    cleaned = re.sub(r"[^\w\s.-]", "", name, flags=re.UNICODE).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned[:80] or "Projekt"


@dataclass
class Project:
    name: str
    path: Path

    @property
    def db_path(self) -> Path:
        return self.path / DB_FILENAME

    @property
    def sources_dir(self) -> Path:
        return self.path / SOURCE_DIRNAME

    @property
    def changed_at(self) -> datetime | None:
        if self.db_path.exists():
            return datetime.fromtimestamp(self.db_path.stat().st_mtime)
        return None

    @property
    def size_mb(self) -> float:
        total = sum(f.stat().st_size for f in self.path.rglob("*") if f.is_file())
        return round(total / (1024 * 1024), 2)


def get(name: str) -> Project:
    return Project(name=name, path=projects_root() / slug(name))


def create(name: str) -> Project:
    """Legt ein Projekt an (idempotent) und gibt es zurueck."""
    project = get(name)
    project.path.mkdir(parents=True, exist_ok=True)
    project.sources_dir.mkdir(parents=True, exist_ok=True)
    (project.path / "projekt.txt").write_text(name, encoding="utf-8")
    return project


def list_projects() -> list[Project]:
    root = projects_root()
    if not root.exists():
        return [create(DEFAULT_PROJECT)]
    projects = []
    for entry in sorted(root.iterdir()):
        if not entry.is_dir():
            continue
        marker = entry / "projekt.txt"
        name = marker.read_text(encoding="utf-8").strip() if marker.exists() else entry.name
        projects.append(Project(name=name, path=entry))
    return projects or [create(DEFAULT_PROJECT)]


def delete(name: str) -> bool:
    """Loescht ein Projekt samt Datenbank und abgelegten Originaldateien."""
    project = get(name)
    if not project.path.exists():
        return False
    shutil.rmtree(project.path)
    return True


def clear(name: str) -> bool:
    """Leert ein Projekt: Datenbank und Originaldateien weg, Projekt bleibt."""
    project = get(name)
    if not project.path.exists():
        return False
    project.db_path.unlink(missing_ok=True)
    if project.sources_dir.exists():
        shutil.rmtree(project.sources_dir)
    project.sources_dir.mkdir(parents=True, exist_ok=True)
    return True


def resolve(name: str | None) -> Project:
    """Projekt nach Namen holen und dabei anlegen, falls noetig."""
    return create(name or DEFAULT_PROJECT)
