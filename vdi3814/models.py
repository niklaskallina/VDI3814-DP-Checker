"""Datenmodell der Extraktion (vor dem Schreiben in die Datenbank).

Bewusst als schlanke Dataclasses gehalten: die Vision-/OCR-Schicht erzeugt
diese Objekte, die UI bearbeitet sie, erst danach werden sie persistiert.
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from datetime import datetime
from enum import Enum
from typing import Any


class PageKind(str, Enum):
    """Ergebnis der Seitenklassifikation."""

    FUNKTIONSLISTE = "funktionsliste"
    SCHEMA = "schema"
    SONSTIGES = "sonstiges"
    FEHLER = "fehler"


@dataclass
class PageClassification:
    kind: PageKind = PageKind.SONSTIGES
    confidence: float = 0.0
    reason: str = ""

    @property
    def is_table(self) -> bool:
        return self.kind is PageKind.FUNKTIONSLISTE


@dataclass
class ColumnHeader:
    """Eine im Dokument tatsaechlich gelesene Spaltenueberschrift.

    In der neuen Fassung traegt der Kopf zusaetzlich zwei Nummernzeilen
    ("Abschnitt" und "Spalte"), aus denen sich die normative Spaltenadresse
    ergibt (z. B. Abschnitt "2.2" + Spalte "10" -> Adresse "2.2.10").
    Diese Adresse ist der zuverlaessigste Anker fuer das Spalten-Mapping.
    """

    index: int
    group: str = ""
    subgroup: str = ""
    label: str = ""
    abbrev: str = ""
    section: str = ""      # Kopfzeile "Abschnitt"
    column_no: str = ""    # Kopfzeile "Spalte"
    x_center: float | None = None   # nur bei geometrischer Extraktion (PDF/Excel)
    footnote_markers: list[str] = field(default_factory=list)

    # Ergebnis des Mappings gegen das Profil (profiles_loader.match_columns)
    column_key: str | None = None
    match_score: float = 0.0
    is_note_column: bool = False

    @property
    def address(self) -> str:
        from .textutil import join_address

        return join_address(self.section, self.column_no)

    def display(self) -> str:
        parts = [p for p in (self.group, self.subgroup, self.label) if p]
        text = " / ".join(parts) or self.abbrev or f"Spalte {self.index}"
        return f"{self.address} {text}".strip() if self.address else text


@dataclass
class Footnote:
    marker: str
    text: str
    referenced_columns: list[str] = field(default_factory=list)


@dataclass
class Cell:
    column_index: int
    raw_value: str = ""
    count: float | None = None
    note: str = ""

    @property
    def is_note(self) -> bool:
        return bool(self.note) and self.count is None


@dataclass
class DataPointRow:
    """Eine Datenpunktzeile der Funktionsliste."""

    row_no: str = ""
    bas: str = ""
    klartext: str = ""
    qualifier: str = ""
    remark: str = ""
    cells: list[Cell] = field(default_factory=list)
    page_index: int = 0
    is_sum_row: bool = False
    confidence: float = 0.0
    warnings: list[str] = field(default_factory=list)

    def value(self, column_index: int) -> float | None:
        for cell in self.cells:
            if cell.column_index == column_index:
                return cell.count
        return None

    def label(self) -> str:
        return self.klartext or self.bas or f"Zeile {self.row_no}"


@dataclass
class DocumentMetadata:
    """Kopf-/Fussbereich des Dokuments."""

    projekt: str = ""
    auftraggeber: str = ""
    anlage: str = ""
    gewerk: str = ""
    planersteller: str = ""
    protokoll: str = ""
    planstand: str = ""
    blatt_nr: str = ""
    blatt_von: str = ""
    datum: str = ""
    vdi_blatt: str = ""
    informationsschwerpunkt: str = ""

    def merge(self, other: "DocumentMetadata") -> None:
        """Uebernimmt fehlende Felder aus einem weiteren Fund (z. B. Folgeseite)."""
        for key, value in asdict(other).items():
            if value and not getattr(self, key):
                setattr(self, key, value)


@dataclass
class SumCheck:
    column_index: int
    reported: float | None
    computed: float
    matches: bool


@dataclass
class PageResult:
    page_index: int
    classification: PageClassification
    image_path: str = ""
    error: str = ""
    rows: list[DataPointRow] = field(default_factory=list)

    @property
    def skipped(self) -> bool:
        return not self.classification.is_table


@dataclass
class DocumentResult:
    """Vollstaendiges Extraktionsergebnis einer Quelldatei."""

    source_path: str
    file_name: str
    file_hash: str = ""
    imported_at: datetime = field(default_factory=datetime.now)
    profile_id: str = ""
    fassung: str = ""
    profile_score: float = 0.0
    metadata: DocumentMetadata = field(default_factory=DocumentMetadata)
    columns: list[ColumnHeader] = field(default_factory=list)
    footnotes: list[Footnote] = field(default_factory=list)
    rows: list[DataPointRow] = field(default_factory=list)
    pages: list[PageResult] = field(default_factory=list)
    sum_checks: list[SumCheck] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    backend: str = ""

    # ---------------- Auswertung ----------------

    def data_rows(self) -> list[DataPointRow]:
        return [r for r in self.rows if not r.is_sum_row]

    def column_by_index(self, index: int) -> ColumnHeader | None:
        for col in self.columns:
            if col.index == index:
                return col
        return None

    def totals(self) -> dict[int, float]:
        """Eigene Spaltensummen ueber alle Datenzeilen (ohne Summenzeile)."""
        totals: dict[int, float] = {}
        for row in self.data_rows():
            for cell in row.cells:
                if cell.count is not None:
                    totals[cell.column_index] = totals.get(cell.column_index, 0.0) + cell.count
        return totals

    def grand_total(self) -> float:
        return float(sum(self.totals().values()))

    def skipped_pages(self) -> list[PageResult]:
        return [p for p in self.pages if p.skipped]

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["imported_at"] = self.imported_at.isoformat()
        return data
