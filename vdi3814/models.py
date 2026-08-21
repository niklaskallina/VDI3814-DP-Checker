"""Datenmodell der Extraktion (vor dem Schreiben in die Datenbank).

Bewusst als schlanke Dataclasses gehalten: die Vision-/OCR-Schicht erzeugt
diese Objekte, die UI bearbeitet sie, erst danach werden sie persistiert.
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from datetime import datetime
from enum import Enum
from typing import Any, Iterable


class RowKind(str, Enum):
    """Art einer Tabellenzeile - entscheidet, ob sie gezaehlt wird."""

    DATEN = "daten"              # echter Datenpunkt -> wird gezaehlt
    LEER = "leer"                # keine Beschriftung, keine Werte
    SUMME = "summe"              # Summen-/Zwischensummenzeile
    UEBERTRAG = "uebertrag"      # Uebertrag von/auf ein anderes Blatt
    FUSSBEREICH = "fussbereich"  # Zeile unterhalb der Tabelle
    KOPF = "kopf"                # wiederholte Kopfzeile


class PageKind(str, Enum):
    """Ergebnis der Seitenklassifikation."""

    FUNKTIONSLISTE = "funktionsliste"
    SUMMENBLATT = "summenblatt"     # nur Summen je Anlage, wird nicht gezaehlt
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
    bbox: tuple[float, float, float, float] | None = None   # Fundstelle auf der Seite

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
    bbox: tuple[float, float, float, float] | None = None    # Fundstelle der Zeile
    kind: RowKind = RowKind.DATEN
    exclusion_reason: str = ""
    confidence: float = 0.0
    warnings: list[str] = field(default_factory=list)

    @property
    def is_sum_row(self) -> bool:
        """Summen- und Uebertragszeilen dienen nur der Kontrolle."""
        return self.kind in (RowKind.SUMME, RowKind.UEBERTRAG)

    @property
    def has_values(self) -> bool:
        return any(cell.count is not None for cell in self.cells)

    @property
    def is_countable(self) -> bool:
        """Nur echte Datenpunkte mit mindestens einem Wert werden gezaehlt."""
        return self.kind is RowKind.DATEN and self.has_values

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
    """Abgleich einer Spaltensumme gegen die Summenzeile - immer je Seite.

    Ein dokumentweiter Vergleich waere falsch: Plansaetze enthalten oft ein
    eigenes Summenblatt mit Summen je Anlage, die einen anderen Umfang haben
    als die Detailseiten (teils sogar Anlagen aus anderen Dateien).
    """

    column_index: int
    reported: float | None
    computed: float
    matches: bool
    page_index: int = 0

    @property
    def difference(self) -> float:
        return (self.reported or 0.0) - self.computed


@dataclass
class PageResult:
    page_index: int
    classification: PageClassification
    image_path: str = ""
    width: float = 0.0        # Seitenbreite in Quellkoordinaten (fuer den Nachweis)
    height: float = 0.0
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
        """Zeilen, die als Datenpunkte gelten (ohne Summen, Uebertraege, Leerzeilen)."""
        return [r for r in self.rows if r.kind is RowKind.DATEN]

    def counted_rows(self) -> list[DataPointRow]:
        """Zeilen, die tatsaechlich in die Summen eingehen."""
        return [r for r in self.rows if r.is_countable]

    def excluded_rows(self) -> list[DataPointRow]:
        """Nicht gezaehlte Zeilen inkl. Begruendung - fuer den Pruefbericht."""
        return [r for r in self.rows if r.kind is not RowKind.DATEN]

    def remove_data_rows(self, positions: Iterable[int]) -> int:
        """Entfernt Datenpunktzeilen anhand ihrer Position in data_rows().

        Genau diese Reihenfolge zeigt die Korrekturtabelle der Oberflaeche, so
        dass eine dort angekreuzte Zeile zuverlaessig die gemeinte Zeile
        trifft. Summen-, Uebertrags- und Leerzeilen bleiben unberuehrt; sie
        werden nicht angezeigt und duerfen sich beim Loeschen nicht
        verschieben. Rueckgabe: Anzahl der tatsaechlich entfernten Zeilen.
        """
        gemeint = {int(p) for p in positions}
        stellen = [i for i, row in enumerate(self.rows) if row.kind is RowKind.DATEN]
        entfernen = {stellen[p] for p in gemeint if 0 <= p < len(stellen)}
        if not entfernen:
            return 0
        self.rows = [row for i, row in enumerate(self.rows) if i not in entfernen]
        return len(entfernen)

    def column_by_index(self, index: int) -> ColumnHeader | None:
        for col in self.columns:
            if col.index == index:
                return col
        return None

    def totals(self) -> dict[int, float]:
        """Eigene Spaltensummen - ausschliesslich ueber echte Datenpunkte."""
        totals: dict[int, float] = {}
        for row in self.counted_rows():
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
