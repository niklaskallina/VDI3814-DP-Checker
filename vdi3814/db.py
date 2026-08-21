"""Lokale SQLite-Datenbank (SQLAlchemy 2.0).

Speichert alle extrahierten Rohdaten inkl. Quelldatei, VDI-Fassung und
Zeitstempel. Die Datenbank ist inkrementell: neue Listen koennen jederzeit
ergaenzt werden, bereits importierte Dateien werden ueber ihren SHA-256
erkannt und wahlweise ersetzt.
"""

from __future__ import annotations

import json
import logging
import shutil
from datetime import datetime
from pathlib import Path
from typing import Iterable

from sqlalchemy import (
    Boolean,
    event,
    DateTime,
    Float,
    ForeignKey,
    Integer,
    String,
    Text,
    create_engine,
    select,
)
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column, relationship

from .config import SETTINGS
from .models import (
    Cell,
    ColumnHeader,
    DataPointRow,
    DocumentMetadata,
    DocumentResult,
    Footnote,
    PageClassification,
    PageKind,
    PageResult,
    RowKind,
)


log = logging.getLogger(__name__)


class Base(DeclarativeBase):
    pass


def _bbox_to_text(bbox) -> str:
    return ",".join(f"{value:.2f}" for value in bbox) if bbox else ""


def bbox_from_text(text: str) -> tuple[float, float, float, float] | None:
    """Fundstelle aus der Datenbank zurueckwandeln."""
    if not text:
        return None
    try:
        werte = [float(part) for part in text.split(",")]
    except ValueError:
        return None
    return tuple(werte) if len(werte) == 4 else None


class Document(Base):
    __tablename__ = "documents"

    id: Mapped[int] = mapped_column(primary_key=True)
    file_name: Mapped[str] = mapped_column(String(255))
    source_path: Mapped[str] = mapped_column(Text, default="")
    file_hash: Mapped[str] = mapped_column(String(64), index=True, default="")
    imported_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.now)
    profile_id: Mapped[str] = mapped_column(String(64), default="")
    fassung: Mapped[str] = mapped_column(String(16), default="")
    backend: Mapped[str] = mapped_column(String(32), default="")
    warnings_json: Mapped[str] = mapped_column(Text, default="[]")
    stored_path: Mapped[str] = mapped_column(Text, default="")   # Kopie der Originaldatei

    projekt: Mapped[str] = mapped_column(Text, default="")
    auftraggeber: Mapped[str] = mapped_column(Text, default="")
    anlage: Mapped[str] = mapped_column(Text, default="")
    gewerk: Mapped[str] = mapped_column(Text, default="")
    planersteller: Mapped[str] = mapped_column(Text, default="")
    protokoll: Mapped[str] = mapped_column(Text, default="")
    planstand: Mapped[str] = mapped_column(Text, default="")
    blatt_nr: Mapped[str] = mapped_column(String(32), default="")
    blatt_von: Mapped[str] = mapped_column(String(32), default="")
    datum: Mapped[str] = mapped_column(String(64), default="")
    vdi_blatt: Mapped[str] = mapped_column(String(64), default="")
    informationsschwerpunkt: Mapped[str] = mapped_column(Text, default="")
    automationsschwerpunkt: Mapped[str] = mapped_column(Text, default="")

    columns: Mapped[list["Column"]] = relationship(back_populates="document", cascade="all, delete-orphan")
    rows: Mapped[list["Row"]] = relationship(back_populates="document", cascade="all, delete-orphan")
    pages: Mapped[list["Page"]] = relationship(back_populates="document", cascade="all, delete-orphan")
    footnotes: Mapped[list["FootnoteRow"]] = relationship(back_populates="document", cascade="all, delete-orphan")

    @property
    def warnings(self) -> list[str]:
        try:
            return json.loads(self.warnings_json)
        except json.JSONDecodeError:
            return []


class Page(Base):
    __tablename__ = "pages"

    id: Mapped[int] = mapped_column(primary_key=True)
    document_id: Mapped[int] = mapped_column(ForeignKey("documents.id", ondelete="CASCADE"), index=True)
    page_index: Mapped[int] = mapped_column(Integer)
    kind: Mapped[str] = mapped_column(String(32), default="")
    confidence: Mapped[float] = mapped_column(Float, default=0.0)
    reason: Mapped[str] = mapped_column(Text, default="")
    width: Mapped[float] = mapped_column(Float, default=0.0)
    height: Mapped[float] = mapped_column(Float, default=0.0)
    schwerpunkt: Mapped[str] = mapped_column(Text, default="")   # ASP/ISP der Seite

    document: Mapped[Document] = relationship(back_populates="pages")


class Column(Base):
    __tablename__ = "columns"

    id: Mapped[int] = mapped_column(primary_key=True)
    document_id: Mapped[int] = mapped_column(ForeignKey("documents.id", ondelete="CASCADE"), index=True)
    idx: Mapped[int] = mapped_column(Integer)
    column_key: Mapped[str] = mapped_column(String(32), default="", index=True)
    address: Mapped[str] = mapped_column(String(16), default="")
    group_name: Mapped[str] = mapped_column(Text, default="")
    subgroup: Mapped[str] = mapped_column(Text, default="")
    label: Mapped[str] = mapped_column(Text, default="")
    abbrev: Mapped[str] = mapped_column(String(32), default="")
    is_note: Mapped[bool] = mapped_column(Boolean, default=False)
    match_score: Mapped[float] = mapped_column(Float, default=0.0)
    footnote_markers: Mapped[str] = mapped_column(String(64), default="")

    document: Mapped[Document] = relationship(back_populates="columns")


class Row(Base):
    __tablename__ = "rows"

    id: Mapped[int] = mapped_column(primary_key=True)
    document_id: Mapped[int] = mapped_column(ForeignKey("documents.id", ondelete="CASCADE"), index=True)
    page_index: Mapped[int] = mapped_column(Integer, default=0)
    row_no: Mapped[str] = mapped_column(String(16), default="")
    bas: Mapped[str] = mapped_column(Text, default="")
    klartext: Mapped[str] = mapped_column(Text, default="")
    qualifier: Mapped[str] = mapped_column(Text, default="")
    remark: Mapped[str] = mapped_column(Text, default="")
    kind: Mapped[str] = mapped_column(String(16), default="daten", index=True)
    # Automations-/Informationsschwerpunkt der Zeile - Grundlage der
    # Auswertung "wie viele Datenpunkte entfallen auf welchen ASP/ISP"
    schwerpunkt: Mapped[str] = mapped_column(String(16), default="", index=True)
    schwerpunkt_text: Mapped[str] = mapped_column(Text, default="")
    exclusion_reason: Mapped[str] = mapped_column(Text, default="")
    bbox: Mapped[str] = mapped_column(String(64), default="")   # Fundstelle "x0,y0,x1,y1"
    confidence: Mapped[float] = mapped_column(Float, default=0.0)

    @property
    def is_countable(self) -> bool:
        return self.kind == "daten" and any(v.count is not None for v in self.values)

    document: Mapped[Document] = relationship(back_populates="rows")
    values: Mapped[list["Value"]] = relationship(back_populates="row", cascade="all, delete-orphan")


class Value(Base):
    __tablename__ = "values"

    id: Mapped[int] = mapped_column(primary_key=True)
    row_id: Mapped[int] = mapped_column(ForeignKey("rows.id", ondelete="CASCADE"), index=True)
    column_idx: Mapped[int] = mapped_column(Integer)
    column_key: Mapped[str] = mapped_column(String(32), default="", index=True)
    count: Mapped[float | None] = mapped_column(Float, nullable=True)
    raw_value: Mapped[str] = mapped_column(Text, default="")
    note: Mapped[str] = mapped_column(Text, default="")
    bbox: Mapped[str] = mapped_column(String(64), default="")   # Fundstelle "x0,y0,x1,y1"

    row: Mapped[Row] = relationship(back_populates="values")


class FootnoteRow(Base):
    __tablename__ = "footnotes"

    id: Mapped[int] = mapped_column(primary_key=True)
    document_id: Mapped[int] = mapped_column(ForeignKey("documents.id", ondelete="CASCADE"), index=True)
    marker: Mapped[str] = mapped_column(String(8), default="")
    text: Mapped[str] = mapped_column(Text, default="")
    referenced_columns: Mapped[str] = mapped_column(Text, default="")

    document: Mapped[Document] = relationship(back_populates="footnotes")


class UnitPrice(Base):
    """Einheitspreis je Funktionsspalte - unabhaengig von den Importen pflegbar."""

    __tablename__ = "unit_prices"

    column_key: Mapped[str] = mapped_column(String(32), primary_key=True)
    profile_id: Mapped[str] = mapped_column(String(64), default="")
    label: Mapped[str] = mapped_column(Text, default="")
    price: Mapped[float] = mapped_column(Float, default=0.0)
    currency: Mapped[str] = mapped_column(String(8), default="EUR")
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.now)


# --------------------------------------------------------------------------
# Verbindung
# --------------------------------------------------------------------------

def _migrate(engine) -> None:
    """Ergaenzt fehlende Spalten in einer aelteren Datenbank.

    So bleiben bereits importierte Listen erhalten, wenn das Datenmodell
    erweitert wird - ein Neuaufbau der Datenbank ist nie noetig.
    """
    from sqlalchemy import text

    with engine.begin() as connection:
        for table in Base.metadata.sorted_tables:
            vorhanden = {
                row[1] for row in connection.execute(text(f"PRAGMA table_info('{table.name}')"))
            }
            if not vorhanden:
                continue                      # Tabelle wird ohnehin neu angelegt
            for column in table.columns:
                if column.name in vorhanden:
                    continue
                typ = column.type.compile(engine.dialect)
                default = column.default.arg if getattr(column.default, "is_scalar", False) else None
                zusatz = ""
                if isinstance(default, str):
                    zusatz = f" DEFAULT '{default}'"
                elif isinstance(default, (int, float)):
                    zusatz = f" DEFAULT {default}"
                connection.execute(
                    text(f"ALTER TABLE {table.name} ADD COLUMN {column.name} {typ}{zusatz}")
                )

        # Alte Kennzeichnung der Summenzeilen uebernehmen
        rows_columns = {row[1] for row in connection.execute(text("PRAGMA table_info('rows')"))}
        if {"is_sum_row", "kind"} <= rows_columns:
            connection.execute(text(
                "UPDATE rows SET kind = 'summe' WHERE is_sum_row = 1 AND (kind IS NULL OR kind = '')"
            ))
            connection.execute(text(
                "UPDATE rows SET kind = 'daten' WHERE kind IS NULL OR kind = ''"
            ))


# Offene Verbindungen je Datenbankdatei. Unter Windows laesst sich eine Datei
# nicht loeschen, solange sie geoeffnet ist - deshalb werden die Verbindungen
# hier verwaltet und vor dem Loeschen eines Projekts geschlossen.
_ENGINES: dict[str, object] = {}


def dispose_engine(db_path: str | Path) -> None:
    """Schliesst alle Verbindungen zu einer Datenbankdatei."""
    schluessel = str(Path(db_path).resolve())
    engine = _ENGINES.pop(schluessel, None)
    if engine is not None:
        engine.dispose()


def dispose_all() -> None:
    for engine in list(_ENGINES.values()):
        engine.dispose()
    _ENGINES.clear()


def make_engine(db_path: str | Path | None = None):
    path = Path(db_path or SETTINGS.db_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    schluessel = str(path.resolve())
    vorhanden = _ENGINES.get(schluessel)
    if vorhanden is not None:
        return vorhanden
    engine = create_engine(f"sqlite:///{path}", future=True)

    @event.listens_for(engine, "connect")
    def _enable_foreign_keys(connection, _record):   # pragma: no cover - Treiberhook
        cursor = connection.cursor()
        cursor.execute("PRAGMA foreign_keys=ON")
        cursor.close()

    Base.metadata.create_all(engine)
    _migrate(engine)
    _ENGINES[schluessel] = engine
    return engine


# --------------------------------------------------------------------------
# Schreiben / Lesen
# --------------------------------------------------------------------------

def find_by_hash(session: Session, file_hash: str) -> Document | None:
    if not file_hash:
        return None
    return session.scalars(select(Document).where(Document.file_hash == file_hash)).first()


def save_document(session: Session, result: DocumentResult, replace: bool = True) -> Document:
    """Schreibt ein Extraktionsergebnis in die Datenbank.

    Ist die Datei (SHA-256) bereits vorhanden, wird der alte Datensatz
    ersetzt - so lassen sich Korrekturen erneut importieren, ohne Dubletten
    zu erzeugen.
    """
    existing = find_by_hash(session, result.file_hash)
    if existing is not None:
        if not replace:
            return existing
        # ueber das ORM loeschen, damit die Kaskade auf Spalten/Zeilen/Werte greift
        session.delete(existing)
        session.flush()

    document = Document(
        file_name=result.file_name,
        source_path=result.source_path,
        file_hash=result.file_hash,
        imported_at=result.imported_at,
        profile_id=result.profile_id,
        fassung=result.fassung,
        backend=result.backend,
        warnings_json=json.dumps(result.warnings, ensure_ascii=False),
        **{name: getattr(result.metadata, name) for name in vars(result.metadata)},
    )
    session.add(document)
    session.flush()

    key_by_index: dict[int, str] = {}
    for column in result.columns:
        key_by_index[column.index] = column.column_key or ""
        session.add(Column(
            document_id=document.id,
            idx=column.index,
            column_key=column.column_key or "",
            address=column.address,
            group_name=column.group,
            subgroup=column.subgroup,
            label=column.label,
            abbrev=column.abbrev,
            is_note=column.is_note_column,
            match_score=column.match_score,
            footnote_markers=",".join(column.footnote_markers),
        ))

    for row in result.rows:
        db_row = Row(
            document_id=document.id,
            page_index=row.page_index,
            row_no=row.row_no,
            bas=row.bas,
            klartext=row.klartext,
            qualifier=row.qualifier,
            remark=row.remark,
            schwerpunkt=row.schwerpunkt,
            schwerpunkt_text=row.schwerpunkt_text,
            kind=row.kind.value,
            exclusion_reason=row.exclusion_reason,
            bbox=_bbox_to_text(row.bbox),
            confidence=row.confidence,
        )
        session.add(db_row)
        session.flush()
        for cell in row.cells:
            session.add(Value(
                row_id=db_row.id,
                column_idx=cell.column_index,
                column_key=key_by_index.get(cell.column_index, ""),
                count=cell.count,
                raw_value=cell.raw_value,
                note=cell.note,
                bbox=_bbox_to_text(cell.bbox),
            ))

    for page in result.pages:
        session.add(Page(
            document_id=document.id,
            page_index=page.page_index,
            kind=page.classification.kind.value,
            confidence=page.classification.confidence,
            reason=page.classification.reason,
            width=page.width,
            height=page.height,
            schwerpunkt=page.schwerpunkt,
        ))

    for footnote in result.footnotes:
        session.add(FootnoteRow(
            document_id=document.id,
            marker=footnote.marker,
            text=footnote.text,
            referenced_columns=",".join(footnote.referenced_columns),
        ))

    _store_source(session, document, result)
    session.commit()
    return document


def source_directory(session: Session) -> Path | None:
    """Ablageordner fuer Originaldateien neben der Datenbankdatei."""
    bind = session.get_bind()
    database = getattr(getattr(bind, "url", None), "database", None)
    if not database:
        return None
    return Path(database).parent / "quellen"


def _store_source(session: Session, document: Document, result: DocumentResult) -> None:
    """Legt eine Kopie der Originaldatei im Projekt ab.

    Nur so laesst sich spaeter im Programm zeigen, an welcher Stelle im PDF ein
    Wert gefunden wurde - auch wenn die Ursprungsdatei verschoben wurde.
    """
    quelle = Path(result.source_path)
    ziel_ordner = source_directory(session)
    if ziel_ordner is None or not quelle.is_file():
        return
    ziel_ordner.mkdir(parents=True, exist_ok=True)
    ziel = ziel_ordner / f"{result.file_hash[:16]}{quelle.suffix.lower()}"
    if not ziel.exists():
        try:
            shutil.copy2(quelle, ziel)
        except OSError as exc:              # pragma: no cover - Rechteproblem
            log.warning("Originaldatei konnte nicht abgelegt werden: %s", exc)
            return
    document.stored_path = str(ziel)


def list_documents(session: Session) -> list[Document]:
    return list(session.scalars(select(Document).order_by(Document.imported_at.desc())))


def delete_document(session: Session, document_id: int) -> None:
    document = session.get(Document, document_id)
    if document is not None:
        session.delete(document)
        session.commit()


def load_document(session: Session, document_id: int) -> DocumentResult:
    """Laedt einen Datensatz zurueck in das Arbeitsmodell (fuer Re-Export/Korrektur)."""
    document = session.get(Document, document_id)
    if document is None:
        raise KeyError(f"Dokument {document_id} nicht gefunden")

    metadata = DocumentMetadata(**{name: getattr(document, name) for name in vars(DocumentMetadata())})
    result = DocumentResult(
        source_path=document.source_path,
        file_name=document.file_name,
        file_hash=document.file_hash,
        imported_at=document.imported_at,
        profile_id=document.profile_id,
        fassung=document.fassung,
        metadata=metadata,
        backend=document.backend,
        warnings=document.warnings,
    )
    result.columns = [
        ColumnHeader(
            index=column.idx,
            group=column.group_name,
            subgroup=column.subgroup,
            label=column.label,
            abbrev=column.abbrev,
            section="", column_no="",
            footnote_markers=[m for m in column.footnote_markers.split(",") if m],
            column_key=column.column_key or None,
            match_score=column.match_score,
            is_note_column=column.is_note,
        )
        for column in sorted(document.columns, key=lambda c: c.idx)
    ]
    # Adresse wird beim Zurueckladen aus dem Profil abgeleitet, nicht gespeichert
    result.rows = [
        DataPointRow(
            row_no=row.row_no,
            bas=row.bas,
            klartext=row.klartext,
            qualifier=row.qualifier,
            remark=row.remark,
            page_index=row.page_index,
            schwerpunkt=row.schwerpunkt or "",
            schwerpunkt_text=row.schwerpunkt_text or "",
            kind=RowKind(row.kind) if row.kind else RowKind.DATEN,
            exclusion_reason=row.exclusion_reason,
            confidence=row.confidence,
            bbox=bbox_from_text(row.bbox),
            cells=[Cell(column_index=v.column_idx, raw_value=v.raw_value, count=v.count,
                        note=v.note, bbox=bbox_from_text(v.bbox))
                   for v in row.values],
        )
        for row in document.rows
    ]
    result.pages = [
        PageResult(page_index=page.page_index,
                   classification=PageClassification(PageKind(page.kind) if page.kind else PageKind.SONSTIGES,
                                                     page.confidence, page.reason),
                   width=page.width, height=page.height,
                   schwerpunkt=page.schwerpunkt or "")
        for page in document.pages
    ]
    result.footnotes = [
        Footnote(marker=f.marker, text=f.text,
                 referenced_columns=[c for c in f.referenced_columns.split(",") if c])
        for f in document.footnotes
    ]
    return result


# --------------------------------------------------------------------------
# Einheitspreise
# --------------------------------------------------------------------------

def get_unit_prices(session: Session) -> dict[str, float]:
    return {price.column_key: price.price for price in session.scalars(select(UnitPrice))}


def set_unit_price(session: Session, column_key: str, price: float, label: str = "",
                   profile_id: str = "", currency: str | None = None) -> None:
    entry = session.get(UnitPrice, column_key)
    if entry is None:
        entry = UnitPrice(column_key=column_key, profile_id=profile_id, label=label)
        session.add(entry)
    entry.price = float(price)
    entry.label = label or entry.label
    entry.profile_id = profile_id or entry.profile_id
    entry.currency = currency or SETTINGS.currency
    entry.updated_at = datetime.now()
    session.commit()


def set_unit_prices(session: Session, prices: dict[str, float], profile_id: str = "") -> None:
    for key, value in prices.items():
        set_unit_price(session, key, value, profile_id=profile_id)
