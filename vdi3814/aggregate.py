"""Aggregation und Drilldown ueber alle importierten Funktionslisten."""

from __future__ import annotations

import pandas as pd
from sqlalchemy import select
from sqlalchemy.orm import Session

from . import db
from .profiles_loader import get_profile, load_profiles

RAW_COLUMNS = [
    "dokument_id", "datei", "fassung", "profil", "projekt", "auftraggeber", "anlage",
    "gewerk", "planersteller", "blatt", "seite", "zeile_nr", "datenpunkt", "bas",
    "qualifikation", "bemerkung", "spalte_adresse", "spalte_key", "gruppe",
    "untergruppe", "spalte", "wert", "rohwert", "notiz", "importiert_am",
]


SCHLUESSEL_SPALTE = "Schlüssel"


def dokument_schluessel(document: db.Document) -> str:
    """Eindeutiger Schluessel einer importierten Liste.

    Dateinamen sind nicht eindeutig - dieselbe Liste kann aus zwei Ordnern
    stammen. Ueber diesen Schluessel verweist die "Uebersicht" im Excel-Export
    auf genau die Zeilen der "GA-Funktionsliste", die zu dieser Liste gehoeren.
    """
    return f"D{document.id}"


def _column_address(column_key: str) -> str:
    for profile in load_profiles():
        column = profile.column(column_key)
        if column is not None:
            return column.address
    return ""


def raw_dataframe(session: Session) -> pd.DataFrame:
    """Ein Datensatz je Zelle - Basis fuer alle Auswertungen und den Detailexport."""
    records: list[dict] = []
    for document in session.scalars(select(db.Document)):
        columns = {c.idx: c for c in document.columns}
        for row in document.rows:
            # Nur echte Datenpunkte: keine Summen-, Uebertrags-, Leer-,
            # Kopf- oder Fussbereichszeilen.
            if row.kind != "daten":
                continue
            for value in row.values:
                column = columns.get(value.column_idx)
                if column is None or column.is_note:
                    continue
                records.append({
                    "dokument_id": document.id,
                    "datei": document.file_name,
                    "fassung": document.fassung,
                    "profil": document.profile_id,
                    "projekt": document.projekt or document.auftraggeber,
                    "auftraggeber": document.auftraggeber,
                    "anlage": document.anlage,
                    "gewerk": document.gewerk,
                    "planersteller": document.planersteller,
                    "blatt": document.blatt_nr,
                    "seite": row.page_index + 1,
                    "zeile_nr": row.row_no,
                    "datenpunkt": row.klartext,
                    "bas": row.bas,
                    "qualifikation": row.qualifier,
                    "bemerkung": row.remark,
                    "spalte_adresse": column.address or _column_address(column.column_key),
                    "spalte_key": column.column_key,
                    "gruppe": column.group_name,
                    "untergruppe": column.subgroup,
                    "spalte": column.label,
                    "wert": value.count if value.count is not None else 0.0,
                    "rohwert": value.raw_value,
                    "notiz": value.note,
                    "importiert_am": document.imported_at,
                })
    frame = pd.DataFrame(records, columns=RAW_COLUMNS)
    return frame


def _profile_order(profile_id: str) -> dict[str, int]:
    profile = get_profile(profile_id)
    if profile is None:
        return {}
    return {column.key: index for index, column in enumerate(profile.columns)}


def column_summary(frame: pd.DataFrame) -> pd.DataFrame:
    """Summen je Funktionsspalte ueber alle importierten Listen."""
    if frame.empty:
        return pd.DataFrame(columns=["profil", "spalte_adresse", "spalte_key", "gruppe",
                                     "untergruppe", "spalte", "menge", "dokumente", "datenpunkte"])
    grouped = (
        frame.groupby(["profil", "spalte_adresse", "spalte_key", "gruppe", "untergruppe", "spalte"],
                      dropna=False)
        .agg(menge=("wert", "sum"),
             dokumente=("dokument_id", "nunique"),
             datenpunkte=("datenpunkt", "count"))
        .reset_index()
    )
    orders = {profil: _profile_order(profil) for profil in grouped["profil"].unique()}
    grouped["_sort"] = [
        orders.get(profil, {}).get(key, 9999)
        for profil, key in zip(grouped["profil"], grouped["spalte_key"])
    ]
    grouped = grouped.sort_values(["profil", "_sort", "spalte_adresse"]).drop(columns="_sort")
    return grouped.reset_index(drop=True)


def group_summary(frame: pd.DataFrame) -> pd.DataFrame:
    """Summen je Spaltengruppe (Ein-/Ausgabe, Verarbeitung, ...)."""
    if frame.empty:
        return pd.DataFrame(columns=["fassung", "gruppe", "menge"])
    return (
        frame.groupby(["fassung", "gruppe"], dropna=False)["wert"].sum()
        .reset_index(name="menge").sort_values(["fassung", "gruppe"]).reset_index(drop=True)
    )


def by_dimension(frame: pd.DataFrame, dimension: str) -> pd.DataFrame:
    """Summen je Projekt / Anlage / Gewerk / Datei."""
    if frame.empty or dimension not in frame.columns:
        return pd.DataFrame(columns=[dimension, "menge", "datenpunkte", "dokumente"])
    return (
        frame.groupby(dimension, dropna=False)
        .agg(menge=("wert", "sum"),
             datenpunkte=("datenpunkt", "count"),
             dokumente=("dokument_id", "nunique"))
        .reset_index().sort_values("menge", ascending=False).reset_index(drop=True)
    )


def pivot_projects(frame: pd.DataFrame) -> pd.DataFrame:
    """Kreuztabelle Projekt/Anlage x Spaltengruppe."""
    if frame.empty:
        return pd.DataFrame()
    return pd.pivot_table(frame, index=["projekt", "anlage"], columns="gruppe",
                          values="wert", aggfunc="sum", fill_value=0, margins=True,
                          margins_name="Gesamt").reset_index()


def documents_frame(session: Session) -> pd.DataFrame:
    """Uebersicht der importierten Dateien inkl. uebersprungener Seiten."""
    records = []
    for document in session.scalars(select(db.Document)):
        pages = document.pages
        skipped = [p for p in pages if p.kind not in ("funktionsliste",)]
        records.append({
            "dokument_id": document.id,
            "datei": document.file_name,
            "fassung": document.fassung,
            "profil": document.profile_id,
            "verfahren": document.backend,
            "projekt": document.projekt,
            "auftraggeber": document.auftraggeber,
            "anlage": document.anlage,
            "gewerk": document.gewerk,
            "planersteller": document.planersteller,
            "protokoll": document.protokoll,
            "planstand": document.planstand,
            "blatt_nr": document.blatt_nr,
            "blatt_von": document.blatt_von,
            "datum": document.datum,
            "seiten": len(pages),
            "seiten_ausgewertet": len(pages) - len(skipped),
            "seiten_uebersprungen": "; ".join(
                f"S.{p.page_index + 1} ({p.kind})" for p in sorted(skipped, key=lambda p: p.page_index)
            ),
            "datenpunkte": len([r for r in document.rows if r.is_countable]),
            "summe_funktionen": sum(v.count or 0.0 for r in document.rows
                                    if r.kind == "daten" for v in r.values),
            "zeilen_ausgeschlossen": "; ".join(
                sorted({f"{r.kind}: {len([x for x in document.rows if x.kind == r.kind])}"
                        for r in document.rows if r.kind != "daten"})
            ),
            "hinweise": " | ".join(document.warnings),
            "importiert_am": document.imported_at,
            "quelle": document.source_path,
        })
    return pd.DataFrame(records)


def footnotes_frame(session: Session) -> pd.DataFrame:
    records = []
    for document in session.scalars(select(db.Document)):
        for footnote in document.footnotes:
            records.append({
                "datei": document.file_name,
                "fassung": document.fassung,
                "marker": footnote.marker,
                "text": footnote.text,
                "bezieht_sich_auf": footnote.referenced_columns,
            })
    return pd.DataFrame(records, columns=["datei", "fassung", "marker", "text", "bezieht_sich_auf"])


def vdi_layout_frame(session: Session, profile_id: str) -> pd.DataFrame:
    """Datenpunkte im Original-Layout der VDI-Funktionsliste.

    Eine Zeile je Datenpunkt, eine Spalte je Funktionsspalte des Profils -
    also genau der Aufbau der Papierliste, nur ueber alle importierten Dateien
    hinweg. Grundlage fuer das Blatt "GA-Funktionsliste" im Excel-Export.
    """
    profile = get_profile(profile_id)
    if profile is None:
        return pd.DataFrame()

    column_titles = [column.label for column in profile.columns]
    kopf = ["Datei", "Blatt", "Zeile Nr.", "Datenpunkt", "Benutzeradresse", "Typ"]
    records: list[dict] = []

    for document in session.scalars(select(db.Document).order_by(db.Document.file_name)):
        if document.profile_id != profile_id:
            continue
        columns = {c.idx: c for c in document.columns}
        for row in document.rows:
            if row.kind != "daten" or not row.is_countable:
                continue
            record = {
                "Datei": document.file_name,
                "Blatt": document.blatt_nr or "",
                "Zeile Nr.": row.row_no,
                "Datenpunkt": row.klartext,
                "Benutzeradresse": row.bas,
                "Typ": row.qualifier,
                SCHLUESSEL_SPALTE: dokument_schluessel(document),
            }
            for title in column_titles:
                record[title] = None
            bemerkungen = [row.remark] if row.remark else []
            for value in row.values:
                column = columns.get(value.column_idx)
                if column is None:
                    continue
                profile_column = profile.column(column.column_key) if column.column_key else None
                if profile_column is None or profile_column.is_note:
                    if value.note:
                        bemerkungen.append(value.note)
                    continue
                if value.count is not None:
                    record[profile_column.label] = value.count
                elif value.note:
                    bemerkungen.append(value.note)
            record["Bemerkungen"] = " | ".join(b for b in bemerkungen if b)
            records.append(record)

    return pd.DataFrame(
        records, columns=kopf + column_titles + ["Bemerkungen", SCHLUESSEL_SPALTE])


def profiles_in_use(session: Session) -> list[str]:
    """Profile, zu denen tatsaechlich Daten in der Datenbank liegen."""
    ids = {d.profile_id for d in session.scalars(select(db.Document)) if d.profile_id}
    return sorted(ids)


def pruefbericht(session: Session) -> pd.DataFrame:
    """Nachvollziehbare Dokumentation der Zaehlung.

    Enthaelt je Datei den Abgleich mit der Zeile "Summe Funktionen" sowie jede
    Zeile, die bewusst nicht gezaehlt wurde - mit Begruendung und Fundstelle.
    """
    records: list[dict] = []
    for document in session.scalars(select(db.Document)):
        columns = {c.idx: c for c in document.columns}
        eigene: dict[int, float] = {}
        gemeldet: dict[int, float] = {}
        for row in document.rows:
            for value in row.values:
                if value.count is None:
                    continue
                if row.kind == "daten":
                    eigene[value.column_idx] = eigene.get(value.column_idx, 0.0) + value.count
                elif row.kind == "summe":
                    gemeldet[value.column_idx] = gemeldet.get(value.column_idx, 0.0) + value.count

        for column_idx in sorted(set(eigene) | set(gemeldet)):
            column = columns.get(column_idx)
            eigener = eigene.get(column_idx, 0.0)
            listen = gemeldet.get(column_idx)
            records.append({
                "datei": document.file_name,
                "art": "Summenabgleich",
                "spalte": column.address if column else str(column_idx),
                "bezeichnung": column.label if column else "",
                "seite": "",
                "zeile": "",
                "wert_liste": listen if listen is not None else "",
                "wert_eigene_zaehlung": eigener,
                "bewertung": ("nicht in der Liste ausgewiesen" if listen is None
                              else "stimmt überein" if abs(listen - eigener) < 1e-6
                              else "ABWEICHUNG"),
                "begruendung": "",
            })

        for row in document.rows:
            if row.kind == "daten":
                continue
            summe = sum(v.count or 0.0 for v in row.values)
            records.append({
                "datei": document.file_name,
                "art": "nicht gezählte Zeile",
                "spalte": "",
                "bezeichnung": row.klartext,
                "seite": row.page_index + 1,
                "zeile": row.row_no,
                "wert_liste": summe,
                "wert_eigene_zaehlung": 0.0,
                "bewertung": row.kind,
                "begruendung": row.exclusion_reason,
            })
    return pd.DataFrame(records, columns=[
        "datei", "art", "spalte", "bezeichnung", "seite", "zeile",
        "wert_liste", "wert_eigene_zaehlung", "bewertung", "begruendung",
    ])


def datei_spalten_matrix(session: Session, profile_id: str) -> pd.DataFrame:
    """Mengen je Datei und Funktionsspalte - die Grundlage der Uebersicht.

    Eine Zeile je importierter Datei, eine Spalte je Funktionsspalte des
    Profils. Genau die Form, in der die Mengen fuer die Kalkulation gebraucht
    werden: was steckt in welcher Liste, und was ergibt das in Summe.
    """
    profile = get_profile(profile_id)
    if profile is None:
        return pd.DataFrame()

    kopf = ["Datei", "Projekt", "Anlage", "Gewerk", "Blatt", "Datenpunkte"]
    records: list[dict] = []
    for document in session.scalars(
        select(db.Document).order_by(db.Document.file_name)
    ):
        if document.profile_id != profile_id:
            continue
        columns = {c.idx: c for c in document.columns}
        record = {
            "Datei": document.file_name,
            "Projekt": document.projekt or document.auftraggeber,
            "Anlage": document.anlage,
            "Gewerk": document.gewerk,
            "Blatt": document.blatt_nr,
            "Datenpunkte": 0,
            SCHLUESSEL_SPALTE: dokument_schluessel(document),
        }
        for column in profile.columns:
            record[column.key] = 0.0

        for row in document.rows:
            if row.kind != "daten" or not row.is_countable:
                continue
            record["Datenpunkte"] += 1
            for value in row.values:
                if value.count is None:
                    continue
                column = columns.get(value.column_idx)
                if column is None or not column.column_key:
                    continue
                if column.column_key in record:
                    record[column.column_key] += value.count
        records.append(record)

    return pd.DataFrame(
        records, columns=kopf + [c.key for c in profile.columns] + [SCHLUESSEL_SPALTE])
