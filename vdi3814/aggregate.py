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
            if row.is_sum_row:
                continue                      # Summenzeilen nie mitzaehlen
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
            "datenpunkte": len([r for r in document.rows if not r.is_sum_row]),
            "summe_funktionen": sum(v.count or 0.0 for r in document.rows if not r.is_sum_row
                                    for v in r.values),
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
