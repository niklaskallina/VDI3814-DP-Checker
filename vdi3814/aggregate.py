"""Aggregation und Drilldown ueber alle importierten Funktionslisten."""

from __future__ import annotations

import pandas as pd
from sqlalchemy import select
from sqlalchemy.orm import Session

from . import db, schwerpunkt
from .profiles_loader import get_profile, load_profiles

RAW_COLUMNS = [
    "dokument_id", "datei", "fassung", "profil", "projekt", "auftraggeber", "anlage",
    "schwerpunkt", "schwerpunkt_bezeichnung", "gewerk", "planersteller", "blatt",
    "seite", "zeile_id", "zeile_nr", "datenpunkt", "bas",
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
        profile = get_profile(document.profile_id) if document.profile_id else None
        for row in document.rows:
            # Nur echte Datenpunkte: keine Summen-, Uebertrags-, Leer-,
            # Kopf- oder Fussbereichszeilen.
            if row.kind != "daten":
                continue
            for value in row.values:
                column = columns.get(value.column_idx)
                if column is None or column.is_note:
                    continue
                # Bezeichnungen einheitlich aus dem Profil: dieselbe Spalte
                # heisst in jeder Datei gleich, egal wie der Kopf gelesen wurde.
                vorlage = (profile.column(column.column_key)
                           if profile is not None and column.column_key else None)
                records.append({
                    "dokument_id": document.id,
                    "datei": document.file_name,
                    "fassung": document.fassung,
                    "profil": document.profile_id,
                    "projekt": document.projekt or document.auftraggeber,
                    "auftraggeber": document.auftraggeber,
                    "anlage": document.anlage,
                    "schwerpunkt": row.schwerpunkt or schwerpunkt.OHNE_ZUORDNUNG,
                    "schwerpunkt_bezeichnung": row.schwerpunkt_text or "",
                    "gewerk": document.gewerk,
                    "planersteller": document.planersteller,
                    "blatt": document.blatt_nr,
                    "seite": row.page_index + 1,
                    "zeile_id": row.id,
                    "zeile_nr": row.row_no,
                    "datenpunkt": row.klartext,
                    "bas": row.bas,
                    "qualifikation": row.qualifier,
                    "bemerkung": row.remark,
                    "spalte_adresse": ((vorlage.address if vorlage else "") or column.address
                                       or _column_address(column.column_key)),
                    "spalte_key": column.column_key,
                    "gruppe": (vorlage.group if vorlage else "") or column.group_name,
                    "untergruppe": (vorlage.subgroup if vorlage else "") or column.subgroup,
                    "spalte": (vorlage.label if vorlage else "") or column.label,
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
    """Summen je Funktionsspalte ueber alle importierten Listen.

    Eine Zeile je Spalte des Profils - zusammengefasst ueber den
    Spaltenschluessel, nicht ueber die gelesene Beschriftung. Liest eine Datei
    die Ueberschrift etwas anders (OCR, Umbruch, Gruppenzeile), entstand
    frueher eine zweite Zeile fuer dieselbe Spalte; im Excel-Export verwiesen
    dann beide auf dieselbe Summe und Menge wie Kosten zaehlten doppelt.
    Bezeichnungen kommen deshalb aus dem Profil. Spalten ohne Zuordnung
    bleiben je Adresse und Beschriftung getrennt.

    "datenpunkte" zaehlt die Datenpunkte (Zeilen) mit einem Wert in der Spalte.
    """
    spalten = ["profil", "spalte_adresse", "spalte_key", "gruppe", "untergruppe", "spalte",
               "menge", "dokumente", "datenpunkte"]
    if frame.empty:
        return pd.DataFrame(columns=spalten)
    daten = frame.copy()
    schluessel = daten["spalte_key"].fillna("").astype(str)
    daten["_gruppe"] = [
        key if key else f"?{adresse}|{spalte}"
        for key, adresse, spalte in zip(schluessel, daten["spalte_adresse"].fillna(""),
                                        daten["spalte"].fillna(""))
    ]
    daten["_belegt"] = daten["zeile_id"].where(daten["wert"].fillna(0) != 0)
    grouped = (
        daten.groupby(["profil", "_gruppe"], dropna=False)
        .agg(spalte_key=("spalte_key", "first"),
             spalte_adresse=("spalte_adresse", "first"),
             gruppe=("gruppe", "first"),
             untergruppe=("untergruppe", "first"),
             spalte=("spalte", "first"),
             menge=("wert", "sum"),
             dokumente=("dokument_id", "nunique"),
             datenpunkte=("_belegt", "nunique"))
        .reset_index()
    )
    # Einheitliche Bezeichnungen aus dem Profil
    for index, record in grouped.iterrows():
        profile = get_profile(str(record["profil"])) if record["profil"] else None
        eintrag = profile.column(str(record["spalte_key"])) if profile and record["spalte_key"] else None
        if eintrag is None:
            continue
        grouped.at[index, "spalte_adresse"] = eintrag.address or record["spalte_adresse"]
        grouped.at[index, "gruppe"] = eintrag.group or record["gruppe"]
        grouped.at[index, "untergruppe"] = eintrag.subgroup or record["untergruppe"]
        grouped.at[index, "spalte"] = eintrag.label or record["spalte"]

    orders = {profil: _profile_order(profil) for profil in grouped["profil"].unique()}
    grouped["_sort"] = [
        orders.get(profil, {}).get(key, 9999)
        for profil, key in zip(grouped["profil"], grouped["spalte_key"])
    ]
    grouped = grouped.sort_values(["profil", "_sort", "spalte_adresse"])
    return grouped[spalten].reset_index(drop=True)


def moegliche_dubletten(frame: pd.DataFrame, mindestanteil: float = 0.8) -> pd.DataFrame:
    """Dateien, die (fast) dieselben Datenpunkte enthalten.

    Derselbe Plansatz liegt oft doppelt vor - als PDF und als Excel, oder in
    zwei Planstaenden. Der Dateivergleich ueber den Hash erkennt das nicht,
    die Mengen zaehlten dann doppelt. Verglichen wird deshalb der Inhalt: je
    Datenpunkt Benennung und belegte Werte. Stimmen mindestens
    ``mindestanteil`` der Datenpunkte der kleineren Datei ueberein, wird das
    Paar gemeldet - entscheiden muss der Anwender.
    """
    spalten = ["datei_a", "datei_b", "dokument_a", "dokument_b", "gemeinsam", "anteil"]
    if frame.empty:
        return pd.DataFrame(columns=spalten)
    from .textutil import normalize

    signaturen: dict[int, set] = {}
    dateien: dict[int, str] = {}
    belegt = frame[frame["wert"].fillna(0) != 0]
    for (dokument_id, _zeile), gruppe in belegt.groupby(["dokument_id", "zeile_id"]):
        werte = frozenset(zip(gruppe["spalte_key"], gruppe["wert"]))
        name = normalize(str(gruppe["datenpunkt"].iloc[0] or ""))
        signaturen.setdefault(int(dokument_id), set()).add((name, werte))
        dateien[int(dokument_id)] = str(gruppe["datei"].iloc[0])

    paare: list[dict] = []
    nummern = sorted(signaturen)
    for position, a in enumerate(nummern):
        for b in nummern[position + 1:]:
            kleiner = min(len(signaturen[a]), len(signaturen[b]))
            if kleiner < 3:
                continue
            gemeinsam = len(signaturen[a] & signaturen[b])
            anteil = gemeinsam / kleiner
            if anteil >= mindestanteil:
                paare.append({"datei_a": dateien[a], "datei_b": dateien[b],
                              "dokument_a": a, "dokument_b": b,
                              "gemeinsam": gemeinsam, "anteil": round(anteil, 3)})
    return pd.DataFrame(paare, columns=spalten)


def group_summary(frame: pd.DataFrame) -> pd.DataFrame:
    """Summen je Spaltengruppe (Ein-/Ausgabe, Verarbeitung, ...)."""
    if frame.empty:
        return pd.DataFrame(columns=["fassung", "gruppe", "menge"])
    return (
        frame.groupby(["fassung", "gruppe"], dropna=False)["wert"].sum()
        .reset_index(name="menge").sort_values(["fassung", "gruppe"]).reset_index(drop=True)
    )


def by_dimension(frame: pd.DataFrame, dimension: str) -> pd.DataFrame:
    """Summen je Projekt / Anlage / Schwerpunkt / Gewerk / Datei.

    "datenpunkte" zaehlt echte Datenpunktzeilen, nicht belegte Zellen -
    deshalb ueber die Zeilen-ID und nicht ueber die Benennung (die kann sich
    in zwei Listen wiederholen).
    """
    if frame.empty or dimension not in frame.columns:
        return pd.DataFrame(columns=[dimension, "menge", "datenpunkte", "dokumente"])
    return (
        frame.groupby(dimension, dropna=False)
        .agg(menge=("wert", "sum"),
             datenpunkte=("zeile_id", "nunique"),
             dokumente=("dokument_id", "nunique"))
        .reset_index().sort_values("menge", ascending=False).reset_index(drop=True)
    )


def schwerpunkt_summary(frame: pd.DataFrame) -> pd.DataFrame:
    """Datenpunkte und Funktionen je Automations-/Informationsschwerpunkt.

    Beantwortet die Frage "wie viel steckt in ASP01, wie viel in ASP02?" -
    quer ueber alle importierten Listen. Traegt eine Liste denselben ASP mit
    unterschiedlicher Bezeichnung, bleiben die Angaben getrennt: dann ist im
    Plansatz etwas uneinheitlich und das soll sichtbar bleiben.
    """
    spalten = ["schwerpunkt", "bezeichnung", "art", "datenpunkte", "menge",
               "dokumente", "dateien"]
    if frame.empty:
        return pd.DataFrame(columns=spalten)
    grouped = (
        frame.groupby(["schwerpunkt", "schwerpunkt_bezeichnung"], dropna=False)
        .agg(datenpunkte=("zeile_id", "nunique"),
             menge=("wert", "sum"),
             dokumente=("dokument_id", "nunique"),
             dateien=("datei", lambda werte: ", ".join(sorted(set(werte)))))
        .reset_index().rename(columns={"schwerpunkt_bezeichnung": "bezeichnung"})
    )
    grouped["art"] = [
        "ASP" if str(kennung).upper().startswith("ASP")
        else "ISP" if str(kennung).upper().startswith("ISP") else ""
        for kennung in grouped["schwerpunkt"]
    ]
    # Zeilen ohne Zuordnung ans Ende, sonst nach Kennung sortieren
    grouped["_sort"] = [1 if k == schwerpunkt.OHNE_ZUORDNUNG else 0
                        for k in grouped["schwerpunkt"]]
    grouped = grouped.sort_values(["_sort", "schwerpunkt"]).drop(columns="_sort")
    return grouped[spalten].reset_index(drop=True)


def pivot_projects(frame: pd.DataFrame) -> pd.DataFrame:
    """Kreuztabelle Projekt/Anlage x Spaltengruppe."""
    if frame.empty:
        return pd.DataFrame()
    return pd.pivot_table(frame, index=["projekt", "anlage"], columns="gruppe",
                          values="wert", aggfunc="sum", fill_value=0, margins=True,
                          margins_name="Gesamt").reset_index()


def _schwerpunkt_der_seite(document: db.Document, page_index: int) -> str:
    for page in document.pages:
        if page.page_index == page_index and page.schwerpunkt:
            return page.schwerpunkt
    return ""


def _schwerpunkte_der_datei(document: db.Document) -> list[str]:
    """Alle Schwerpunkte einer Datei in der Reihenfolge des Dokuments."""
    gefunden: list[str] = []
    for row in document.rows:
        if row.kind != "daten" or not row.schwerpunkt:
            continue
        if row.schwerpunkt not in gefunden:
            gefunden.append(row.schwerpunkt)
    return gefunden


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
            "schwerpunkte": ", ".join(_schwerpunkte_der_datei(document)),
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
    kopf = ["Schwerpunkt", "Bezeichnung", "Datei", "Blatt", "Zeile Nr.",
            "Datenpunkt", "Benutzeradresse", "Typ"]
    records: list[dict] = []

    for document in session.scalars(select(db.Document).order_by(db.Document.file_name)):
        if document.profile_id != profile_id:
            continue
        columns = {c.idx: c for c in document.columns}
        for row in document.rows:
            if row.kind != "daten" or not row.is_countable:
                continue
            record = {
                "Schwerpunkt": row.schwerpunkt or "",
                "Bezeichnung": row.schwerpunkt_text or "",
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

    Enthaelt je Datei und Seite den Abgleich mit der Zeile "Summe Funktionen"
    sowie jede Zeile, die bewusst nicht gezaehlt wurde - mit Begruendung und
    Fundstelle.

    Verglichen wird je Seite und mit Uebertrag (evidence.seitenabgleich). Der
    fruehere Abgleich ueber die ganze Datei hat bei mehrseitigen Listen jede
    Seitensumme aufaddiert und dadurch Abweichungen gemeldet, die es nicht gab.
    """
    from .evidence import seitenabgleich

    records: list[dict] = []
    for document in session.scalars(select(db.Document)):
        columns = {c.idx: c for c in document.columns}
        abgleich = seitenabgleich(document)
        for eintrag in abgleich:
            column = columns.get(eintrag["spalte_idx"])
            listen, eigener = eintrag["gemeldet"], eintrag["eigen"]
            records.append({
                "datei": document.file_name,
                "schwerpunkt": _schwerpunkt_der_seite(document, eintrag["seite"]),
                "art": "Summenabgleich",
                "spalte": column.address if column else str(eintrag["spalte_idx"]),
                "bezeichnung": column.label if column else "",
                "seite": eintrag["seite"] + 1,
                "zeile": eintrag["zeile"].row_no if eintrag["zeile"] is not None else "",
                "wert_liste": listen if listen is not None else "",
                "wert_eigene_zaehlung": eigener,
                "bewertung": ("nicht in der Liste ausgewiesen" if listen is None
                              else "stimmt überein" if abs(listen - eigener) < 1e-6
                              else "ABWEICHUNG"),
                "begruendung": (f"inkl. Übertrag {eintrag['uebertrag']:g}"
                                if eintrag["uebertrag"] else ""),
            })
        if not abgleich:
            records.append({
                "datei": document.file_name,
                "schwerpunkt": ", ".join(_schwerpunkte_der_datei(document)),
                "art": "kein Summenabgleich",
                "spalte": "", "bezeichnung": "", "seite": "", "zeile": "",
                "wert_liste": "", "wert_eigene_zaehlung": "",
                "bewertung": "nicht prüfbar",
                "begruendung": "Die Liste enthält keine Summenzeile zum Abgleich.",
            })

        for row in document.rows:
            if row.kind == "daten":
                continue
            summe = sum(v.count or 0.0 for v in row.values)
            records.append({
                "datei": document.file_name,
                "schwerpunkt": row.schwerpunkt or "",
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
        "datei", "schwerpunkt", "art", "spalte", "bezeichnung", "seite", "zeile",
        "wert_liste", "wert_eigene_zaehlung", "bewertung", "begruendung",
    ])


def schwerpunkt_matrix(session: Session, profile_id: str) -> pd.DataFrame:
    """Mengen je Automations-/Informationsschwerpunkt und Funktionsspalte.

    Aufbau wie datei_spalten_matrix(), nur ist die Zeile hier kein Dokument,
    sondern ein Schwerpunkt. Gruppiert wird je Datei und Schwerpunkt: die
    laufende Nummer eines ASP gilt nur innerhalb eines Plansatzes, ein
    dateiuebergreifendes Zusammenfassen von "ASP01" wuerde zwei verschiedene
    Schwerpunkte vermischen. Die Gesamtsicht ueber alle Listen liefert
    schwerpunkt_summary().
    """
    profile = get_profile(profile_id)
    if profile is None:
        return pd.DataFrame()

    kopf = ["Schwerpunkt", "Bezeichnung", "Datei", "Projekt", "Anlage", "Datenpunkte"]
    records: list[dict] = []
    for document in session.scalars(
        select(db.Document).order_by(db.Document.file_name)
    ):
        if document.profile_id != profile_id:
            continue
        columns = {c.idx: c for c in document.columns}
        je_schwerpunkt: dict[str, dict] = {}
        for row in document.rows:
            if row.kind != "daten" or not row.is_countable:
                continue
            kennung = row.schwerpunkt or schwerpunkt.OHNE_ZUORDNUNG
            record = je_schwerpunkt.get(kennung)
            if record is None:
                record = {
                    "Schwerpunkt": kennung,
                    "Bezeichnung": row.schwerpunkt_text or "",
                    "Datei": document.file_name,
                    "Projekt": document.projekt or document.auftraggeber,
                    "Anlage": document.anlage,
                    "Datenpunkte": 0,
                    # Aus welcher Liste stammt die Zeile? Der Excel-Export
                    # zieht darueber genau deren Zeilen zusammen.
                    SCHLUESSEL_SPALTE: dokument_schluessel(document),
                }
                for column in profile.columns:
                    record[column.key] = 0.0
                je_schwerpunkt[kennung] = record
            if not record["Bezeichnung"] and row.schwerpunkt_text:
                record["Bezeichnung"] = row.schwerpunkt_text
            record["Datenpunkte"] += 1
            for value in row.values:
                if value.count is None:
                    continue
                column = columns.get(value.column_idx)
                if column is None or not column.column_key:
                    continue
                if column.column_key in record:
                    record[column.column_key] += value.count
        records.extend(je_schwerpunkt.values())

    frame = pd.DataFrame(
        records, columns=kopf + [c.key for c in profile.columns] + [SCHLUESSEL_SPALTE])
    if frame.empty:
        return frame
    # Nach Kennung sortieren, Zeilen ohne Zuordnung ans Ende
    frame["_sort"] = [1 if k == schwerpunkt.OHNE_ZUORDNUNG else 0
                      for k in frame["Schwerpunkt"]]
    frame = frame.sort_values(["_sort", "Schwerpunkt", "Datei"]).drop(columns="_sort")
    return frame.reset_index(drop=True)


def datei_spalten_matrix(session: Session, profile_id: str) -> pd.DataFrame:
    """Mengen je Datei und Funktionsspalte - die Grundlage der Uebersicht.

    Eine Zeile je importierter Datei, eine Spalte je Funktionsspalte des
    Profils. Genau die Form, in der die Mengen fuer die Kalkulation gebraucht
    werden: was steckt in welcher Liste, und was ergibt das in Summe.
    """
    profile = get_profile(profile_id)
    if profile is None:
        return pd.DataFrame()

    kopf = ["Datei", "Projekt", "Anlage", "Schwerpunkt", "Gewerk", "Blatt", "Datenpunkte"]
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
            # Enthaelt die Datei mehrere Schwerpunkte, stehen sie hier alle;
            # aufgeschluesselt sind sie in schwerpunkt_matrix().
            "Schwerpunkt": ", ".join(_schwerpunkte_der_datei(document)),
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
