"""Lokale Oberflaeche (Streamlit).

Start:  streamlit run vdi3814/ui/app.py   bzw. Doppelklick auf die EXE.

Aufbau: Projekt waehlen -> Dateien importieren -> Ergebnis pruefen und ggf.
korrigieren -> Auswertung, Kosten, Export. Alles laeuft lokal auf diesem
Rechner; es werden keine Daten nach aussen gegeben.
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from sqlalchemy.orm import Session                                     # noqa: E402

from vdi3814 import aggregate, db, evidence, export_excel, projects     # noqa: E402
from vdi3814.config import SETTINGS                                     # noqa: E402
from vdi3814.costs import apply_prices, price_template, total_cost      # noqa: E402
from vdi3814.models import Cell, DocumentResult                         # noqa: E402
from vdi3814.pipeline import process_file                               # noqa: E402
from vdi3814.profiles_loader import load_profiles                       # noqa: E402
from vdi3814.vision import ocr                                          # noqa: E402

ROW_FIELDS = ["Zeile", "Datenpunkt", "Benutzeradresse", "Typ", "Bemerkung"]
ART_SYMBOL = {"abweichung": "🔴", "ausgeschlossen": "🟠", "hinweis": "🔵"}

st.set_page_config(page_title="VDI 3814 DP-Checker", page_icon="📋", layout="wide")


# --------------------------------------------------------------------------
# Grundlagen
# --------------------------------------------------------------------------

@st.cache_resource(show_spinner=False)
def get_engine(path: str):
    return db.make_engine(path)


def state() -> dict:
    st.session_state.setdefault("results", {})
    return st.session_state


def zahl(wert: float, nachkomma: int = 2) -> str:
    text = f"{wert:,.{nachkomma}f}"
    return text.replace(",", "X").replace(".", ",").replace("X", ".")


# --------------------------------------------------------------------------
# Seitenleiste: Projektwahl
# --------------------------------------------------------------------------

with st.sidebar:
    st.title("VDI 3814 DP-Checker")
    st.caption("Datenpunkte aus GA-Funktionslisten auszählen – offline auf diesem Rechner")

    projektliste = projects.list_projects()
    namen = [p.name for p in projektliste]
    aktuell = st.session_state.get("projekt", namen[0])
    if aktuell not in namen:
        aktuell = namen[0]
    projektname = st.selectbox("Projekt", namen, index=namen.index(aktuell),
                               help="Jedes Projekt hat eine eigene Datenbank. "
                                    "So bleiben Auswertungen sauber getrennt.")
    st.session_state["projekt"] = projektname
    projekt = projects.resolve(projektname)

    with st.form("neues_projekt", clear_on_submit=True):
        neuer_name = st.text_input("Neues Projekt anlegen", placeholder="z. B. Neubau Musterstadt")
        if st.form_submit_button("Anlegen") and neuer_name.strip():
            projects.create(neuer_name.strip())
            st.session_state["projekt"] = neuer_name.strip()
            st.rerun()

    st.divider()
    st.caption(f"Ablage: `{projekt.path}`")
    st.caption(f"Größe: {projekt.size_mb} MB")
    if ocr.available():
        st.caption("✅ Excel, PDF und Scans auswertbar – keine Zusatzsoftware nötig")
    else:
        st.caption("✅ Excel und PDF mit Textebene auswertbar\n\n"
                   "⚠️ Für gescannte Listen fehlt die Texterkennung "
                   "(in der installierten Programmversion ist sie enthalten)")

engine = get_engine(str(projekt.db_path))


def load_raw() -> pd.DataFrame:
    with Session(engine) as session:
        return aggregate.raw_dataframe(session)


tab_import, tab_preview, tab_check, tab_overview, tab_costs, tab_export, tab_data = st.tabs(
    ["1 · Import", "2 · Prüfen & korrigieren", "3 · Nachweis & Differenzen",
     "4 · Gesamtübersicht", "5 · Kostenschätzung", "6 · Export", "7 · Projekt & Daten"]
)


# --------------------------------------------------------------------------
# 1 - Import
# --------------------------------------------------------------------------

with tab_import:
    st.subheader("Funktionslisten importieren")
    st.write(
        f"Dateien in **{projektname}** einlesen. Unterstützt werden PDF (auch mehrseitig "
        "und gescannt), PNG/JPEG sowie die Excel-Vorlagen. Regelschemata in denselben "
        "Dateien werden erkannt und übersprungen."
    )
    with Session(engine) as session:
        vorhandene = len(db.list_documents(session))
    if vorhandene == 0:
        st.info(
            "**So gehen Sie vor**\n\n"
            f"1. Unten die Funktionslisten auswählen – sie landen im Projekt „{projektname}“.\n"
            "2. Reiter 2: Ergebnis prüfen, bei Bedarf Werte korrigieren und speichern.\n"
            "3. Reiter 3: Auffälligkeiten anklicken – die Stelle wird im Original gezeigt.\n"
            "4. Reiter 5 und 6: Einheitspreise eintragen und Excel-Datei erzeugen.\n\n"
            "Für ein anderes Bauvorhaben links ein neues Projekt anlegen."
        )

    uploads = st.file_uploader(
        "Dateien auswählen (mehrere gleichzeitig möglich)", accept_multiple_files=True,
        type=["pdf", "png", "jpg", "jpeg", "tif", "tiff", "xlsx", "xlsm", "xls"],
    )
    spalte_a, spalte_b = st.columns([1, 2])
    with spalte_a:
        sofort_speichern = st.checkbox("Nach der Erkennung direkt speichern", value=True)

    if uploads and st.button("Erkennung starten", type="primary"):
        fortschritt = st.progress(0.0)
        protokoll = st.empty()
        ablage = Path(tempfile.mkdtemp(prefix="vdi3814_"))
        for index, upload in enumerate(uploads, start=1):
            pfad = ablage / upload.name
            pfad.write_bytes(upload.getbuffer())
            protokoll.info(f"({index}/{len(uploads)}) {upload.name} …")
            ergebnis = process_file(pfad, progress=lambda text: protokoll.info(text))
            state()["results"][upload.name] = ergebnis
            if sofort_speichern and ergebnis.columns:
                with Session(engine) as session:
                    db.save_document(session, ergebnis)
            fortschritt.progress(index / len(uploads))
        protokoll.empty()
        st.success(f"{len(uploads)} Datei(en) verarbeitet.")

    if state()["results"]:
        st.divider()
        st.markdown("**Ergebnis der Erkennung**")
        zeilen = []
        for name, ergebnis in state()["results"].items():
            zeilen.append({
                "Datei": name,
                "Fassung": ergebnis.fassung or "?",
                "Verfahren": ergebnis.backend,
                "Datenpunkte": len(ergebnis.counted_rows()),
                "Funktionen": ergebnis.grand_total(),
                "Nicht gezählt": len(ergebnis.excluded_rows()),
                "Summenprüfung": "OK" if all(c.matches for c in ergebnis.sum_checks) else "Abweichung",
                "Übersprungen": ", ".join(f"S.{p.page_index + 1}" for p in ergebnis.skipped_pages()),
            })
        st.dataframe(pd.DataFrame(zeilen), width="stretch", hide_index=True)
        for name, ergebnis in state()["results"].items():
            for hinweis in ergebnis.warnings:
                st.warning(f"{name}: {hinweis}")


# --------------------------------------------------------------------------
# 2 - Vorschau und Korrektur
# --------------------------------------------------------------------------

def result_to_frame(result: DocumentResult) -> pd.DataFrame:
    columns = [c for c in result.columns if not c.is_note_column]
    records = []
    for row in result.rows:
        if row.is_sum_row:
            continue
        record = {
            "Zeile": row.row_no,
            "Datenpunkt": row.klartext,
            "Benutzeradresse": row.bas,
            "Typ": row.qualifier,
            "Bemerkung": row.remark,
        }
        for column in columns:
            record[_column_title(column)] = row.value(column.index)
        records.append(record)
    return pd.DataFrame(records, columns=ROW_FIELDS + [_column_title(c) for c in columns])


def _column_title(column) -> str:
    prefix = column.address or str(column.index)
    return f"{prefix} · {column.label or column.column_key or ''}"[:60]


def frame_to_result(frame: pd.DataFrame, result: DocumentResult) -> DocumentResult:
    columns = [c for c in result.columns if not c.is_note_column]
    title_to_index = {_column_title(c): c.index for c in columns}
    data_rows = [row for row in result.rows if not row.is_sum_row]
    for position, record in enumerate(frame.to_dict("records")):
        if position >= len(data_rows):
            break
        row = data_rows[position]
        row.row_no = str(record.get("Zeile", "") or "")
        row.klartext = str(record.get("Datenpunkt", "") or "")
        row.bas = str(record.get("Benutzeradresse", "") or "")
        row.qualifier = str(record.get("Typ", "") or "")
        row.remark = str(record.get("Bemerkung", "") or "")
        alt = {cell.column_index: cell for cell in row.cells}
        neu: list[Cell] = []
        for titel, column_index in title_to_index.items():
            wert = record.get(titel)
            if wert is None or (isinstance(wert, float) and pd.isna(wert)) or wert == "":
                continue
            vorher = alt.get(column_index)
            neu.append(Cell(column_index=column_index, raw_value=str(wert), count=float(wert),
                            bbox=vorher.bbox if vorher else None))
        row.cells = neu
    return result


with tab_preview:
    st.subheader("Erkanntes Ergebnis prüfen und korrigieren")
    ergebnisse = state()["results"]
    if not ergebnisse:
        st.info("Noch nichts importiert – bitte zuerst Reiter 1.")
    else:
        name = st.selectbox("Datei", list(ergebnisse.keys()), key="vorschau_datei")
        ergebnis = ergebnisse[name]

        kopf = st.columns(4)
        kopf[0].metric("Datenpunkte", len(ergebnis.counted_rows()))
        kopf[1].metric("Funktionen", f"{ergebnis.grand_total():g}")
        kopf[2].metric("Nicht gezählt", len(ergebnis.excluded_rows()))
        kopf[3].metric("Fassung", ergebnis.fassung or "?")

        with st.expander("Kopf-/Fußdaten des Dokuments"):
            spalten = st.columns(4)
            for index, feld in enumerate(vars(ergebnis.metadata)):
                with spalten[index % 4]:
                    setattr(ergebnis.metadata, feld,
                            st.text_input(feld, value=getattr(ergebnis.metadata, feld),
                                          key=f"meta_{name}_{feld}"))

        if ergebnis.excluded_rows():
            with st.expander(f"Nicht gezählte Zeilen ({len(ergebnis.excluded_rows())}) – mit Begründung"):
                st.dataframe(pd.DataFrame([
                    {"Zeile": r.row_no, "Beschriftung": r.klartext, "Art": r.kind.value,
                     "Grund": r.exclusion_reason,
                     "Summe der Zeile": sum(c.count or 0 for c in r.cells)}
                    for r in ergebnis.excluded_rows()
                ]), width="stretch", hide_index=True)

        st.caption("Zahlenwerte je Funktionsspalte – Korrekturen wirken sofort auf Summen, "
                   "Kosten und Export. Spaltentitel: Abschnitt.Spalte · Bezeichnung")
        tabelle = st.data_editor(result_to_frame(ergebnis), width="stretch",
                                 num_rows="fixed", key=f"editor_{name}", height=420)
        frame_to_result(tabelle, ergebnis)

        if st.button("In Datenbank speichern", type="primary", key=f"save_{name}"):
            with Session(engine) as session:
                dokument = db.save_document(session, ergebnis)
            st.success(f"Gespeichert in Projekt „{projektname}“ (Dokument {dokument.id}).")


# --------------------------------------------------------------------------
# 3 - Nachweis und Differenzen
# --------------------------------------------------------------------------

with tab_check:
    st.subheader("Nachweis: woher kommt welcher Wert?")
    st.write("Jeder Befund lässt sich anklicken – die Originalseite wird an der "
             "betreffenden Stelle markiert angezeigt.")

    with Session(engine) as session:
        dokumente = {d.id: d for d in db.list_documents(session)}
        befunde = evidence.all_findings(session)

        if not dokumente:
            st.info("Noch keine gespeicherten Dateien in diesem Projekt.")
        else:
            filter_spalten = st.columns([2, 2, 1])
            with filter_spalten[0]:
                datei_filter = st.multiselect("Datei", sorted({d.file_name for d in dokumente.values()}))
            with filter_spalten[1]:
                art_filter = st.multiselect("Art", ["abweichung", "ausgeschlossen", "hinweis"],
                                            default=["abweichung", "ausgeschlossen"])
            with filter_spalten[2]:
                ganze_seite = st.checkbox("Ganze Seite", value=False)

            gefiltert = [
                b for b in befunde
                if (not datei_filter or b.datei in datei_filter)
                and (not art_filter or b.art in art_filter)
            ]
            zusammenfassung = st.columns(3)
            zusammenfassung[0].metric("Abweichungen", sum(1 for b in befunde if b.art == "abweichung"))
            zusammenfassung[1].metric("Nicht gezählte Zeilen mit Werten",
                                      sum(1 for b in befunde if b.art == "ausgeschlossen"))
            zusammenfassung[2].metric("Hinweise", sum(1 for b in befunde if b.art == "hinweis"))

            if not gefiltert:
                st.success("Keine Befunde – die Zählung deckt sich mit den Listen.")
            else:
                auswahl = st.radio(
                    "Befund auswählen",
                    options=list(range(len(gefiltert))),
                    format_func=lambda i: (
                        f"{ART_SYMBOL.get(gefiltert[i].art, '•')} {gefiltert[i].datei} "
                        f"S.{gefiltert[i].anzeige_seite} – {gefiltert[i].titel}"
                    ),
                    key="befund_auswahl",
                )
                befund = gefiltert[auswahl]
                st.info(f"**Warum wurde das erkannt?**\n\n{befund.erklaerung}")
                if befund.werte:
                    werte_spalten = st.columns(len(befund.werte))
                    for spalte, (bezeichnung, wert) in zip(werte_spalten, befund.werte.items()):
                        spalte.metric(bezeichnung, f"{wert:g}")

                dokument = dokumente.get(befund.dokument_id)
                bild = evidence.render_finding(dokument, befund, ganze_seite=ganze_seite) \
                    if dokument else None
                if bild is not None:
                    st.image(bild, caption=f"{befund.datei} – Seite {befund.anzeige_seite}",
                             width="stretch")
                elif dokument is not None and evidence.source_file(dokument) is None:
                    st.warning("Die Originaldatei liegt nicht mehr im Projekt – "
                               "bitte erneut importieren, um den Nachweis zu sehen.")
                else:
                    st.caption("Für Excel-Quellen gibt es kein Seitenbild; "
                               "die Fundstelle steht in den Rohdaten.")


# --------------------------------------------------------------------------
# 4 - Gesamtuebersicht
# --------------------------------------------------------------------------

with tab_overview:
    st.subheader(f"Gesamtübersicht – Projekt „{projektname}“")
    raw = load_raw()
    if raw.empty:
        st.info("Noch keine Daten gespeichert.")
    else:
        filter_spalten = st.columns(4)
        with filter_spalten[0]:
            fassung = st.multiselect("Fassung", sorted(raw["fassung"].unique()))
        with filter_spalten[1]:
            projekt_filter = st.multiselect("Projektangabe", sorted(raw["projekt"].fillna("").unique()))
        with filter_spalten[2]:
            anlage = st.multiselect("Anlage", sorted(raw["anlage"].fillna("").unique()))
        with filter_spalten[3]:
            gewerk = st.multiselect("Gewerk", sorted(raw["gewerk"].fillna("").unique()))
        gefiltert = raw
        for spalte, werte in (("fassung", fassung), ("projekt", projekt_filter),
                              ("anlage", anlage), ("gewerk", gewerk)):
            if werte:
                gefiltert = gefiltert[gefiltert[spalte].isin(werte)]

        kennzahlen = st.columns(4)
        kennzahlen[0].metric("Dateien", gefiltert["dokument_id"].nunique())
        kennzahlen[1].metric("Datenpunkte", gefiltert["datenpunkt"].nunique())
        kennzahlen[2].metric("Funktionen gesamt", f"{gefiltert['wert'].sum():g}")
        kennzahlen[3].metric("Belegte Funktionsspalten", gefiltert["spalte_key"].nunique())

        st.markdown("**Summen je Funktionsspalte**")
        st.dataframe(aggregate.column_summary(gefiltert), width="stretch", hide_index=True)
        links, rechts = st.columns(2)
        with links:
            st.markdown("**Je Spaltengruppe**")
            st.dataframe(aggregate.group_summary(gefiltert), width="stretch", hide_index=True)
        with rechts:
            dimension = st.selectbox("Aufschlüsselung nach", ["projekt", "anlage", "gewerk", "datei"])
            st.dataframe(aggregate.by_dimension(gefiltert, dimension),
                         width="stretch", hide_index=True)
        with st.expander("Drilldown auf jede einzelne Zelle"):
            st.dataframe(gefiltert, width="stretch", hide_index=True)


# --------------------------------------------------------------------------
# 5 - Kostenschaetzung
# --------------------------------------------------------------------------

with tab_costs:
    st.subheader("Kostenschätzung")
    st.write("Einheitspreis je Funktionsspalte eintragen – die Kosten werden sofort neu "
             "berechnet. Gespeicherte Preise gelten auch für den Excel-Export.")

    raw = load_raw()
    with Session(engine) as session:
        gespeicherte_preise = db.get_unit_prices(session)
    summary = aggregate.column_summary(raw)

    alle_spalten = st.checkbox("Alle Spalten des Profils anzeigen (auch ohne Mengen)",
                               value=summary.empty)
    if alle_spalten:
        profil_id = st.selectbox("Profil", [p.id for p in load_profiles()])
        basis = price_template(profil_id, gespeicherte_preise)
        mengen = dict(zip(summary["spalte_key"], summary["menge"])) if not summary.empty else {}
        basis["menge"] = [float(mengen.get(key, 0.0)) for key in basis["spalte_key"]]
        basis = basis[["spalte_adresse", "gruppe", "untergruppe", "spalte", "spalte_key",
                       "menge", "einheitspreis"]]
    elif summary.empty:
        st.info("Noch keine Daten – Haken oben setzen, um Preise vorab zu pflegen.")
        basis = None
    else:
        basis = summary[["spalte_adresse", "gruppe", "untergruppe", "spalte",
                         "spalte_key", "menge"]].copy()
        basis["einheitspreis"] = [float(gespeicherte_preise.get(key, 0.0))
                                  for key in basis["spalte_key"]]

    if basis is not None:
        bearbeitet = st.data_editor(
            basis, width="stretch", hide_index=True, num_rows="fixed", height=420,
            key="preis_editor",
            column_config={
                "spalte_adresse": st.column_config.TextColumn("Abschnitt.Spalte"),
                "gruppe": st.column_config.TextColumn("Gruppe"),
                "untergruppe": st.column_config.TextColumn("Untergruppe"),
                "spalte": st.column_config.TextColumn("Funktionsspalte"),
                "spalte_key": st.column_config.TextColumn("Schlüssel"),
                "menge": st.column_config.NumberColumn("Menge", format="%.0f"),
                "einheitspreis": st.column_config.NumberColumn(
                    f"Einheitspreis [{SETTINGS.currency}]", min_value=0.0, step=1.0, format="%.2f"),
            },
            disabled=["spalte_adresse", "gruppe", "untergruppe", "spalte", "spalte_key", "menge"],
        )
        eingegeben = {str(key): float(value or 0.0)
                      for key, value in zip(bearbeitet["spalte_key"], bearbeitet["einheitspreis"])}
        ergebnis_preise = apply_prices(summary, {**gespeicherte_preise, **eingegeben}) \
            if not summary.empty else bearbeitet.assign(kosten=0.0)

        kennzahlen = st.columns(3)
        kennzahlen[0].metric("Funktionen gesamt",
                             f"{summary['menge'].sum():g}" if not summary.empty else "0")
        kennzahlen[1].metric("Spalten mit Einheitspreis",
                             sum(1 for value in eingegeben.values() if value > 0))
        kennzahlen[2].metric("Geschätzte Gesamtkosten",
                             f"{zahl(total_cost(ergebnis_preise))} {SETTINGS.currency}")

        if not summary.empty:
            anzeige = ergebnis_preise[ergebnis_preise["kosten"] > 0] \
                if (ergebnis_preise["kosten"] > 0).any() else ergebnis_preise
            st.dataframe(anzeige[["spalte_adresse", "gruppe", "spalte", "menge",
                                  "einheitspreis", "kosten"]],
                         width="stretch", hide_index=True)

        if st.button("Einheitspreise speichern", type="primary"):
            with Session(engine) as session:
                db.set_unit_prices(session, eingegeben)
            st.success("Einheitspreise gespeichert.")


# --------------------------------------------------------------------------
# 6 - Export
# --------------------------------------------------------------------------

with tab_export:
    st.subheader("Excel-Export")
    st.write("Der Export enthält die Funktionsliste im Original-Layout, die Summen je "
             "Funktionsspalte, ein Kostenblatt mit Formeln sowie alle Rohdaten mit "
             "Quellenangabe.")
    raw = load_raw()
    if raw.empty:
        st.info("Noch keine Daten gespeichert.")
    else:
        ziel = st.text_input("Zieldatei", value=str(projekt.path / f"{projekt.path.name}_Auswertung.xlsx"))
        if st.button("Excel erzeugen", type="primary"):
            with Session(engine) as session:
                pfad = export_excel.export_workbook(
                    ziel,
                    summary=aggregate.column_summary(raw),
                    raw=raw,
                    documents=aggregate.documents_frame(session),
                    footnotes=aggregate.footnotes_frame(session),
                    prices=db.get_unit_prices(session),
                    projects=aggregate.pivot_projects(raw),
                    layouts={profil: aggregate.vdi_layout_frame(session, profil)
                             for profil in aggregate.profiles_in_use(session)},
                    pruefung=aggregate.pruefbericht(session),
                )
            st.success(f"Geschrieben: {pfad}")
            st.download_button("Datei herunterladen", data=Path(pfad).read_bytes(),
                               file_name=Path(pfad).name,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# --------------------------------------------------------------------------
# 7 - Projekt und Daten
# --------------------------------------------------------------------------

with tab_data:
    st.subheader(f"Projekt „{projektname}“ verwalten")

    with Session(engine) as session:
        dokumente = aggregate.documents_frame(session)

    if dokumente.empty:
        st.info("Dieses Projekt enthält noch keine Dateien.")
    else:
        st.markdown("**Importierte Dateien** – zum Löschen ankreuzen")
        auswahl_frame = dokumente[["dokument_id", "datei", "fassung", "verfahren",
                                   "datenpunkte", "summe_funktionen", "importiert_am"]].copy()
        auswahl_frame.insert(0, "löschen", False)
        bearbeitet = st.data_editor(
            auswahl_frame, width="stretch", hide_index=True, num_rows="fixed",
            key="dokument_auswahl",
            disabled=["dokument_id", "datei", "fassung", "verfahren", "datenpunkte",
                      "summe_funktionen", "importiert_am"],
        )
        zu_loeschen = [int(row["dokument_id"]) for _, row in bearbeitet.iterrows() if row["löschen"]]
        if zu_loeschen and st.button(f"{len(zu_loeschen)} Datei(en) löschen", type="primary"):
            with Session(engine) as session:
                for dokument_id in zu_loeschen:
                    db.delete_document(session, dokument_id)
            st.success(f"{len(zu_loeschen)} Datei(en) gelöscht.")
            st.rerun()

        with st.expander("Details je Datei (übersprungene Seiten, Hinweise)"):
            st.dataframe(dokumente, width="stretch", hide_index=True)

    st.divider()
    st.markdown("**Projekt zurücksetzen oder löschen**")
    links, rechts = st.columns(2)
    with links:
        sicher_leeren = st.checkbox("Ja, alle Daten dieses Projekts entfernen",
                                    key="sicher_leeren")
        if st.button("Projekt leeren", disabled=not sicher_leeren):
            get_engine.clear()
            db.dispose_all()
            projects.clear(projektname)
            st.success(f"Projekt „{projektname}“ ist jetzt leer.")
            st.rerun()
    with rechts:
        sicher_loeschen = st.checkbox("Ja, dieses Projekt vollständig löschen",
                                      key="sicher_loeschen")
        if st.button("Projekt löschen", disabled=not sicher_loeschen):
            get_engine.clear()
            db.dispose_all()
            projects.delete(projektname)
            st.session_state.pop("projekt", None)
            st.success(f"Projekt „{projektname}“ wurde gelöscht.")
            st.rerun()

    st.divider()
    st.markdown("**Alle Projekte**")
    st.dataframe(pd.DataFrame([
        {"Projekt": p.name, "Größe (MB)": p.size_mb,
         "Zuletzt geändert": p.changed_at.strftime("%d.%m.%Y %H:%M") if p.changed_at else "-",
         "Ablage": str(p.path)}
        for p in projects.list_projects()
    ]), width="stretch", hide_index=True)
