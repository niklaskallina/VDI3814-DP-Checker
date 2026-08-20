"""Lokale Oberflaeche (Streamlit).

Start:  streamlit run vdi3814/ui/app.py

Warum Streamlit und keine Desktop-GUI: die Oberflaeche laeuft ausschliesslich
lokal (127.0.0.1), braucht keine Installation ausser dem Python-Paket, und
tabellarisches Bearbeiten grosser Funktionslisten ist mit st.data_editor
deutlich komfortabler als mit Tkinter/Qt-Tabellen. Es geht dabei nichts ins
Netz - der Browser ist nur die Anzeige.
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from sqlalchemy.orm import Session                                    # noqa: E402

from vdi3814 import aggregate, db, export_excel                        # noqa: E402
from vdi3814.config import SETTINGS                                    # noqa: E402
from vdi3814.costs import apply_prices, price_template, total_cost     # noqa: E402
from vdi3814.models import Cell, DocumentResult                        # noqa: E402
from vdi3814.pipeline import process_file                              # noqa: E402
from vdi3814.profiles_loader import load_profiles                      # noqa: E402
from vdi3814.vision import ocr                                         # noqa: E402
from vdi3814.vision.ollama_client import OllamaUnavailable, OllamaVisionBackend  # noqa: E402

ROW_FIELDS = ["Zeile", "Datenpunkt", "Benutzeradresse", "Typ", "Bemerkung"]

st.set_page_config(page_title="VDI 3814 DP-Checker", page_icon="📋", layout="wide")


# --------------------------------------------------------------------------
# Zustand
# --------------------------------------------------------------------------

@st.cache_resource
def get_engine(path: str):
    return db.make_engine(path)


def state() -> dict:
    st.session_state.setdefault("results", {})      # Dateiname -> DocumentResult
    return st.session_state


def make_backend(use_model: bool, model: str):
    if not use_model:
        return None, "Vision-Modell deaktiviert"
    backend = OllamaVisionBackend(model=model)
    try:
        backend.ensure_ready()
    except OllamaUnavailable as exc:
        return None, str(exc)
    return backend, f"Modell bereit: {model}"


# --------------------------------------------------------------------------
# Umwandlung DocumentResult <-> DataFrame (fuer die Korrekturansicht)
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
            value = row.value(column.index)
            record[_column_title(column)] = value
        records.append(record)
    return pd.DataFrame(records, columns=ROW_FIELDS + [_column_title(c) for c in columns])


def _column_title(column) -> str:
    prefix = column.address or str(column.index)
    label = column.label or (column.column_key or "")
    return f"{prefix} · {label}"[:60]


def frame_to_result(frame: pd.DataFrame, result: DocumentResult) -> DocumentResult:
    """Uebernimmt Korrekturen aus der Tabelle zurueck in das Ergebnisobjekt."""
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
        cells: list[Cell] = []
        for title, column_index in title_to_index.items():
            value = record.get(title)
            if value is None or (isinstance(value, float) and pd.isna(value)) or value == "":
                continue
            cells.append(Cell(column_index=column_index, raw_value=str(value), count=float(value)))
        row.cells = cells
    return result


# --------------------------------------------------------------------------
# Seitenleiste
# --------------------------------------------------------------------------

with st.sidebar:
    st.title("VDI 3814 DP-Checker")
    st.caption("Offline-Auswertung von GA-Funktionslisten")
    db_path = st.text_input("Datenbank", value=str(SETTINGS.db_path))
    use_model = st.checkbox("Vision-Modell für Scans/Bilder verwenden", value=True)
    model_name = st.text_input("Ollama-Modell", value=SETTINGS.vision_model, disabled=not use_model)
    backend, backend_note = make_backend(use_model, model_name)
    (st.success if backend else st.info)(backend_note)
    st.caption(f"OCR (Tesseract): {'verfügbar' if ocr.available() else 'nicht installiert'}")
    st.caption("Profile: " + ", ".join(f"{p.id} ({len(p.columns)} Sp.)" for p in load_profiles()))

engine = get_engine(db_path)

tab_import, tab_preview, tab_overview, tab_costs, tab_export, tab_db = st.tabs(
    ["1 · Import", "2 · Vorschau & Korrektur", "3 · Gesamtübersicht",
     "4 · Kostenschätzung", "5 · Export", "6 · Datenbank"]
)


# --------------------------------------------------------------------------
# 1 - Import
# --------------------------------------------------------------------------

with tab_import:
    st.subheader("Funktionslisten importieren")
    st.write("PDF (auch mehrseitig/gescannt), PNG, JPEG sowie die Excel-Vorlagen "
             "(.xlsx/.xls). Regelschemata in denselben Dateien werden erkannt und "
             "übersprungen.")
    uploads = st.file_uploader("Dateien auswählen", accept_multiple_files=True,
                               type=["pdf", "png", "jpg", "jpeg", "tif", "tiff", "xlsx", "xlsm", "xls"])
    if uploads and st.button("Erkennung starten", type="primary"):
        progress = st.progress(0.0)
        log = st.empty()
        with tempfile.TemporaryDirectory() as tmp:
            for index, upload in enumerate(uploads, start=1):
                path = Path(tmp) / upload.name
                path.write_bytes(upload.getbuffer())
                log.info(f"({index}/{len(uploads)}) {upload.name} …")
                result = process_file(path, backend=backend,
                                      progress=lambda message: log.info(message))
                result.source_path = upload.name
                state()["results"][upload.name] = result
                progress.progress(index / len(uploads))
        log.empty()
        st.success(f"{len(uploads)} Datei(en) verarbeitet. Weiter mit „Vorschau & Korrektur“.")

    if state()["results"]:
        st.divider()
        rows = []
        for name, result in state()["results"].items():
            rows.append({
                "Datei": name,
                "Fassung": result.fassung or "?",
                "Profil": result.profile_id or "-",
                "Verfahren": result.backend,
                "Spalten": len(result.columns),
                "Datenpunkte": len(result.data_rows()),
                "Funktionen": result.grand_total(),
                "Übersprungen": ", ".join(f"S.{p.page_index + 1} ({p.classification.kind.value})"
                                          for p in result.skipped_pages()),
                "Hinweise": len(result.warnings),
            })
        st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
        for name, result in state()["results"].items():
            for warning in result.warnings:
                st.warning(f"{name}: {warning}")


# --------------------------------------------------------------------------
# 2 - Vorschau & Korrektur
# --------------------------------------------------------------------------

with tab_preview:
    st.subheader("Vorschau und manuelle Korrektur")
    results = state()["results"]
    if not results:
        st.info("Noch nichts importiert.")
    else:
        name = st.selectbox("Datei", list(results.keys()))
        result = results[name]

        with st.expander("Kopf-/Fußdaten des Dokuments", expanded=True):
            columns = st.columns(4)
            fields = list(vars(result.metadata))
            for index, field_name in enumerate(fields):
                with columns[index % 4]:
                    setattr(result.metadata, field_name,
                            st.text_input(field_name, value=getattr(result.metadata, field_name),
                                          key=f"meta_{name}_{field_name}"))

        st.caption("Zahlenwerte je Funktionsspalte – Änderungen wirken sich sofort auf "
                   "Summen, Kosten und Export aus. Spaltentitel: Abschnitt.Spalte · Bezeichnung")
        frame = result_to_frame(result)
        edited = st.data_editor(frame, width="stretch", num_rows="fixed",
                                key=f"editor_{name}", height=420)
        frame_to_result(edited, result)

        left, right = st.columns(2)
        with left:
            st.metric("Funktionen (eigene Summe)", f"{result.grand_total():g}")
        with right:
            mismatches = [c for c in result.sum_checks if not c.matches]
            st.metric("Abweichungen zur Summenzeile", len(mismatches))
        if result.sum_checks:
            check_frame = pd.DataFrame([
                {
                    "Spalte": (result.column_by_index(c.column_index).display()
                               if result.column_by_index(c.column_index) else c.column_index),
                    "Dokument": c.reported,
                    "Eigene Summe": c.computed,
                    "OK": "✓" if c.matches else "✗",
                }
                for c in result.sum_checks
            ])
            with st.expander("Prüfung gegen die Zeile „Summe Funktionen“"):
                st.dataframe(check_frame, width="stretch", hide_index=True)
        if result.footnotes:
            with st.expander("Fußnoten / Zählregeln"):
                st.dataframe(pd.DataFrame([
                    {"Marker": f.marker, "Text": f.text,
                     "Betrifft Spalten": ", ".join(f.referenced_columns)}
                    for f in result.footnotes
                ]), width="stretch", hide_index=True)

        if st.button("In Datenbank speichern", type="primary", key=f"save_{name}"):
            with Session(engine) as session:
                document = db.save_document(session, result)
            st.success(f"Gespeichert als Dokument {document.id}.")


# --------------------------------------------------------------------------
# 3 - Gesamtuebersicht
# --------------------------------------------------------------------------

def load_raw() -> pd.DataFrame:
    with Session(engine) as session:
        return aggregate.raw_dataframe(session)


with tab_overview:
    st.subheader("Gesamtübersicht über alle importierten Listen")
    raw = load_raw()
    if raw.empty:
        st.info("Noch keine Daten in der Datenbank.")
    else:
        filters = st.columns(4)
        with filters[0]:
            fassung = st.multiselect("Fassung", sorted(raw["fassung"].unique()))
        with filters[1]:
            projekt = st.multiselect("Projekt", sorted(raw["projekt"].fillna("").unique()))
        with filters[2]:
            anlage = st.multiselect("Anlage", sorted(raw["anlage"].fillna("").unique()))
        with filters[3]:
            gewerk = st.multiselect("Gewerk", sorted(raw["gewerk"].fillna("").unique()))
        filtered = raw
        for column, values in (("fassung", fassung), ("projekt", projekt),
                               ("anlage", anlage), ("gewerk", gewerk)):
            if values:
                filtered = filtered[filtered[column].isin(values)]

        metrics = st.columns(4)
        metrics[0].metric("Dateien", filtered["dokument_id"].nunique())
        metrics[1].metric("Datenpunkte", filtered["datenpunkt"].nunique())
        metrics[2].metric("Funktionen gesamt", f"{filtered['wert'].sum():g}")
        metrics[3].metric("Funktionsspalten belegt", filtered["spalte_key"].nunique())

        st.markdown("**Summen je Funktionsspalte**")
        st.dataframe(aggregate.column_summary(filtered), width="stretch", hide_index=True)
        left, right = st.columns(2)
        with left:
            st.markdown("**Je Spaltengruppe**")
            st.dataframe(aggregate.group_summary(filtered), width="stretch", hide_index=True)
        with right:
            dimension = st.selectbox("Aufschlüsselung nach", ["projekt", "anlage", "gewerk", "datei"])
            st.dataframe(aggregate.by_dimension(filtered, dimension),
                         width="stretch", hide_index=True)
        with st.expander("Drilldown auf Einzelwerte"):
            st.dataframe(filtered, width="stretch", hide_index=True)


# --------------------------------------------------------------------------
# 4 - Kostenschaetzung
# --------------------------------------------------------------------------

with tab_costs:
    st.subheader("Kostenschätzung")
    st.write("Einheitspreise können jederzeit gepflegt werden – auch bevor Daten "
             "importiert wurden. Die Kosten werden sofort neu berechnet.")
    with Session(engine) as session:
        prices = db.get_unit_prices(session)
    profile_id = st.selectbox("Profil", [p.id for p in load_profiles()])
    template = price_template(profile_id, prices)
    edited_prices = st.data_editor(
        template, width="stretch", hide_index=True, num_rows="fixed",
        column_config={"einheitspreis": st.column_config.NumberColumn(
            f"Einheitspreis [{SETTINGS.currency}]", min_value=0.0, step=1.0, format="%.2f")},
        disabled=["spalte_key", "spalte_adresse", "gruppe", "untergruppe", "spalte"],
        key=f"prices_{profile_id}", height=420,
    )
    if st.button("Einheitspreise speichern", type="primary"):
        with Session(engine) as session:
            db.set_unit_prices(session,
                               dict(zip(edited_prices["spalte_key"], edited_prices["einheitspreis"])),
                               profile_id=profile_id)
        st.success("Einheitspreise gespeichert.")

    raw = load_raw()
    if not raw.empty:
        current = dict(zip(edited_prices["spalte_key"], edited_prices["einheitspreis"]))
        current = {**prices, **current}
        priced = apply_prices(aggregate.column_summary(raw), current)
        st.metric("Geschätzte Gesamtkosten",
                  f"{total_cost(priced):,.2f} {SETTINGS.currency}".replace(",", "."))
        st.dataframe(priced[["spalte_adresse", "gruppe", "spalte", "menge",
                             "einheitspreis", "kosten"]],
                     width="stretch", hide_index=True)


# --------------------------------------------------------------------------
# 5 - Export
# --------------------------------------------------------------------------

with tab_export:
    st.subheader("Excel-Export")
    raw = load_raw()
    if raw.empty:
        st.info("Noch keine Daten in der Datenbank.")
    else:
        target = st.text_input("Zieldatei", value=str(Path("out") / "VDI3814_Auswertung.xlsx"))
        if st.button("Excel erzeugen", type="primary"):
            with Session(engine) as session:
                path = export_excel.export_workbook(
                    target,
                    summary=aggregate.column_summary(raw),
                    raw=raw,
                    documents=aggregate.documents_frame(session),
                    footnotes=aggregate.footnotes_frame(session),
                    prices=db.get_unit_prices(session),
                    projects=aggregate.pivot_projects(raw),
                )
            st.success(f"Geschrieben: {path}")
            st.download_button("Datei herunterladen", data=Path(path).read_bytes(),
                               file_name=Path(path).name,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# --------------------------------------------------------------------------
# 6 - Datenbank
# --------------------------------------------------------------------------

with tab_db:
    st.subheader("Datenbank")
    with Session(engine) as session:
        documents = aggregate.documents_frame(session)
    if documents.empty:
        st.info("Datenbank ist leer.")
    else:
        st.dataframe(documents, width="stretch", hide_index=True)
        target_id = st.number_input("Dokument-ID löschen", min_value=0, step=1, value=0)
        if target_id and st.button("Löschen", type="secondary"):
            with Session(engine) as session:
                db.delete_document(session, int(target_id))
            st.success(f"Dokument {int(target_id)} gelöscht.")
            st.rerun()
