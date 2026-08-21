"""Korrekturtabelle der Oberflaeche: Ergebnis <-> DataFrame.

Bewusst ohne Streamlit-Abhaengigkeit. Die Umrechnung zwischen dem
Extraktionsergebnis und der bearbeitbaren Tabelle ist die einzige Stelle, an
der Korrekturen des Anwenders in das Datenmodell zurueckfliessen - sie laesst
sich hier unabhaengig von der Oberflaeche pruefen.
"""

from __future__ import annotations

import pandas as pd

from vdi3814.models import Cell, ColumnHeader, DocumentResult, RowKind

ROW_FIELDS = ["Zeile", "Datenpunkt", "Benutzeradresse", "Typ", "Bemerkung"]
DELETE_FIELD = "Löschen"   # Ankreuzspalte zum Entfernen ganzer Zeilen


def column_title(column: ColumnHeader) -> str:
    prefix = column.address or str(column.index)
    return f"{prefix} · {column.label or column.column_key or ''}"[:60]


def result_to_frame(result: DocumentResult) -> pd.DataFrame:
    columns = [c for c in result.columns if not c.is_note_column]
    records = []
    for row in result.rows:
        # Nur echte Datenpunkte zeigen. Summen-, Uebertrags-, Leer- und
        # Fussbereichszeilen stehen nachvollziehbar unter "Nicht gezaehlte
        # Zeilen" und haben in der Korrekturtabelle nichts verloren.
        if row.kind is not RowKind.DATEN:
            continue
        record = {
            DELETE_FIELD: False,
            "Zeile": row.row_no,
            "Datenpunkt": row.klartext,
            "Benutzeradresse": row.bas,
            "Typ": row.qualifier,
            "Bemerkung": row.remark,
        }
        for column in columns:
            record[column_title(column)] = row.value(column.index)
        records.append(record)
    spalten = [DELETE_FIELD] + ROW_FIELDS + [column_title(c) for c in columns]
    return pd.DataFrame(records, columns=spalten)


def frame_to_result(frame: pd.DataFrame, result: DocumentResult) -> DocumentResult:
    columns = [c for c in result.columns if not c.is_note_column]
    title_to_index = {column_title(c): c.index for c in columns}
    data_rows = [row for row in result.rows if row.kind is RowKind.DATEN]
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


def marked_rows(frame: pd.DataFrame) -> list[int]:
    """Positionen der zum Loeschen angekreuzten Zeilen.

    Die Position zaehlt die angezeigten Datenzeilen - also genau das, was
    DocumentResult.remove_data_rows erwartet.
    """
    if DELETE_FIELD not in frame.columns:
        return []
    return [position for position, wert in enumerate(frame[DELETE_FIELD].fillna(False))
            if bool(wert)]
