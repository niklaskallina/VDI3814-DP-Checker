"""Excel-Export (.xlsx) mit lebenden Formeln.

Blaetter:
    Übersicht       - Summen je Funktionsspalte ueber alle importierten Listen
    Kostenschätzung - Einheitspreis je Spalte, Menge per Formel aus "Übersicht",
                      Kosten = Menge * Einheitspreis, Gesamtsumme per SUM()
    Rohdaten        - jede Einzelzelle mit Quelldatei-, Seiten- und Zeilenbezug
    Dokumente       - importierte Dateien inkl. uebersprungener Seiten/Hinweise
    Fußnoten        - Zaehlregeln der Kopfzeile je Spalte
    Projekte        - Kreuztabelle Projekt/Anlage x Spaltengruppe

Die Kostenspalten sind bewusst Formeln und keine festen Werte, damit in Excel
mit geaenderten Einheitspreisen sofort weitergerechnet werden kann.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .config import SETTINGS

HEADER_FILL = PatternFill("solid", fgColor="DDE5F0")
TOTAL_FILL = PatternFill("solid", fgColor="F2F2F2")
HEADER_FONT = Font(bold=True)


def _write_frame(sheet: Worksheet, frame: pd.DataFrame, freeze: str = "A2") -> None:
    sheet.append([str(column) for column in frame.columns])
    for cell in sheet[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for record in frame.itertuples(index=False):
        sheet.append([_clean(value) for value in record])
    sheet.freeze_panes = freeze
    _autosize(sheet)


def _clean(value):
    if isinstance(value, datetime):
        return value.replace(microsecond=0)
    if isinstance(value, (list, tuple)):
        return ", ".join(str(item) for item in value)
    if pd.isna(value):
        return ""
    return value


def _autosize(sheet: Worksheet, maximum: int = 55) -> None:
    for column_cells in sheet.columns:
        letter = get_column_letter(column_cells[0].column)
        width = max((len(str(cell.value)) for cell in column_cells[:200] if cell.value is not None),
                    default=10)
        sheet.column_dimensions[letter].width = min(max(10, width + 2), maximum)


def _sheet_overview(workbook: Workbook, summary: pd.DataFrame) -> int:
    sheet = workbook.create_sheet("Übersicht")
    frame = summary.rename(columns={
        "spalte_adresse": "Abschnitt.Spalte",
        "spalte_key": "Spaltenschluessel",
        "gruppe": "Gruppe",
        "untergruppe": "Untergruppe",
        "spalte": "Funktionsspalte",
        "menge": "Menge",
        "dokumente": "Dateien",
        "datenpunkte": "Datenpunkte",
        "profil": "Profil",
        "fussnoten": "Zaehlregeln (Fussnoten)",
    })
    _write_frame(sheet, frame)

    last_row = sheet.max_row
    menge_col = list(frame.columns).index("Menge") + 1
    letter = get_column_letter(menge_col)
    total_row = last_row + 1
    sheet.cell(row=total_row, column=1, value="Summe Funktionen").font = HEADER_FONT
    total_cell = sheet.cell(row=total_row, column=menge_col,
                            value=f"=SUM({letter}2:{letter}{last_row})")
    total_cell.font = HEADER_FONT
    for column in range(1, sheet.max_column + 1):
        sheet.cell(row=total_row, column=column).fill = TOTAL_FILL
    return menge_col


def _sheet_costs(workbook: Workbook, summary: pd.DataFrame, prices: dict[str, float],
                 menge_col: int, currency: str) -> None:
    sheet = workbook.create_sheet("Kostenschätzung")
    headers = ["Abschnitt.Spalte", "Gruppe", "Untergruppe", "Funktionsspalte",
               "Spaltenschluessel", "Menge", f"Einheitspreis [{currency}]", f"Kosten [{currency}]"]
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    menge_letter = get_column_letter(menge_col)
    for offset, record in enumerate(summary.itertuples(index=False), start=2):
        # Menge als Formel auf das Uebersichtsblatt - aendert sich dort etwas,
        # rechnet die Kostenschaetzung automatisch mit.
        sheet.cell(row=offset, column=1, value=record.spalte_adresse)
        sheet.cell(row=offset, column=2, value=record.gruppe)
        sheet.cell(row=offset, column=3, value=record.untergruppe)
        sheet.cell(row=offset, column=4, value=record.spalte)
        sheet.cell(row=offset, column=5, value=record.spalte_key)
        sheet.cell(row=offset, column=6, value=f"='Übersicht'!{menge_letter}{offset}")
        price_cell = sheet.cell(row=offset, column=7, value=float(prices.get(record.spalte_key, 0.0)))
        price_cell.number_format = "#,##0.00"
        cost_cell = sheet.cell(row=offset, column=8, value=f"=F{offset}*G{offset}")
        cost_cell.number_format = "#,##0.00"

    last_row = sheet.max_row
    total_row = last_row + 1
    sheet.cell(row=total_row, column=1, value="Gesamtkosten").font = HEADER_FONT
    sheet.cell(row=total_row, column=6, value=f"=SUM(F2:F{last_row})").font = HEADER_FONT
    total = sheet.cell(row=total_row, column=8, value=f"=SUM(H2:H{last_row})")
    total.font = HEADER_FONT
    total.number_format = "#,##0.00"
    for column in range(1, 9):
        sheet.cell(row=total_row, column=column).fill = TOTAL_FILL
    sheet.freeze_panes = "A2"
    _autosize(sheet)


def export_workbook(path: str | Path,
                    summary: pd.DataFrame,
                    raw: pd.DataFrame,
                    documents: pd.DataFrame,
                    footnotes: pd.DataFrame,
                    prices: dict[str, float] | None = None,
                    projects: pd.DataFrame | None = None,
                    currency: str | None = None) -> Path:
    """Schreibt die vollstaendige Auswertung als .xlsx."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    prices = prices or {}
    currency = currency or SETTINGS.currency

    workbook = Workbook()
    workbook.remove(workbook.active)

    menge_col = _sheet_overview(workbook, summary)
    _sheet_costs(workbook, summary, prices, menge_col, currency)

    _write_frame(workbook.create_sheet("Rohdaten"), raw)
    _write_frame(workbook.create_sheet("Dokumente"), documents)
    _write_frame(workbook.create_sheet("Fußnoten"), footnotes)
    if projects is not None and not projects.empty:
        _write_frame(workbook.create_sheet("Projekte"), projects)

    info = workbook.create_sheet("Info")
    info.append(["VDI 3814 DP-Checker"])
    info.append(["Erstellt am", datetime.now().replace(microsecond=0)])
    info.append(["Dateien", int(documents.shape[0]) if not documents.empty else 0])
    info.append(["Datenpunkte", int(raw["datenpunkt"].nunique()) if not raw.empty else 0])
    info.append(["Funktionen gesamt", float(summary["menge"].sum()) if not summary.empty else 0.0])
    info.append([])
    info.append(["Hinweis", "Mengen und Kosten sind Formeln - Einheitspreise koennen "
                            "direkt in 'Kostenschaetzung' geaendert werden."])
    info["A1"].font = Font(bold=True, size=14)
    _autosize(info)

    workbook.save(path)
    return path
