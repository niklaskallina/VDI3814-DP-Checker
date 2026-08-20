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
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .config import SETTINGS

HEADER_FILL = PatternFill("solid", fgColor="DDE5F0")
GROUP_FILL = PatternFill("solid", fgColor="C9D6E8")
THIN = Side(style="thin", color="9BA7B4")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
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
                    currency: str | None = None,
                    layouts: dict[str, pd.DataFrame] | None = None) -> Path:
    """Schreibt die vollstaendige Auswertung als .xlsx."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    prices = prices or {}
    currency = currency or SETTINGS.currency

    workbook = Workbook()
    workbook.remove(workbook.active)

    # Zuerst die Funktionslisten im Original-Layout - das ist die Sicht, in der
    # die Datenpunkte geprueft und kalkuliert werden.
    from .profiles_loader import get_profile

    for profile_id, frame in (layouts or {}).items():
        profile = get_profile(profile_id)
        if profile is None or frame is None or frame.empty:
            continue
        titel = f"GA-Funktionsliste {profile.fassung}"[:31]
        _sheet_vdi_layout(workbook, titel, frame, profile, prices, currency)

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


# --------------------------------------------------------------------------
# Blatt im Layout der GA-Funktionsliste
# --------------------------------------------------------------------------

ROW_FIELD_COUNT = 6      # Datei, Blatt, Zeile Nr., Datenpunkt, Benutzeradresse, Typ


def _sheet_vdi_layout(workbook: Workbook, title: str, frame: pd.DataFrame, profile,
                      prices: dict[str, float], currency: str) -> None:
    """Schreibt die Datenpunkte im Aufbau der Original-Funktionsliste.

    Kopfbereich wie in der VDI-Vorlage: Gruppe, Untergruppe, senkrechte
    Spaltenbezeichnung sowie die Nummernzeilen "Abschnitt" und "Spalte".
    Unten die Zeile "Summe Funktionen" als Formel, darunter eine Zeile fuer
    Einheitspreise und die daraus berechneten Kosten - so laesst sich direkt
    in der Liste kalkulieren.
    """
    sheet = workbook.create_sheet(title)
    columns = list(profile.columns)
    first_data_column = ROW_FIELD_COUNT + 1
    remark_column = first_data_column + len(columns)

    kopf_titel = ["Datei", "Blatt", "Zeile Nr.", "Datenpunkt", "Benutzeradresse", "Typ"]

    # --- Zeile 1-2: Gruppe und Untergruppe, jeweils zusammengefasst ---
    for zeile, attribut in ((1, "group"), (2, "subgroup")):
        start = first_data_column
        for index in range(len(columns) + 1):
            aktuell = getattr(columns[index], attribut) if index < len(columns) else None
            vorher = getattr(columns[index - 1], attribut) if index else None
            if index and aktuell != vorher:
                ende = first_data_column + index - 1
                cell = sheet.cell(row=zeile, column=start, value=vorher)
                cell.font = HEADER_FONT
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.fill = GROUP_FILL
                if ende > start:
                    sheet.merge_cells(start_row=zeile, start_column=start,
                                      end_row=zeile, end_column=ende)
                start = first_data_column + index

    # --- Zeile 3: Spaltenbezeichnung, senkrecht wie im Original ---
    for index, column in enumerate(columns):
        cell = sheet.cell(row=3, column=first_data_column + index, value=column.label)
        cell.alignment = Alignment(textRotation=90, horizontal="center", vertical="bottom")
        cell.font = Font(size=8)
        cell.border = BOX
    sheet.row_dimensions[3].height = 150

    # --- Zeile 4/5: Abschnitt und Spalte ---
    for zeile, beschriftung in ((4, "Abschnitt"), (5, "Spalte")):
        cell = sheet.cell(row=zeile, column=ROW_FIELD_COUNT, value=beschriftung)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="right")
    for index, column in enumerate(columns):
        abschnitt, _, nummer = column.address.rpartition(".")
        for zeile, wert in ((4, abschnitt or column.address), (5, nummer)):
            cell = sheet.cell(row=zeile, column=first_data_column + index, value=wert)
            cell.alignment = Alignment(horizontal="center")
            cell.font = Font(size=8, bold=True)
            cell.fill = HEADER_FILL
            cell.border = BOX

    # --- Beschriftung der Zeilenfelder: ueber die Kopfzeilen 1-3 zusammengefasst,
    #     damit die Nummernzeilen "Abschnitt"/"Spalte" frei bleiben (wie im Original) ---
    for index, titel in enumerate(kopf_titel):
        cell = sheet.cell(row=1, column=1 + index, value=titel)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        sheet.merge_cells(start_row=1, start_column=1 + index, end_row=3, end_column=1 + index)
    bemerkung = sheet.cell(row=1, column=remark_column, value="Bemerkungen")
    bemerkung.font = HEADER_FONT
    bemerkung.fill = HEADER_FILL
    bemerkung.alignment = Alignment(vertical="center", wrap_text=True)
    sheet.merge_cells(start_row=1, start_column=remark_column, end_row=3, end_column=remark_column)

    # --- Datenzeilen ---
    erste_datenzeile = 6
    zeile = erste_datenzeile
    for record in frame.to_dict("records"):
        for index, titel in enumerate(kopf_titel):
            sheet.cell(row=zeile, column=1 + index, value=record.get(titel, ""))
        for index, column in enumerate(columns):
            wert = record.get(column.label)
            if wert is None or (isinstance(wert, float) and pd.isna(wert)):
                continue
            zelle = sheet.cell(row=zeile, column=first_data_column + index, value=wert)
            zelle.alignment = Alignment(horizontal="center")
        sheet.cell(row=zeile, column=remark_column, value=record.get("Bemerkungen", ""))
        zeile += 1
    letzte_datenzeile = zeile - 1

    # --- Summenzeile als Formel ---
    summenzeile = zeile
    sheet.cell(row=summenzeile, column=1, value="Summe Funktionen").font = HEADER_FONT
    for index in range(len(columns)):
        spalte = first_data_column + index
        buchstabe = get_column_letter(spalte)
        formel = (f"=SUM({buchstabe}{erste_datenzeile}:{buchstabe}{letzte_datenzeile})"
                  if letzte_datenzeile >= erste_datenzeile else 0)
        zelle = sheet.cell(row=summenzeile, column=spalte, value=formel)
        zelle.font = HEADER_FONT
        zelle.fill = TOTAL_FILL
    gesamt_spalte = get_column_letter(remark_column)
    sheet.cell(row=summenzeile, column=remark_column,
               value=f"=SUM({get_column_letter(first_data_column)}{summenzeile}:"
                     f"{get_column_letter(remark_column - 1)}{summenzeile})").font = HEADER_FONT

    # --- Einheitspreise und Kosten ---
    preiszeile = summenzeile + 1
    kostenzeile = summenzeile + 2
    sheet.cell(row=preiszeile, column=1, value=f"Einheitspreis [{currency}]").font = HEADER_FONT
    sheet.cell(row=kostenzeile, column=1, value=f"Kosten [{currency}]").font = HEADER_FONT
    for index, column in enumerate(columns):
        spalte = first_data_column + index
        buchstabe = get_column_letter(spalte)
        preis = sheet.cell(row=preiszeile, column=spalte,
                           value=float(prices.get(column.key, 0.0)))
        preis.number_format = "#,##0.00"
        kosten = sheet.cell(row=kostenzeile, column=spalte,
                            value=f"={buchstabe}{summenzeile}*{buchstabe}{preiszeile}")
        kosten.number_format = "#,##0.00"
    gesamt = sheet.cell(
        row=kostenzeile, column=remark_column,
        value=f"=SUM({get_column_letter(first_data_column)}{kostenzeile}:"
              f"{get_column_letter(remark_column - 1)}{kostenzeile})")
    gesamt.font = HEADER_FONT
    gesamt.number_format = "#,##0.00"
    sheet.cell(row=kostenzeile - 1, column=remark_column).fill = TOTAL_FILL
    for spalte in range(1, remark_column + 1):
        sheet.cell(row=summenzeile, column=spalte).fill = TOTAL_FILL

    # --- Darstellung ---
    sheet.freeze_panes = sheet.cell(row=erste_datenzeile, column=first_data_column)
    for index in range(len(columns)):
        sheet.column_dimensions[get_column_letter(first_data_column + index)].width = 4.5
    for index, breite in enumerate((26, 7, 8, 34, 22, 12), start=1):
        sheet.column_dimensions[get_column_letter(index)].width = breite
    sheet.column_dimensions[gesamt_spalte].width = 40
