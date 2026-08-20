"""Erkennung von Nicht-Datenzeilen.

Fuer die Kalkulation zaehlt nur, was tatsaechlich ein Datenpunkt ist.
Alles andere - Summen-, Zwischensummen- und Uebertragszeilen, wiederholte
Tabellenkoepfe, Zeilen des Fussbereichs und leere Zeilen - darf die Auswertung
nicht beeinflussen. Diese Regeln gelten fuer alle Extraktionswege gleich.
"""

from __future__ import annotations

import re

from ..models import RowKind
from ..textutil import normalize

# Zeilen, deren Werte bereits an anderer Stelle gezaehlt wurden.
SUM_PATTERNS = (
    "summe", "summen", "zwischensumme", "gesamtsumme", "endsumme",
    "summe funktionen", "anzahl funktionen", "total", "gesamt",
)
CARRY_PATTERNS = (
    "uebertrag", "ubertrag", "vortrag", "uebertrag von", "uebertrag auf",
    "seitensumme", "seitenuebertrag", "zwischenuebertrag",
)

# Beschriftungen des Kopf-/Fussbereichs - sie stehen ausserhalb der Tabelle.
FOOTER_PATTERNS = (
    "ausgabedatum", "planersteller", "informationsschwerpunkt", "geprueft",
    "auftraggeber", "planstand", "blatt nr", "msr zeichnung nr", "zeichnungs nr",
    "datenkommunikationsprotokoll", "aenderungsindex", "index", "revision",
    "rev", "bearbeiter", "erstellt", "freigabe", "massstab", "plannummer",
    "anmerkung", "definition der funktionen", "bibbs", "legende",
)
HEADER_PATTERNS = (
    "benennung", "datenpunkt", "benutzeradresse", "zeile nr", "abschnitt",
    "spalte", "bemerkung", "gewerk", "anlage", "funktionsliste",
    "ein ausgabefunktionen", "verarbeitungsfunktionen", "anwendungsfunktionen",
    "managementfunktionen", "bedienfunktionen", "b a funktionen",
)

# Strassen-/Ortszeilen des Planerstellers ("22941 Bargteheide", "Vosskuhlenweg 2")
_ADDRESS_RE = re.compile(r"^\s*\d{4,5}\s+[A-Za-zÄÖÜäöüß]")


def classify_label(text: str, *, has_values: bool = False) -> RowKind:
    """Bestimmt die Art einer Zeile allein anhand ihrer Beschriftung."""
    norm = normalize(text)
    if not norm:
        return RowKind.LEER if not has_values else RowKind.DATEN

    for pattern in CARRY_PATTERNS:
        if norm.startswith(pattern) or f" {pattern}" in norm:
            return RowKind.UEBERTRAG
    for pattern in SUM_PATTERNS:
        if norm.startswith(pattern):
            return RowKind.SUMME
    for pattern in FOOTER_PATTERNS:
        if norm.startswith(pattern):
            return RowKind.FUSSBEREICH
    if _ADDRESS_RE.match(text or ""):
        return RowKind.FUSSBEREICH
    for pattern in HEADER_PATTERNS:
        if norm == pattern:
            return RowKind.KOPF
    return RowKind.DATEN


def is_footer_label(text: str) -> bool:
    """True, wenn die Beschriftung eindeutig zum Fussbereich gehoert."""
    return classify_label(text) is RowKind.FUSSBEREICH


def classify_row(label: str, has_values: bool, remark: str = "") -> tuple[RowKind, str]:
    """Art der Zeile plus Begruendung (fuer den Pruefbericht)."""
    kind = classify_label(label, has_values=has_values)
    if kind is RowKind.SUMME:
        return kind, "Summenzeile - Werte sind bereits in den Einzelzeilen enthalten"
    if kind is RowKind.UEBERTRAG:
        return kind, "Uebertragszeile - Werte stammen von einem anderen Blatt"
    if kind is RowKind.FUSSBEREICH:
        return kind, "Zeile des Kopf-/Fussbereichs, kein Datenpunkt"
    if kind is RowKind.KOPF:
        return kind, "Wiederholte Kopfzeile"
    if not has_values:
        if not normalize(label):
            return RowKind.LEER, "Leere Zeile"
        return RowKind.DATEN, ""
    return RowKind.DATEN, ""
