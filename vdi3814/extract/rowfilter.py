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


def markiere_unbeschriftete_summenzeilen(rows: list) -> list:
    """Erkennt Summenzeilen, die keine Beschriftung tragen.

    In vielen Plansaetzen steht am Ende der Seite eine Summenzeile ohne den
    Text "Summe Funktionen" - nur die Zahlen. Ohne Erkennung wuerde sie als
    Datenpunkt gezaehlt und die Seite damit doppelt.

    Drei Bedingungen muessen zugleich erfuellt sein - sonst wuerden etwa zwei
    gleiche Zeilen hintereinander faelschlich als Summe gelten:

    1. Die Zeile traegt KEINE Beschriftung. Jeder echte Datenpunkt hat einen
       Namen; beschriftete Summenzeilen erkennt bereits die Schlagwortregel.
    2. Die Werte entsprechen SPALTENWEISE genau der Summe der Zeilen darueber,
       und zwar in jeder belegten Spalte.
    3. Die Zeile steht am Ende der Seite ODER es stehen mindestens zwei
       Datenzeilen darueber (Zwischensummen mehrerer Bloecke).
    """
    from ..models import RowKind

    letzte_mit_werten = -1
    for index, row in enumerate(rows):
        if row.has_values:
            letzte_mit_werten = index

    laufend: dict[int, float] = {}
    gezaehlte_zeilen = 0
    for index, row in enumerate(rows):
        if row.kind is RowKind.SUMME or row.kind is RowKind.UEBERTRAG:
            laufend = {}                    # danach beginnt ein neuer Block
            gezaehlte_zeilen = 0
            continue
        if row.kind is not RowKind.DATEN or not row.has_values:
            continue

        werte = {c.column_index: c.count for c in row.cells if c.count is not None}
        ohne_beschriftung = not normalize(row.klartext) and not normalize(row.bas)
        am_ende = index == letzte_mit_werten
        if (ohne_beschriftung and gezaehlte_zeilen >= 1
                and (am_ende or gezaehlte_zeilen >= 2)
                and _entspricht_summe(werte, laufend)):
            row.kind = RowKind.SUMME
            row.exclusion_reason = (
                "Summenzeile ohne Beschriftung erkannt: die Werte entsprechen "
                "spaltenweise der Summe der Zeilen darueber"
            )
            laufend = {}
            gezaehlte_zeilen = 0
            continue
        for index, wert in werte.items():
            laufend[index] = laufend.get(index, 0.0) + wert
        gezaehlte_zeilen += 1
    return rows


def _entspricht_summe(werte: dict[int, float], laufend: dict[int, float]) -> bool:
    """True, wenn die Zeilenwerte genau der bisherigen Spaltensumme entsprechen."""
    belegte = {index: wert for index, wert in laufend.items() if wert}
    if len(belegte) < 2 or len(werte) < 2:
        return False                        # zu wenig Belege, kein sicherer Schluss
    if set(werte) != set(belegte):
        return False
    return all(abs(werte[index] - belegte[index]) < 1e-6 for index in belegte)
