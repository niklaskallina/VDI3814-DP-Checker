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
#
# Verglichen wird immer WORTWEISE. Ein Teilstring-Vergleich hat echte
# Datenpunkte verschluckt: "Summenstoerung", "Gesamtstoerung",
# "Totalausfall", "Waermeuebertrager" oder "Stoerung Uebertragung" sind
# Datenpunkte, keine Summen- oder Uebertragszeilen.
SUM_PATTERNS = (
    "summe", "summen", "zwischensumme", "gesamtsumme", "endsumme", "teilsumme",
    "summe funktionen", "anzahl funktionen",
)
# "Gesamt"/"Total" sind nur dann eine Summe, wenn danach nichts weiter steht
# als Summenvokabular, eine Zahl oder eine ASP/ISP-Kennung ("Gesamt",
# "Gesamt Funktionen", "Total ASP01") - "Gesamt Stoerung" ist ein Datenpunkt.
_GESAMT_WORTE = ("gesamt", "total")
_GESAMT_ERLAUBT = {"funktionen", "funktion", "summe", "anzahl", "seite", "blatt",
                   "liste", "alle", "aller", "der", "je", "dp", "datenpunkte"}
_KENNUNG_RE = re.compile(r"^(asp|isp)?\d+[a-z]?$")
CARRY_PATTERNS = (
    "uebertrag", "ubertrag", "vortrag", "seitensumme", "seitenuebertrag",
    "zwischenuebertrag",
)

# Beschriftungen des Kopf-/Fussbereichs - sie stehen ausserhalb der Tabelle.
# Eindeutig: so heisst kein Datenpunkt.
FOOTER_PATTERNS = (
    "ausgabedatum", "planersteller", "informationsschwerpunkt",
    "auftraggeber", "planstand", "blatt nr", "msr zeichnung nr", "zeichnungs nr",
    "zeichnungsnummer", "datenkommunikationsprotokoll", "aenderungsindex",
    "massstab", "plannummer", "definition der funktionen", "bibbs", "legende",
)
# Mehrdeutig: "Freigabe Kessel", "Index Sollwert" oder "Erstellt am" koennen
# Datenpunkte sein. Als Fussbereich gelten diese Woerter nur, wenn die Zeile
# keine Werte traegt UND die Beschriftung wie ein Feld des Schriftkopfs
# aussieht: das Wort allein oder mit Doppelpunkt/Punkt ("Freigabe:", "Rev.").
FOOTER_PATTERNS_MEHRDEUTIG = (
    "freigabe", "index", "revision", "rev", "erstellt", "bearbeiter", "geprueft",
    "gezeichnet", "anmerkung", "anmerkungen",
)
HEADER_PATTERNS = (
    "benennung", "datenpunkt", "benutzeradresse", "zeile nr", "abschnitt",
    "spalte", "bemerkung", "gewerk", "anlage", "funktionsliste",
    "ein ausgabefunktionen", "verarbeitungsfunktionen", "anwendungsfunktionen",
    "managementfunktionen", "bedienfunktionen", "b a funktionen",
)

# Strassen-/Ortszeilen des Planerstellers ("22941 Bargteheide", "Vosskuhlenweg 2")
_ADDRESS_RE = re.compile(r"^\s*\d{5}\s+[A-Za-zÄÖÜäöüß]")


def _beginnt_mit(norm: str, muster: str) -> bool:
    """True, wenn ``norm`` mit dem Muster als ganzem Wort (bzw. Wortfolge) beginnt."""
    return norm == muster or norm.startswith(muster + " ")


def _enthaelt_wort(norm: str, muster: str) -> bool:
    return muster in norm.split()


def _ist_gesamtzeile(norm: str) -> bool:
    worte = norm.split()
    if not worte or worte[0] not in _GESAMT_WORTE:
        return False
    return all(wort in _GESAMT_ERLAUBT or _KENNUNG_RE.match(wort) for wort in worte[1:])


def _sieht_aus_wie_kopffeld(text: str, muster: str) -> bool:
    """'Freigabe', 'Freigabe:', 'Rev.', 'Erstellt: 12.03.2024' - aber nicht
    'Freigabe Kessel' oder 'Index Sollwert'."""
    roh = (text or "").strip()
    if normalize(roh) == muster:
        return True
    treffer = re.match(r"^\s*([A-Za-zÄÖÜäöüß]+)\s*[:.]", roh)
    return bool(treffer) and normalize(treffer.group(1)) == muster


def classify_label(text: str, *, has_values: bool = False) -> RowKind:
    """Bestimmt die Art einer Zeile allein anhand ihrer Beschriftung."""
    norm = normalize(text)
    if not norm:
        return RowKind.LEER if not has_values else RowKind.DATEN

    for pattern in CARRY_PATTERNS:
        if _enthaelt_wort(norm, pattern):
            return RowKind.UEBERTRAG
    for pattern in SUM_PATTERNS:
        if _beginnt_mit(norm, pattern):
            return RowKind.SUMME
    if _ist_gesamtzeile(norm):
        return RowKind.SUMME
    for pattern in FOOTER_PATTERNS:
        if _beginnt_mit(norm, pattern):
            return RowKind.FUSSBEREICH
    if not has_values:
        for pattern in FOOTER_PATTERNS_MEHRDEUTIG:
            if _beginnt_mit(norm, pattern) and _sieht_aus_wie_kopffeld(text, pattern):
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
        if ohne_beschriftung and gezaehlte_zeilen >= 1:
            passt = _entspricht_summe(werte, laufend)
            # Die letzte Wertezeile einer Seite ohne Beschriftung ist die
            # Summenzeile - auch dann, wenn die Werte nicht exakt aufgehen.
            # Sie wird nie gezaehlt; eine Abweichung meldet stattdessen die
            # Summenpruefung. Das verhindert, dass eine unsauber gelesene
            # Summenzeile die Menge still verdoppelt.
            # Aber nur, wenn sie wie eine Summe AUSSIEHT: Werte in den meisten
            # bisher belegten Spalten. Ein letzter Datenpunkt, dessen Name im
            # Scan nicht lesbar war, traegt nur wenige Werte - er bleibt drin.
            if passt or (am_ende and gezaehlte_zeilen >= 2
                         and _sieht_aus_wie_summe(werte, laufend)):
                row.kind = RowKind.SUMME
                row.exclusion_reason = (
                    "Summenzeile ohne Beschriftung erkannt: die Werte entsprechen "
                    "spaltenweise der Summe der Zeilen darueber"
                    if passt else
                    "Unbeschriftete Schlusszeile - als Summenzeile gewertet, "
                    "die Werte gehen jedoch nicht auf (siehe Summenpruefung)"
                )
                laufend = {}
                gezaehlte_zeilen = 0
                continue
        for index_, wert in werte.items():
            laufend[index_] = laufend.get(index_, 0.0) + wert
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


def _sieht_aus_wie_summe(werte: dict[int, float], laufend: dict[int, float]) -> bool:
    """Deckt die Zeile die meisten belegten Spalten ab, ohne darunter zu liegen?"""
    belegte = {index for index, wert in laufend.items() if wert}
    if not belegte:
        return False
    abgedeckt = belegte & set(werte)
    if len(abgedeckt) < 0.6 * len(belegte):
        return False
    # Eine Summe ist nie kleiner als ein einzelner Summand der Spalte; liegt
    # die Zeile ueberwiegend UNTER den bisherigen Spaltensummen, ist sie eher
    # ein Datenpunkt.
    darunter = sum(1 for index in abgedeckt if werte[index] < laufend[index] * 0.5)
    return darunter <= len(abgedeckt) // 3
