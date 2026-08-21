"""Automations- (ASP) und Informationsschwerpunkt (ISP) erkennen und zuordnen.

In den VDI-Funktionslisten steht praktisch immer irgendwo, zu welchem
Schwerpunkt eine Seite oder eine Zeile gehoert:

* als Kopf-/Fussfeld ("ASP:", "Informationsschwerpunkt: ISP 02"),
* als Bestandteil der Anlagenbezeichnung ("Anlage: ASP01 Heizung"),
* auf Summen- und Uebersichtsblaettern als eigene Zeile, dort meist zweizeilig:
  "ASP01" und darunter "Heizungstechnik 2.UG".

Damit die Auswertung sagen kann, wie viele Datenpunkte und Funktionen auf
welchen ASP bzw. ISP entfallen, bekommt jede Zeile hier ihre Kennung
("ASP01") und die zugehoerige Bezeichnung ("Heizungstechnik 2.UG").

Die Erkennung ist bewusst streng: nur ein alleinstehendes ASP/ISP bzw. das
ausgeschriebene Wort gilt als Treffer - "Aspiration" oder "Wasp" nicht.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

# Anzeige fuer Zeilen, denen sich kein Schwerpunkt zuordnen laesst.
OHNE_ZUORDNUNG = "ohne Zuordnung"

_TOKEN = r"asp|isp|automationsschwerpunkt|informationsschwerpunkt"
_TRENNER = r"[\s.:;,_\-–/#]*"
_MUSTER = re.compile(
    rf"(?<![A-Za-zÄÖÜäöüß0-9])({_TOKEN})"
    rf"{_TRENNER}(?:nr\.?|no\.?)?{_TRENNER}(\d{{1,3}})?(?![A-Za-z0-9])",
    re.IGNORECASE,
)
_RAENDER = " \t:;,.-–—/|*()[]"


@dataclass(frozen=True)
class Schwerpunkt:
    """Ein Automations- oder Informationsschwerpunkt."""

    kennung: str            # normiert, z. B. "ASP01" (ohne Nummer: "ASP")
    bezeichnung: str = ""   # Klartext, z. B. "Heizungstechnik 2.UG"

    @property
    def art(self) -> str:
        return self.kennung[:3].upper()

    def anzeige(self) -> str:
        return f"{self.kennung} {self.bezeichnung}".strip()


def _saubere_bezeichnung(text: str | None, maximum: int = 60) -> str:
    """Rest eines Fundes als Bezeichnung aufbereiten."""
    if not text:
        return ""
    sauber = re.sub(r"\s+", " ", str(text)).strip(_RAENDER).strip()
    return sauber[:maximum].strip(_RAENDER).strip()


def parse(text: str | None, *, mit_bezeichnung: bool = True) -> Schwerpunkt | None:
    """Liest "ASP01 Heizungstechnik" o. Ae. aus einem beliebigen Text.

    Rueckgabe None, wenn der Text keinen Schwerpunkt nennt.
    """
    if not text:
        return None
    text = str(text)
    funde = list(_MUSTER.finditer(text))
    if not funde:
        return None
    # "Informationsschwerpunkt: ISP04" nennt die Kennung zweimal - die Angabe
    # mit laufender Nummer ist die gemeinte.
    treffer = next((f for f in funde if f.group(2)), funde[0])
    art = "ASP" if treffer.group(1)[0].lower() == "a" else "ISP"
    nummer = treffer.group(2)
    kennung = f"{art}{int(nummer):02d}" if nummer else art
    if not mit_bezeichnung:
        return Schwerpunkt(kennung)
    # Die Bezeichnung steht hinter der Kennung; nur wenn dort nichts mehr
    # folgt, wird der Text davor genommen - ohne die Feldbeschriftung
    # ("Anlage: ASP03 Lueftung" -> "Lueftung", "RLT 01 ASP03" -> "RLT 01").
    bezeichnung = _saubere_bezeichnung(text[treffer.end():])
    if not bezeichnung:
        davor = text[:treffer.start()]
        bezeichnung = _saubere_bezeichnung(davor.rpartition(":")[2])
    return Schwerpunkt(kennung, bezeichnung)


def aus_metadaten(metadata) -> Schwerpunkt | None:
    """Schwerpunkt aus dem Kopf-/Fussbereich eines Blattes.

    Reihenfolge: ausdrueckliches ASP-/ISP-Feld, dann die Anlagenbezeichnung -
    dort steht der Schwerpunkt in vielen Vorlagen mit drin.
    """
    felder = (
        ("ASP", getattr(metadata, "automationsschwerpunkt", "")),
        ("ISP", getattr(metadata, "informationsschwerpunkt", "")),
        ("", getattr(metadata, "anlage", "")),
        ("", getattr(metadata, "gewerk", "")),
    )
    for art, wert in felder:
        if not wert:
            continue
        gefunden = parse(wert)
        if gefunden is None and art:
            # Das Feld heisst bereits "ASP"/"Informationsschwerpunkt", im Wert
            # steht nur noch die Nummer bzw. der Klartext.
            gefunden = parse(f"{art} {wert}")
        if gefunden is None:
            continue
        if not gefunden.bezeichnung:
            # Steht im ASP-Feld nur die Nummer, ist die Anlagenbezeichnung des
            # Blattes die Bezeichnung des Schwerpunkts.
            ersatz = _saubere_bezeichnung(getattr(metadata, "anlage", ""))
            if ersatz:
                gefunden = Schwerpunkt(gefunden.kennung, ersatz)
        return gefunden
    return None


def _aus_zeile(row) -> Schwerpunkt | None:
    """Schwerpunkt, den eine Zeile selbst nennt.

    Nur die Beschriftung liefert auch eine Bezeichnung: steht das Kuerzel in
    der Benutzeradresse ("*ASP01*L01*TEMP"), ist der Rest der Adresse kein
    Klartext des Schwerpunkts, sondern der Name des Datenpunkts.
    """
    gefunden = parse(getattr(row, "klartext", ""))
    if gefunden is not None:
        return gefunden
    for feld in ("bas", "qualifier", "remark"):
        gefunden = parse(getattr(row, feld, ""), mit_bezeichnung=False)
        if gefunden is not None:
            return gefunden
    return None


def zuordnen(rows: list, vorgaben: dict[int, Schwerpunkt] | None = None,
             vorgabe: Schwerpunkt | None = None) -> dict[str, str]:
    """Traegt in jede Zeile den Schwerpunkt ein, zu dem sie gehoert.

    Regeln, in dieser Reihenfolge:

    1. Nennt die Zeile selbst einen Schwerpunkt, gilt dieser - und ab da auch
       fuer die folgenden Zeilen derselben Seite (Zwischenueberschriften).
    2. Sonst gilt der Schwerpunkt aus dem Kopf-/Fussbereich der Seite.
    3. Nennt eine Seite ohne Kopfangabe genau einen Schwerpunkt, gilt dieser
       fuer die ganze Seite - auch fuer die Zeilen davor.
    4. Sonst der Schwerpunkt des Dokuments; findet sich keiner, bleibt das
       Feld leer.

    Rueckgabe: je Kennung die gefundene Bezeichnung.
    """
    vorgaben = dict(vorgaben or {})
    eigene: list[Schwerpunkt | None] = [_aus_zeile(row) for row in rows]

    # Auf Uebersichtsblaettern steht die Kennung allein in der Zelle und die
    # Bezeichnung eine Zeile darunter ("ASP01" / "Heizungstechnik 2.UG").
    for index, gefunden in enumerate(eigene):
        if gefunden is None or gefunden.bezeichnung:
            continue
        for folge in range(index + 1, min(index + 3, len(rows))):
            if eigene[folge] is not None or rows[folge].has_values:
                break
            text = _saubere_bezeichnung(rows[folge].klartext)
            if text:
                eigene[index] = Schwerpunkt(gefunden.kennung, text)
                break

    # Regel 3: eine Seite mit genau einem genannten Schwerpunkt
    je_seite: dict[int, set[str]] = {}
    for row, gefunden in zip(rows, eigene):
        if gefunden is not None:
            je_seite.setdefault(row.page_index, set()).add(gefunden.kennung)
    for row, gefunden in zip(rows, eigene):
        if (gefunden is not None and row.page_index not in vorgaben
                and len(je_seite.get(row.page_index, ())) == 1):
            vorgaben[row.page_index] = gefunden

    bezeichnungen: dict[str, str] = {}
    aktuelle_seite: int | None = None
    seitenwert: Schwerpunkt | None = None
    aktuell: Schwerpunkt | None = None
    for row, gefunden in zip(rows, eigene):
        if row.page_index != aktuelle_seite:
            aktuelle_seite = row.page_index
            seitenwert = vorgaben.get(row.page_index) or vorgabe
            aktuell = seitenwert
        if getattr(row, "is_sum_row", False):
            # Eine Summen- oder Uebertragszeile schliesst den Block ab: sie
            # fasst die Zeilen darueber zusammen und gehoert damit nicht zum
            # zuletzt genannten Schwerpunkt, sondern zur Seite.
            aktuell = seitenwert
        elif gefunden is not None:
            aktuell = gefunden
        row.schwerpunkt = aktuell.kennung if aktuell else ""
        row.schwerpunkt_text = aktuell.bezeichnung if aktuell else ""
        if aktuell and aktuell.bezeichnung:
            bezeichnungen.setdefault(aktuell.kennung, aktuell.bezeichnung)

    # Wurde die Bezeichnung nur an einer Stelle genannt, gilt sie ueberall.
    for row in rows:
        if row.schwerpunkt and not row.schwerpunkt_text:
            row.schwerpunkt_text = bezeichnungen.get(row.schwerpunkt, "")
    return bezeichnungen


def je_seite(rows: list) -> dict[int, str]:
    """Welche Schwerpunkte stehen auf welcher Seite? (fuer die Seitenuebersicht)"""
    gefunden: dict[int, list[str]] = {}
    for row in rows:
        kennung = getattr(row, "schwerpunkt", "")
        if kennung and kennung not in gefunden.setdefault(row.page_index, []):
            gefunden[row.page_index].append(kennung)
    return {seite: ", ".join(sorted(kennungen)) for seite, kennungen in gefunden.items()}
