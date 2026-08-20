"""Extraktion aus Scans und Bildern - ohne KI, nur mit lokalem OCR.

Gescannte Funktionslisten haben keine Textebene. Statt ein Sprachmodell raten
zu lassen, wird die Seite mit Tesseract gelesen: das liefert Woerter MIT
Koordinaten, und damit greift genau dieselbe geometrische Rekonstruktion wie
bei digitalen PDFs (wordgrid.py).

Zusaetzlich werden die Tabellenlinien im Bild gesucht, um den Datenbereich
ebenso hart zu begrenzen wie bei einem PDF mit gezeichneten Spaltentrennern.
"""

from __future__ import annotations

import logging
import re
from collections import Counter
from dataclasses import dataclass

import numpy as np
from PIL import Image

from ..config import SETTINGS
from ..ingest.preprocess import prepare_for_ocr
from ..textutil import normalize
from .base import ExtractionMode, RawTable
from ..models import ColumnHeader
from .wordgrid import _HeaderGrid, Word, extract_table_from_words

log = logging.getLogger(__name__)

MIN_CONFIDENCE = 30.0        # Tesseract-Sicherheit in Prozent


def available() -> bool:
    from ..vision import ocr

    return ocr.available()


_INT_RE = re.compile(r"^\d{1,3}$")


def _durchgang(image: Image.Image, lang: str, config: str, start_index: int) -> list[Word]:
    import pytesseract

    daten = pytesseract.image_to_data(
        image, lang=lang, config=config, output_type=pytesseract.Output.DICT,
    )
    words: list[Word] = []
    for index, text in enumerate(daten["text"]):
        text = (text or "").strip()
        if not text:
            continue
        try:
            sicherheit = float(daten["conf"][index])
        except (TypeError, ValueError):
            sicherheit = -1.0
        if sicherheit < MIN_CONFIDENCE:
            continue
        x, y = float(daten["left"][index]), float(daten["top"][index])
        breite, hoehe = float(daten["width"][index]), float(daten["height"][index])
        words.append((x, y, x + breite, y + hoehe, text, 0, 0, start_index + index))
    return words


def _ueberlappt(a: Word, b: Word, anteil: float = 0.5) -> bool:
    breite = min(a[2], b[2]) - max(a[0], b[0])
    hoehe = min(a[3], b[3]) - max(a[1], b[1])
    if breite <= 0 or hoehe <= 0:
        return False
    schnitt = breite * hoehe
    kleinste = min((a[2] - a[0]) * (a[3] - a[1]), (b[2] - b[0]) * (b[3] - b[1])) or 1.0
    return schnitt / kleinste >= anteil


def ocr_words(image: Image.Image, lang: str | None = None) -> tuple[list[Word], Image.Image]:
    """Liest die Seite und liefert Woerter mit Koordinaten (in Bildpixeln).

    Zwei Durchgaenge, weil eine Funktionsliste zwei sehr verschiedene Inhalte
    hat:

    * zusammenhaengender Text (Beschriftungen, Kopf-/Fussbereich)  -> psm 6
    * einzeln stehende Ziffern in schmalen Zellen                  -> psm 11

    Ein Durchgang allein verliert jeweils den anderen Teil: mit psm 6 fehlen
    die meisten Zellenwerte, mit psm 11 die Textzeilen. Die Ergebnisse werden
    ueber ihre Position zusammengefuehrt, Doppelfunde entfallen.
    """
    sprache = lang or SETTINGS.ocr_lang
    aufbereitet = prepare_for_ocr(image)

    text_woerter = _durchgang(aufbereitet, sprache, "--psm 6", 0)
    ziffern = _durchgang(
        aufbereitet, sprache,
        "--psm 11 -c tessedit_char_whitelist=0123456789xX",
        100000,
    )

    words = list(text_woerter)
    for kandidat in ziffern:
        if not _INT_RE.match(kandidat[4]) and kandidat[4].lower() != "x":
            continue
        if any(_ueberlappt(kandidat, vorhanden) for vorhanden in text_woerter):
            continue
        words.append(kandidat)
    words.sort(key=lambda w: (w[1], w[0]))
    return words, aufbereitet


def _line_positions(maske: np.ndarray, achse: int, mindestlaenge: int) -> list[tuple[int, int, int]]:
    """Findet zusammenhaengende dunkle Strecken entlang einer Achse."""
    treffer: list[tuple[int, int, int]] = []
    anzahl = maske.shape[1] if achse == 0 else maske.shape[0]
    for position in range(anzahl):
        spur = maske[:, position] if achse == 0 else maske[position, :]
        start = None
        for index, gesetzt in enumerate(spur):
            if gesetzt and start is None:
                start = index
            elif not gesetzt and start is not None:
                if index - start >= mindestlaenge:
                    treffer.append((position, start, index))
                start = None
        if start is not None and len(spur) - start >= mindestlaenge:
            treffer.append((position, start, len(spur)))
    return treffer


def make_table_bottom_finder(image: Image.Image):
    """Unterkante des Datenbereichs aus den Tabellenlinien des Scans.

    Gesucht werden senkrechte dunkle Strecken - im Scan sind das die
    Spaltentrenner. Damit gilt fuer Scans dieselbe harte Grenze zwischen
    Tabelle und Fussbereich wie bei digitalen PDFs.
    """
    grau = np.asarray(image.convert("L"), dtype=np.uint8)
    if grau.size == 0:
        return lambda *_: None
    schwelle = max(60, int(grau.mean() * 0.6))
    dunkel = grau < schwelle
    mindestlaenge = max(20, int(grau.shape[0] * 0.08))
    senkrecht = _line_positions(dunkel, achse=0, mindestlaenge=mindestlaenge)

    def finder(body_top: float, x_left: float, x_right: float,
               pitch: float, row_pitch: float) -> float | None:
        if not senkrecht:
            return None
        start_tolerance = max(pitch * 0.8, row_pitch * 1.2, 6.0)
        links, rechts = x_left - pitch, x_right + pitch
        unterkanten = [
            ende for x, start, ende in senkrecht
            if links <= x <= rechts
            and abs(start - body_top) <= start_tolerance
            and ende > body_top + start_tolerance
        ]
        if not unterkanten:
            return None
        unterkanten.sort()
        # Median statt Maximum: einzelne durchgezogene Rahmenlinien sollen die
        # Grenze nicht nach unten ziehen.
        return float(unterkanten[len(unterkanten) // 2])

    return finder


def _waagerechte_linien(dunkel: np.ndarray, mindestbreite: int) -> list[int]:
    """y-Positionen der waagerechten Tabellenlinien."""
    treffer = _line_positions(dunkel, achse=1, mindestlaenge=mindestbreite)
    if not treffer:
        return []
    ys = sorted({y for y, _, _ in treffer})
    zusammengefasst: list[int] = []
    for y in ys:
        if zusammengefasst and y - zusammengefasst[-1] <= 3:
            continue
        zusammengefasst.append(y)
    return zusammengefasst


def _hat_inhalt(arr: np.ndarray, schwelle: float = 0.012) -> bool:
    """Schnelle Vorpruefung: steht in der Zelle ueberhaupt etwas?"""
    if arr.size == 0:
        return False
    return float((arr < 128).mean()) > schwelle


def _ist_linienrest(arr: np.ndarray) -> bool:
    """True, wenn die Schwaerzung nur vom Zellrahmen stammt.

    Reste der Tabellenlinien werden sonst gern als "1" oder "7" gelesen und
    erzeugen Werte, die in der Liste gar nicht stehen.
    """
    if arr.size == 0:
        return True
    dunkel = arr < 128
    spalten = dunkel.mean(axis=0)
    zeilen = dunkel.mean(axis=1)
    rand = max(1, arr.shape[1] // 5)
    innen = spalten[rand:-rand] if spalten.size > 2 * rand else spalten
    if innen.size == 0 or innen.max() < 0.15:
        return True                       # nichts in der Zellmitte
    if zeilen.size and float(zeilen.mean()) > 0.75:
        return True                       # durchgehend dunkel: Linie, keine Ziffer
    return False


def _lies_ziffer(zelle: Image.Image, mindest_konfidenz: float = 55.0) -> str:
    """Liest eine einzelne Zelle und verwirft unsichere Treffer."""
    import pytesseract

    daten = pytesseract.image_to_data(
        zelle, config="--psm 10 -c tessedit_char_whitelist=0123456789xX",
        output_type=pytesseract.Output.DICT,
    )
    bester, beste_konfidenz = "", -1.0
    for index, text in enumerate(daten["text"]):
        text = (text or "").strip()
        if not text:
            continue
        try:
            konfidenz = float(daten["conf"][index])
        except (TypeError, ValueError):
            konfidenz = -1.0
        if konfidenz > beste_konfidenz:
            bester, beste_konfidenz = text, konfidenz
    return bester if beste_konfidenz >= mindest_konfidenz else ""


_LABEL_MUELL = re.compile(r"[|\\\[\]{}_~^`]+")


def _bereinige_beschriftung(text: str) -> str:
    """Entfernt Linienreste aus einer gelesenen Zeilenbeschriftung.

    Der Zuschnitt enthaelt die Zellrahmen; Tesseract macht daraus gern
    Klammern oder senkrechte Striche ("[Aussentemperatur", "6]Ventil").
    Die fuehrende Zeilennummer bleibt erhalten - sie wird spaeter abgetrennt.
    """
    text = _LABEL_MUELL.sub(" ", text or "")
    text = re.sub(r"\s+", " ", text).strip(" .,:;-")
    # Ein senkrechter Zellrahmen wird gern als einzelner Buchstabe gelesen
    # ("lAussentemperatur", "g FBH Kita"). Einzelne Buchstaben am Anfang
    # gehoeren nie zu einer Datenpunktbezeichnung - Ziffern dagegen schon,
    # das ist die Zeilennummer.
    # (Die fuehrende Zeilennummer bleibt dabei erhalten.)
    # Nur KLEINbuchstaben entfernen - Abkuerzungen wie "WT Heizung" oder
    # "STW" beginnen mit Grossbuchstaben und muessen erhalten bleiben.
    text = re.sub(r"^(\d{1,3}\s+)?[a-zäöü](?=[A-ZÄÖÜ])", r"\1", text)
    text = re.sub(r"^(\d{1,3}\s+)?[a-zäöü]\s+(?=\S)", r"\1", text)
    return text.strip()


def zellen_ocr(image: Image.Image, spalten_kanten: list[float], zeilen_kanten: list[int],
               label_rechts: float, lang: str) -> list[Word]:
    """Liest die Tabelle Zelle fuer Zelle.

    Bei gescannten Funktionslisten ist das der zuverlaessigste Weg: die Ziffern
    stehen einzeln in schmalen, umrandeten Zellen - dort scheitert die
    seitenweise Texterkennung regelmaessig, waehrend ein Zuschnitt auf die
    einzelne Zelle sehr sicher gelesen wird. Leere Zellen werden vorab ueber
    ihren Schwaerzungsgrad aussortiert, das haelt den Aufwand klein.

    Rueckgabe: Woerter mit Koordinaten - dasselbe Format wie bei einem PDF,
    sodass die weitere Auswertung identisch ablaeuft.
    """
    import pytesseract

    grau = image.convert("L")
    arr = np.asarray(grau, dtype=np.uint8)
    woerter: list[Word] = []
    laufende_nummer = 0

    for oben, unten in zip(zeilen_kanten, zeilen_kanten[1:]):
        if unten - oben < 8:
            continue

        # Beschriftung der Zeile (links der ersten Funktionsspalte). Bewusst
        # zeilenweise: ein Sammeldurchgang ueber die ganze Spalte liest die
        # Datenpunktnamen deutlich schlechter.
        # Rand grosszuegig wegschneiden und vergroessern: die Zellrahmen sonst
        # als Buchstaben gelesen ("2jwr Heizung" statt "2 WT Heizung").
        rand = max(2, int((unten - oben) * 0.12))
        bereich = arr[oben + rand:unten - rand, rand:int(label_rechts) - rand]
        if _hat_inhalt(bereich, schwelle=0.004):
            zuschnitt = grau.crop((rand, oben + rand, int(label_rechts) - rand, unten - rand))
            if zuschnitt.width > 10 and zuschnitt.height > 6:
                zuschnitt = zuschnitt.resize(
                    (zuschnitt.width * 3, zuschnitt.height * 3), Image.LANCZOS)
                text = _bereinige_beschriftung(pytesseract.image_to_string(
                    zuschnitt, lang=lang, config="--psm 7"))
                if text:
                    laufende_nummer += 1
                    woerter.append((5.0, float(oben), label_rechts, float(unten),
                                    text, 0, 0, laufende_nummer))

        # Zellen der Funktionsspalten
        for links, rechts in zip(spalten_kanten, spalten_kanten[1:]):
            if rechts - links < 6:
                continue
            ausschnitt = arr[oben + 3:unten - 2, int(links) + 3:int(rechts) - 2]
            if not _hat_inhalt(ausschnitt) or _ist_linienrest(ausschnitt):
                continue
            zelle = grau.crop((int(links) + 3, oben + 3, int(rechts) - 2, unten - 2))
            if zelle.width < 4 or zelle.height < 4:
                continue
            zelle = zelle.resize((zelle.width * 4, zelle.height * 4), Image.LANCZOS)
            text = _lies_ziffer(zelle)
            if not text:
                continue
            laufende_nummer += 1
            woerter.append((float(links) + 2, float(oben) + 2, float(rechts) - 1, float(unten) - 1,
                            text, 0, 0, laufende_nummer))
    return woerter


def _grid_aus_linien(raster: "ScanRaster"):
    """Baut aus dem vermessenen Raster ein fertiges Kopf-Raster fuer wordgrid."""
    mitten = raster.spalten_mitten
    profil = raster.profil
    body_top, body_bottom = raster.body_top, raster.body_bottom
    abstand = (max(mitten) - min(mitten)) / max(1, len(mitten) - 1)

    spalten = []
    for index, (mitte, profilspalte) in enumerate(zip(mitten, profil.columns)):
        abschnitt, _, nummer = profilspalte.address.rpartition(".")
        spalten.append(ColumnHeader(
            index=index,
            group=profilspalte.group,
            subgroup=profilspalte.subgroup,
            label=profilspalte.label,
            abbrev=profilspalte.abbrev,
            section=abschnitt or profilspalte.address,
            column_no=nummer,
            x_center=mitte,
        ))
    grid = _HeaderGrid(
        columns=spalten,
        body_top=body_top,
        label_left=raster.label_rechts,
        note_left=max(mitten) + abstand,
        labels_top=body_top,
    )
    return grid


_ZELLWERT_RE = re.compile(r"^(\d{1,3}|[xX✓])$")


def _bereinigen(words: list[Word], label_left: float) -> list[Word]:
    """Entfernt OCR-Rauschen aus dem Zellbereich.

    Die Tabellenlinien werden von Tesseract gern als Buchstabenketten gelesen
    ("EEE", "IIl"). Rechts der Beschriftungsspalte kann eine Zelle aber nur
    eine Zahl oder ein Kreuz enthalten - alles andere ist Rauschen und wird
    verworfen. Links davon bleibt der Text erhalten, dort stehen die
    Datenpunktbezeichnungen.
    """
    bereinigt: list[Word] = []
    for x0, y0, x1, y1, text, a, b, c in words:
        text = text.replace("|", " ").replace("_", " ").strip(" .:;")
        if not text:
            continue
        if (x0 + x1) / 2 > label_left:
            if not _ZELLWERT_RE.match(text):
                continue
        elif len(text) == 1 and not text.isalnum():
            continue
        bereinigt.append((x0, y0, x1, y1, text, a, b, c))
    return bereinigt


def extract_ocr(image: Image.Image, page_index: int, lang: str | None = None) -> RawTable | None:
    """Liest eine gescannte Funktionsliste ueber lokales OCR aus.

    Zuerst wird versucht, das Spaltenraster aus den Tabellenlinien zu bauen -
    das ist bei Scans deutlich zuverlaessiger als der Versuch, die winzigen
    Kopfzeilen zu lesen. Gelingt das nicht, greift die normale Worterkennung.
    """
    if not available():
        return None
    sprache = lang or SETTINGS.ocr_lang
    aufbereitet = prepare_for_ocr(image)
    raster = spaltenraster_aus_linien(aufbereitet)

    if raster is not None and raster.zeilen_gefunden:
        # Bevorzugter Weg: das Raster ist vollstaendig vermessen, jede Zelle
        # wird einzeln gelesen. Bei umrandeten Tabellen ist das deutlich
        # genauer als seitenweises OCR.
        grid = _grid_aus_linien(raster)
        words = zellen_ocr(aufbereitet, raster.spalten_kanten, raster.zeilen_kanten,
                           raster.label_rechts, sprache)
        tabelle = extract_table_from_words(
            words, float(aufbereitet.width), page_index, ExtractionMode.OCR,
            grid=grid, table_bottom=raster.body_bottom,
        )
    else:
        # Tabelle ohne waagerechte Linien: Zeilen aus den erkannten Woertern
        words, aufbereitet = ocr_words(image, sprache)
        if len(words) < 20:
            return None
        grid = _grid_aus_linien(raster) if raster is not None else None
        if grid is not None:
            words = _bereinigen(words, grid.label_left)
        tabelle = extract_table_from_words(
            words, float(aufbereitet.width), page_index, ExtractionMode.OCR,
            table_bottom_finder=make_table_bottom_finder(aufbereitet),
            grid=grid,
            table_bottom=raster.body_bottom if raster is not None else None,
        )
    if tabelle is not None:
        tabelle.warnings.append(
            "Seite wurde per OCR gelesen - Werte bitte in der Vorschau stichprobenartig pruefen."
        )
    return tabelle


SCHEMA_KEYWORDS = ("regelschema", "prinzipschaltbild", "anlagenschema", "strangschema",
                   "funktionsschema", "schaltplan")
TABLE_KEYWORDS = ("summe funktionen", "funktionsliste", "datenpunktliste", "abschnitt",
                  "verarbeitungsfunktionen", "anwendungsfunktionen", "benennung")


def classify_image(image: Image.Image, lang: str | None = None) -> tuple[str, float, str]:
    """Seitenklassifikation ohne Modell, allein aus dem OCR-Text."""
    from ..vision import ocr

    if not available():
        return "unbekannt", 0.0, "Kein OCR verfuegbar"
    text = normalize(ocr.page_text(image, lang))
    if len(text) < 40:
        return "regelschema", 0.6, "Kaum Text auf der Seite - vermutlich eine Zeichnung"
    if any(normalize(k) in text for k in SCHEMA_KEYWORDS):
        return "regelschema", 0.85, "Schema-Stichwort im Text gefunden"
    treffer = [k for k in TABLE_KEYWORDS if normalize(k) in text]
    if treffer:
        return "funktionsliste", 0.8, f"Tabellen-Stichworte gefunden: {', '.join(treffer[:3])}"
    return "sonstiges", 0.5, "Weder Tabellen- noch Schema-Stichworte gefunden"


# --------------------------------------------------------------------------
# Spaltenraster aus den Tabellenlinien (fuer Scans, deren Kopfzeilen das OCR
# nicht sicher lesen kann)
# --------------------------------------------------------------------------

def _spalten_trenner(maske: np.ndarray, mindestanteil: float = 0.15) -> list[tuple[float, int, int]]:
    """Senkrechte Trennlinien: (x-Mitte, Oberkante, Unterkante)."""
    hoehe = maske.shape[0]
    mindestlaenge = max(20, int(hoehe * mindestanteil))
    treffer = _line_positions(maske, achse=0, mindestlaenge=mindestlaenge)
    if not treffer:
        return []
    treffer.sort()
    gruppen: list[list[tuple[int, int, int]]] = []
    for eintrag in treffer:
        if gruppen and eintrag[0] - gruppen[-1][-1][0] <= 3:
            gruppen[-1].append(eintrag)
        else:
            gruppen.append([eintrag])
    linien = []
    for gruppe in gruppen:
        x = sum(e[0] for e in gruppe) / len(gruppe)
        oben = min(e[1] for e in gruppe)
        unten = max(e[2] for e in gruppe)
        linien.append((x, oben, unten))
    return linien


def _zelle_lesen(band: Image.Image, links: float, rechts: float) -> int | None:
    """Liest eine einzelne Zahl aus einer schmalen Kopfzelle."""
    import pytesseract

    if rechts - links < 8:
        return None
    zelle = band.crop((int(links) + 2, 1, int(rechts) - 1, band.height - 1))
    if zelle.width < 4 or zelle.height < 4:
        return None
    zelle = zelle.resize((zelle.width * 4, zelle.height * 4), Image.LANCZOS)
    text = pytesseract.image_to_string(
        zelle, config="--psm 10 -c tessedit_char_whitelist=0123456789"
    ).strip()
    return int(text) if text.isdigit() and len(text) <= 2 else None


def _blockgroessen(werte: list[int | None]) -> list[int]:
    """Groessen der Abschnittsbloecke aus den gelesenen Spaltennummern.

    Eine "1" beginnt einen neuen Abschnitt. Genau diese Eins ist in Scans am
    haeufigsten unleserlich - sie wird daran erkannt, dass die naechste
    lesbare Zahl eine 2 ist.
    """
    bloecke: list[int] = []
    for index, wert in enumerate(werte):
        naechster = werte[index + 1] if index + 1 < len(werte) else None
        beginnt = wert == 1 or (wert is None and naechster == 2)
        if beginnt or not bloecke:
            bloecke.append(1)
        else:
            bloecke[-1] += 1
    return bloecke


@dataclass
class ScanRaster:
    """Vermessenes Tabellenraster eines Scans."""

    spalten_kanten: list[float]     # x-Grenzen der Funktionsspalten (n+1 Werte)
    profil: object                  # erkanntes Spalten-Profil
    body_top: float
    body_bottom: float
    zeilen_kanten: list[int]        # y-Grenzen der Datenzeilen
    label_rechts: float             # x-Grenze der Beschriftungsspalte
    zeilen_gefunden: bool = True    # False: Tabelle ohne waagerechte Linien

    @property
    def spalten_mitten(self) -> list[float]:
        return [(a + b) / 2 for a, b in zip(self.spalten_kanten, self.spalten_kanten[1:])]


def _pruefe_kopfband(band: Image.Image, zellen: list[tuple[float, float]]):
    """Liest ein Kopfband Zelle fuer Zelle und prueft es gegen die Profile.

    Rueckgabe: (Spaltenkanten, Profil, Guete, Anzahl gelesener Ziffern) oder None.
    Die Randspalten (Beschriftung links, Bemerkung rechts) werden probeweise
    abgeschnitten, bis die Blockstruktur zu einer hinterlegten Vorlage passt.
    """
    from ..profiles_loader import profile_by_block_sizes

    werte = [_zelle_lesen(band, a, b) for a, b in zellen]
    if sum(1 for wert in werte if wert is not None) < 6:
        return None

    anzahl = len(zellen)
    bester = None
    for start in range(0, min(8, anzahl)):
        for ende in range(anzahl, max(anzahl - 8, start + 4) - 1, -1):
            ausschnitt = werte[start:ende]
            if len(ausschnitt) < 8:
                continue
            profil = profile_by_block_sizes(_blockgroessen(ausschnitt))
            if profil is None:
                continue
            erwartet = [int(spalte.address.rpartition(".")[2]) for spalte in profil.columns]
            if len(erwartet) != len(ausschnitt):
                continue
            lesbar = [(gelesen, soll) for gelesen, soll in zip(ausschnitt, erwartet)
                      if gelesen is not None]
            if not lesbar:
                continue
            guete = sum(1 for gelesen, soll in lesbar if gelesen == soll) / len(lesbar)
            bewertung = (guete, len(lesbar), -start)
            if guete >= 0.7 and (bester is None or bewertung > bester[0]):
                kanten = [zellen[i][0] for i in range(start, ende)] + [zellen[ende - 1][1]]
                bester = (bewertung, kanten, profil, guete, len(lesbar))
    if bester is None:
        return None
    _, kanten, profil, guete, lesbar_anzahl = bester
    return kanten, profil, guete, lesbar_anzahl


def spaltenraster_aus_linien(image: Image.Image) -> "ScanRaster | None":
    """Vermisst das Tabellenraster eines Scans.

    Ablauf:
    1. senkrechte Trennlinien suchen -> Spaltengrenzen und Tabellenbereich
    2. waagerechte Linien suchen -> moegliche Kopf- und Zeilenbaender
    3. jedes Kopfband Zelle fuer Zelle lesen und pruefen, ob die Nummernfolge
       zur Blockstruktur einer hinterlegten Vorlage passt
       (alte Fassung 5-5-6-5-8-13-4-4, neue Fassung 1-4-2-5-2-12-9-5-4-5-1-6)

    Erst wenn dieser Abgleich aufgeht, gilt das Raster als erkannt - so wird
    nie geraten. Danach stehen alle Spaltenadressen fest, auch wenn einzelne
    Ziffern im Scan unleserlich waren.
    """
    grau = np.asarray(image.convert("L"), dtype=np.uint8)
    if grau.size == 0:
        return None
    dunkel = grau < max(90, int(grau.mean() * 0.75))
    linien = _spalten_trenner(dunkel)
    if len(linien) < 10:
        return None

    def _haeufigster(werte: list[int], raster: int = 4) -> float:
        gezaehlt = Counter(int(round(wert / raster)) * raster for wert in werte)
        return float(gezaehlt.most_common(1)[0][0])

    rahmen_oben = _haeufigster([l[1] for l in linien])
    rahmen_unten = _haeufigster([l[2] for l in linien])
    toleranz = max(6.0, grau.shape[0] * 0.02)
    linien = [l for l in linien
              if l[1] <= rahmen_oben + toleranz and l[2] >= rahmen_unten - toleranz]
    if len(linien) < 10 or (rahmen_unten - rahmen_oben) < grau.shape[0] * 0.15:
        return None

    xs = [l[0] for l in linien]
    abstaende = sorted(b - a for a, b in zip(xs, xs[1:]) if b - a > 4)
    if not abstaende:
        return None
    spaltenbreite = abstaende[len(abstaende) // 2]
    zellen = [(a, b) for a, b in zip(xs, xs[1:]) if b - a > spaltenbreite * 0.4]
    if len(zellen) < 8:
        return None

    waagerecht = [y for y in _waagerechte_linien(dunkel, int(image.width * 0.2))
                  if rahmen_oben - 4 <= y <= rahmen_unten + 4]
    grenzen = sorted({int(rahmen_oben), *waagerecht, int(rahmen_unten)})

    grau_bild = image.convert("L")
    # Nur Baender in Zeilenhoehe kommen als Nummernzeile infrage.
    kandidaten = [(oben, unten) for oben, unten in zip(grenzen, grenzen[1:])
                  if 10 <= unten - oben <= spaltenbreite * 3][:10]
    treffer = None
    for oben, unten in kandidaten:
        band = grau_bild.crop((0, oben, image.width, unten))
        ergebnis = _pruefe_kopfband(band, zellen)
        if ergebnis is not None:
            treffer = (ergebnis, oben, unten)
            break

    if treffer is None:
        return None
    (kanten, profil, guete, lesbar_anzahl), kopf_oben, kopf_unten = treffer
    log.info("Scan-Raster: %d Spalten, Profil %s, %d/%d Kopfziffern bestaetigt (%.0f%%)",
             len(kanten) - 1, profil.id, round(guete * lesbar_anzahl), lesbar_anzahl, guete * 100)

    body_top = float(kopf_unten)
    zeilen = [y for y in waagerecht if body_top - 2 <= y <= rahmen_unten + 2]
    zeilen_gefunden = len(zeilen) >= 5

    return ScanRaster(
        zeilen_gefunden=zeilen_gefunden,
        spalten_kanten=kanten,
        profil=profil,
        body_top=body_top,
        body_bottom=float(rahmen_unten),
        zeilen_kanten=sorted(zeilen),
        label_rechts=kanten[0],
    )
