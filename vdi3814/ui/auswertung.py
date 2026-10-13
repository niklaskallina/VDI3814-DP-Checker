"""Bausteine der Auswerteseite (Reiter 4 "Auswertung").

Bewusst ohne Streamlit-Aufrufe: hier entstehen Kennzahlen, Qualitaetsbefunde,
HTML-Kacheln und Altair-Diagramme aus den Rohdaten. Die Oberflaeche setzt sie
nur noch zusammen - und die Tests koennen jeden Baustein einzeln pruefen.

Gestaltungsregeln (gelten fuer alle Diagramme dieser Seite):

* Jedes Diagramm zeigt EINE Groesse in EINER Farbe - die Balken tragen die
  Akzentfarbe, Beschriftungen bleiben in Textfarbe.
* Mengen-Raster (Datei x Funktionsgruppe) in einer Farbe von hell nach dunkel.
* Statusfarben (gut / Warnung / kritisch) nur fuer Pruefergebnisse, immer mit
  Symbol und Text - nie Farbe allein.
* Duenne Balken mit abgerundetem Ende, zurueckhaltende Hilfslinien, Werte am
  Balkenende, alles Weitere im Tooltip und in der Tabelle darunter.
"""

from __future__ import annotations

import html
from dataclasses import dataclass, field

import altair as alt
import pandas as pd

from vdi3814 import schwerpunkt

# --------------------------------------------------------------------------
# Farben - hell und dunkel jeweils eigens gestuft
# --------------------------------------------------------------------------

FARBEN = {
    "hell": {
        "flaeche": "#fcfcfb",
        "karte": "#ffffff",
        "text": "#0b0b0b",
        "text2": "#52514e",
        "gedaempft": "#898781",
        "raster": "#e1e0d9",
        "achse": "#c3c2b7",
        "rand": "rgba(11,11,11,0.10)",
        "akzent": "#2a78d6",
        "akzent_hauch": "#eef4fd",
        "verlauf": ["#cde2fb", "#86b6ef", "#3987e5", "#1c5cab", "#104281"],
        "gut": "#006300",
        "warnung": "#8a5a00",
        "kritisch": "#b42323",
        "gut_flaeche": "#e8f5e8",
        "warnung_flaeche": "#fdf3dc",
        "kritisch_flaeche": "#fbe9e9",
    },
    "dunkel": {
        "flaeche": "#1a1a19",
        "karte": "#232322",
        "text": "#f0efec",
        "text2": "#c3c2b7",
        "gedaempft": "#898781",
        "raster": "#2c2c2a",
        "achse": "#383835",
        "rand": "rgba(255,255,255,0.10)",
        "akzent": "#3987e5",
        "akzent_hauch": "#1b2a3d",
        "verlauf": ["#0d366b", "#184f95", "#2a78d6", "#6da7ec", "#b7d3f6"],
        "gut": "#0ca30c",
        "warnung": "#fab219",
        "kritisch": "#e66767",
        "gut_flaeche": "#132b13",
        "warnung_flaeche": "#33280f",
        "kritisch_flaeche": "#3a1818",
    },
}

SCHRIFT = 'system-ui, -apple-system, "Segoe UI", sans-serif'


def farben(dunkel: bool) -> dict:
    return FARBEN["dunkel" if dunkel else "hell"]


def zahl(wert: float, nachkomma: int = 0) -> str:
    """Deutsche Schreibweise: 12.480 bzw. 1.234,50."""
    text = f"{wert:,.{nachkomma}f}"
    return text.replace(",", "X").replace(".", ",").replace("X", ".")


# --------------------------------------------------------------------------
# Kennzahlen
# --------------------------------------------------------------------------

@dataclass
class Kennzahlen:
    funktionen: float = 0.0
    datenpunkte: int = 0
    dateien: int = 0
    schwerpunkte: int = 0
    spalten_belegt: int = 0
    kosten: float | None = None          # None: noch keine Einheitspreise
    spalten_mit_preis: int = 0

    @property
    def funktionen_je_datenpunkt(self) -> float:
        return self.funktionen / self.datenpunkte if self.datenpunkte else 0.0


def kennzahlen(roh: pd.DataFrame, preise: dict[str, float] | None = None) -> Kennzahlen:
    """Eckwerte der (gefilterten) Rohdaten."""
    if roh.empty:
        return Kennzahlen()
    preise = preise or {}
    mengen = roh.groupby("spalte_key")["wert"].sum()
    bepreist = {key: menge for key, menge in mengen.items() if preise.get(key, 0.0) > 0}
    kosten = sum(menge * preise[key] for key, menge in bepreist.items()) if bepreist else None
    return Kennzahlen(
        funktionen=float(roh["wert"].sum()),
        datenpunkte=int(roh.loc[roh["wert"].fillna(0) != 0, "zeile_id"].nunique()),
        dateien=int(roh["dokument_id"].nunique()),
        schwerpunkte=int(roh.loc[roh["schwerpunkt"] != schwerpunkt.OHNE_ZUORDNUNG,
                                 "schwerpunkt"].nunique()),
        spalten_belegt=int(roh.loc[roh["wert"].fillna(0) != 0, "spalte_key"].nunique()),
        kosten=kosten,
        spalten_mit_preis=len(bepreist),
    )


# --------------------------------------------------------------------------
# Qualitaet der Zaehlung
# --------------------------------------------------------------------------

@dataclass
class Befund:
    stufe: str            # gut | warnung | kritisch
    titel: str
    text: str


@dataclass
class Qualitaet:
    seiten_geprueft: int = 0
    seiten_exakt: int = 0
    abweichende_seiten: list[str] = field(default_factory=list)
    ausgeschlossen_mit_werten: int = 0
    entfallen: int = 0
    ohne_zuordnung_menge: float = 0.0
    unbekannte_spalten_menge: float = 0.0
    dubletten: list[tuple[str, str, float]] = field(default_factory=list)
    nur_ocr: list[str] = field(default_factory=list)

    def befunde(self) -> list[Befund]:
        liste: list[Befund] = []
        if self.seiten_geprueft == 0:
            liste.append(Befund("warnung", "Keine Summenzeile zum Abgleich",
                                "Die Listen enthalten keine Zeile „Summe Funktionen“ – "
                                "die Zählung lässt sich nicht gegenprüfen."))
        elif self.abweichende_seiten:
            liste.append(Befund(
                "kritisch",
                f"Summenprüfung: {len(self.abweichende_seiten)} von "
                f"{self.seiten_geprueft} Seiten weichen ab",
                "Betroffen: " + ", ".join(self.abweichende_seiten[:6])
                + (" …" if len(self.abweichende_seiten) > 6 else "")
                + ". Einzelheiten und Fundstelle in Reiter 3."))
        else:
            liste.append(Befund("gut", f"Summenprüfung: alle {self.seiten_geprueft} "
                                       f"Seite(n) exakt",
                                "Die eigene Zählung deckt sich mit der Zeile "
                                "„Summe Funktionen“ jeder geprüften Seite."))
        if self.dubletten:
            a, b, anteil = self.dubletten[0]
            liste.append(Befund(
                "kritisch", f"{len(self.dubletten)} mögliche doppelte Liste(n)",
                f"„{a}“ und „{b}“ enthalten zu {anteil:.0%} dieselben Datenpunkte – "
                "die Mengen zählen sonst doppelt. Eine der Dateien in Reiter 7 löschen, "
                "falls es derselbe Plansatz ist."))
        if self.ausgeschlossen_mit_werten:
            liste.append(Befund(
                "warnung", f"{self.ausgeschlossen_mit_werten} nicht gezählte Zeile(n) mit Werten",
                "Summen-, Übertrags- oder Kopfzeilen mit Zahlen – bewusst nicht gezählt. "
                "Kurzer Blick in Reiter 3 lohnt sich."))
        if self.entfallen:
            liste.append(Befund(
                "warnung", f"{self.entfallen} entfallene Datenpunkt(e)",
                "In Excel ausgeblendet oder durchgestrichen – nicht gezählt."))
        if self.unbekannte_spalten_menge:
            liste.append(Befund(
                "warnung", f"{zahl(self.unbekannte_spalten_menge)} Funktionen ohne Spaltenzuordnung",
                "Diese Werte stehen in Spalten, die keiner Spalte der VDI-Vorlage "
                "zugeordnet werden konnten – sie fehlen in der Kostenschätzung."))
        if self.ohne_zuordnung_menge:
            liste.append(Befund(
                "warnung", f"{zahl(self.ohne_zuordnung_menge)} Funktionen ohne ASP/ISP",
                "Für diese Datenpunkte ist kein Automations-/Informationsschwerpunkt "
                "angegeben – nachtragen in Reiter 2."))
        if self.nur_ocr:
            liste.append(Befund(
                "warnung", f"{len(self.nur_ocr)} Datei(en) per Texterkennung gelesen",
                "Gescannte Listen – Werte bitte stichprobenartig gegen das Original "
                "prüfen (Reiter 2 und 3)."))
        return liste

    @property
    def gesamtstufe(self) -> str:
        stufen = [b.stufe for b in self.befunde()]
        if "kritisch" in stufen:
            return "kritisch"
        return "warnung" if "warnung" in stufen else "gut"


def qualitaet(session, roh: pd.DataFrame, dokument_ids: set[int] | None = None) -> Qualitaet:
    """Sammelt alles, was an der Zaehlung pruefenswert ist."""
    from sqlalchemy import select

    from vdi3814 import aggregate, db, evidence

    ergebnis = Qualitaet()
    for document in session.scalars(select(db.Document)):
        if dokument_ids is not None and document.id not in dokument_ids:
            continue
        abgleich = evidence.seitenabgleich(document)
        seiten = sorted({e["seite"] for e in abgleich})
        abweichend = sorted({e["seite"] for e in abgleich
                             if e["gemeldet"] is not None
                             and abs(e["gemeldet"] - e["eigen"]) >= 1e-6})
        ergebnis.seiten_geprueft += len(seiten)
        ergebnis.seiten_exakt += len(seiten) - len(abweichend)
        ergebnis.abweichende_seiten += [f"{document.file_name} S.{s + 1}" for s in abweichend]
        for row in document.rows:
            if row.kind == "entfallen":
                ergebnis.entfallen += 1
            elif row.kind not in ("daten", "leer") and any(
                    v.count for v in row.values):
                if row.kind == "summe" and row.page_index in seiten \
                        and row.page_index not in abweichend:
                    continue      # passende Summenzeile - normal
                ergebnis.ausgeschlossen_mit_werten += 1
        if document.backend and "ocr" in document.backend:
            ergebnis.nur_ocr.append(document.file_name)

    if not roh.empty:
        ergebnis.ohne_zuordnung_menge = float(
            roh.loc[roh["schwerpunkt"] == schwerpunkt.OHNE_ZUORDNUNG, "wert"].sum())
        ergebnis.unbekannte_spalten_menge = float(
            roh.loc[roh["spalte_key"].fillna("") == "", "wert"].sum())
        dubletten = aggregate.moegliche_dubletten(roh)
        ergebnis.dubletten = [(d.datei_a, d.datei_b, float(d.anteil))
                              for d in dubletten.itertuples()]
    return ergebnis


# --------------------------------------------------------------------------
# HTML-Bausteine
# --------------------------------------------------------------------------

def _stil(f: dict) -> str:
    return f"""
<style>
.vdi-kpi {{ display:grid; grid-template-columns: minmax(220px,1.4fr) repeat(auto-fit, minmax(150px,1fr));
  gap:12px; font-family:{SCHRIFT}; margin: 4px 0 8px; }}
.vdi-kachel {{ background:{f['karte']}; border:1px solid {f['rand']}; border-radius:12px;
  padding:14px 16px; min-width:0; }}
.vdi-kachel .l {{ color:{f['text2']}; font-size:0.82rem; margin-bottom:6px; }}
.vdi-kachel .v {{ color:{f['text']}; font-size:1.6rem; font-weight:600; line-height:1.15; }}
.vdi-kachel .s {{ color:{f['gedaempft']}; font-size:0.78rem; margin-top:6px; }}
.vdi-kachel.hero {{ background:{f['akzent_hauch']}; border-color:transparent; }}
.vdi-kachel.hero .v {{ font-size:3rem; font-weight:650; letter-spacing:-0.01em; }}
.vdi-kachel.hero .balken {{ height:4px; border-radius:2px; background:{f['akzent']};
  width:56px; margin-bottom:10px; }}
.vdi-qs {{ display:flex; flex-wrap:wrap; gap:8px; font-family:{SCHRIFT}; margin: 2px 0 6px; }}
.vdi-chip {{ display:flex; gap:8px; align-items:flex-start; border-radius:10px; padding:9px 12px;
  border:1px solid {f['rand']}; flex: 1 1 280px; min-width:0; }}
.vdi-chip .i {{ font-weight:700; font-size:0.95rem; line-height:1.3; }}
.vdi-chip .t {{ color:{f['text']}; font-weight:600; font-size:0.86rem; }}
.vdi-chip .d {{ color:{f['text2']}; font-size:0.8rem; margin-top:2px; }}
.vdi-chip.gut {{ background:{f['gut_flaeche']}; }} .vdi-chip.gut .i {{ color:{f['gut']}; }}
.vdi-chip.warnung {{ background:{f['warnung_flaeche']}; }} .vdi-chip.warnung .i {{ color:{f['warnung']}; }}
.vdi-chip.kritisch {{ background:{f['kritisch_flaeche']}; }} .vdi-chip.kritisch .i {{ color:{f['kritisch']}; }}
.vdi-titel {{ font-family:{SCHRIFT}; color:{f['text']}; font-weight:600; font-size:1.02rem; margin:18px 0 2px; }}
.vdi-unter {{ font-family:{SCHRIFT}; color:{f['text2']}; font-size:0.82rem; margin:0 0 6px; }}
</style>"""


def kacheln_html(werte: Kennzahlen, waehrung: str, dunkel: bool = False) -> str:
    """KPI-Zeile: die Gesamtmenge gross, daneben die uebrigen Eckwerte."""
    f = farben(dunkel)
    kosten_wert = zahl(werte.kosten) if werte.kosten is not None else "–"
    kosten_zeile = (f"{werte.spalten_mit_preis} Spalte(n) mit Einheitspreis"
                    if werte.kosten is not None else "Einheitspreise in Reiter 5 eintragen")
    kacheln = [
        ("Datenpunkte", zahl(werte.datenpunkte),
         f"Ø {zahl(werte.funktionen_je_datenpunkt, 1)} Funktionen je Datenpunkt"),
        ("Dateien", zahl(werte.dateien), f"{werte.spalten_belegt} Funktionsspalten belegt"),
        ("Schwerpunkte", zahl(werte.schwerpunkte), "ASP/ISP mit Zuordnung"),
        ("Geschätzte Kosten", kosten_wert,
         f"{waehrung} · {kosten_zeile}" if werte.kosten is not None else kosten_zeile),
    ]
    teile = [
        '<div class="vdi-kachel hero"><div class="balken"></div>'
        '<div class="l">Funktionen gesamt</div>'
        f'<div class="v">{zahl(werte.funktionen)}</div>'
        '<div class="s">gezählte Funktionen über alle gefilterten Listen</div></div>'
    ]
    for titel, wert, unter in kacheln:
        teile.append(f'<div class="vdi-kachel"><div class="l">{html.escape(titel)}</div>'
                     f'<div class="v">{html.escape(wert)}</div>'
                     f'<div class="s">{html.escape(unter)}</div></div>')
    return _stil(f) + '<div class="vdi-kpi">' + "".join(teile) + "</div>"


_SYMBOL = {"gut": "✓", "warnung": "!", "kritisch": "✕"}
_STUFE_TEXT = {"gut": "In Ordnung", "warnung": "Hinweis", "kritisch": "Prüfen"}


def qualitaet_html(befunde: list[Befund], dunkel: bool = False) -> str:
    """Pruefergebnisse als Streifen - Symbol, Stufe und Text, nie Farbe allein."""
    f = farben(dunkel)
    teile = []
    for befund in befunde:
        teile.append(
            f'<div class="vdi-chip {befund.stufe}" role="status" '
            f'aria-label="{_STUFE_TEXT[befund.stufe]}">'
            f'<div class="i" aria-hidden="true">{_SYMBOL[befund.stufe]}</div>'
            f'<div><div class="t">{html.escape(befund.titel)}</div>'
            f'<div class="d">{html.escape(befund.text)}</div></div></div>')
    return _stil(f) + '<div class="vdi-qs">' + "".join(teile) + "</div>"


def titel_html(titel: str, unter: str = "", dunkel: bool = False) -> str:
    f = farben(dunkel)
    text = f'<div class="vdi-titel">{html.escape(titel)}</div>'
    if unter:
        text += f'<div class="vdi-unter">{html.escape(unter)}</div>'
    return _stil(f) + text


# --------------------------------------------------------------------------
# Diagramme
# --------------------------------------------------------------------------

# Achsenwerte in deutscher Schreibweise: 1.000 statt 1,000 bzw. 1k
_DEUTSCHE_ZAHL = "replace(format(datum.value, ',d'), ',', '.')"


def _konfiguration(chart: alt.Chart, f: dict) -> alt.Chart:
    return (chart
            .configure(background="transparent", font=SCHRIFT)
            .configure_view(stroke=None)
            .configure_axis(labelColor=f["text2"], titleColor=f["text2"], gridColor=f["raster"],
                            domainColor=f["achse"], tickColor=f["achse"], labelFontSize=11,
                            titleFontSize=11, titleFontWeight="normal", labelLimit=260)
            .configure_legend(labelColor=f["text2"], titleColor=f["text2"]))


def _balken(daten: pd.DataFrame, kategorie: str, kategorie_titel: str, f: dict,
            tooltip: list, sortierung="-x", zeilenhoehe: int = 30) -> alt.Chart:
    """Waagerechte Balken einer Groesse - eine Farbe, Wert am Balkenende."""
    hoehe = max(90, zeilenhoehe * len(daten))
    basis = alt.Chart(daten).encode(
        y=alt.Y(f"{kategorie}:N", sort=sortierung, title=None,
                axis=alt.Axis(ticks=False, domain=False, labelPadding=8)),
        x=alt.X("menge:Q", title="Funktionen",
                axis=alt.Axis(grid=True, tickCount=5, domain=False, ticks=False,
                              labelExpr=_DEUTSCHE_ZAHL),
                scale=alt.Scale(nice=True)),
    )
    balken = basis.mark_bar(color=f["akzent"], cornerRadiusEnd=4,
                            size=min(20, zeilenhoehe - 10)).encode(tooltip=tooltip)
    werte = basis.mark_text(align="left", dx=6, color=f["text2"], fontSize=11).encode(
        text=alt.Text("menge_text:N"))
    return (balken + werte).properties(height=hoehe)


def diagramm_gruppen(roh: pd.DataFrame, dunkel: bool = False) -> alt.Chart | None:
    """Funktionen je Funktionsgruppe (Ein-/Ausgabe, Verarbeitung, ...)."""
    if roh.empty:
        return None
    f = farben(dunkel)
    daten = (roh.groupby("gruppe", dropna=False)["wert"].sum()
             .reset_index(name="menge"))
    daten = daten[daten["menge"] != 0]
    if daten.empty:
        return None
    gesamt = daten["menge"].sum()
    daten["gruppe"] = daten["gruppe"].fillna("").replace("", "ohne Gruppe")
    daten["anteil"] = daten["menge"] / gesamt
    daten["menge_text"] = [zahl(m) for m in daten["menge"]]
    chart = _balken(daten, "gruppe", "Funktionsgruppe", f, tooltip=[
        alt.Tooltip("gruppe:N", title="Funktionsgruppe"),
        alt.Tooltip("menge:Q", title="Funktionen", format=",.0f"),
        alt.Tooltip("anteil:Q", title="Anteil", format=".1%"),
    ])
    return _konfiguration(chart, f)


def diagramm_schwerpunkte(roh: pd.DataFrame, dunkel: bool = False) -> alt.Chart | None:
    """Funktionen je ASP/ISP - in der Reihenfolge der Kennungen."""
    if roh.empty:
        return None
    f = farben(dunkel)
    daten = (roh.groupby(["schwerpunkt", "schwerpunkt_bezeichnung"], dropna=False)
             .agg(menge=("wert", "sum"), datenpunkte=("zeile_id", "nunique"))
             .reset_index())
    daten = daten[daten["menge"] != 0]
    if daten.empty:
        return None
    daten["schwerpunkt_bezeichnung"] = daten["schwerpunkt_bezeichnung"].fillna("")
    daten["beschriftung"] = [
        (f"{k} · {b}" if b else k)[:48]
        for k, b in zip(daten["schwerpunkt"], daten["schwerpunkt_bezeichnung"])]
    daten["anteil"] = daten["menge"] / daten["menge"].sum()
    daten["menge_text"] = [zahl(m) for m in daten["menge"]]
    daten["_ende"] = daten["schwerpunkt"] == schwerpunkt.OHNE_ZUORDNUNG
    daten = daten.sort_values(["_ende", "schwerpunkt", "schwerpunkt_bezeichnung"])
    reihenfolge = list(daten["beschriftung"])
    chart = _balken(daten.drop(columns="_ende"), "beschriftung", "Schwerpunkt", f,
                    sortierung=reihenfolge, tooltip=[
                        alt.Tooltip("schwerpunkt:N", title="Schwerpunkt"),
                        alt.Tooltip("schwerpunkt_bezeichnung:N", title="Bezeichnung"),
                        alt.Tooltip("menge:Q", title="Funktionen", format=",.0f"),
                        alt.Tooltip("datenpunkte:Q", title="Datenpunkte"),
                        alt.Tooltip("anteil:Q", title="Anteil", format=".1%"),
                    ])
    return _konfiguration(chart, f)


def diagramm_spalten(zusammenfassung: pd.DataFrame, anzahl: int = 12,
                     dunkel: bool = False) -> alt.Chart | None:
    """Die meistbelegten Funktionsspalten."""
    if zusammenfassung.empty:
        return None
    f = farben(dunkel)
    daten = zusammenfassung[zusammenfassung["menge"] != 0].nlargest(anzahl, "menge").copy()
    if daten.empty:
        return None
    daten["beschriftung"] = [f"{a} · {s}"[:52] for a, s in
                             zip(daten["spalte_adresse"].fillna(""), daten["spalte"].fillna(""))]
    daten["menge_text"] = [zahl(m) for m in daten["menge"]]
    chart = _balken(daten, "beschriftung", "Funktionsspalte", f, tooltip=[
        alt.Tooltip("spalte_adresse:N", title="Abschnitt.Spalte"),
        alt.Tooltip("spalte:N", title="Funktionsspalte"),
        alt.Tooltip("gruppe:N", title="Gruppe"),
        alt.Tooltip("menge:Q", title="Funktionen", format=",.0f"),
        alt.Tooltip("datenpunkte:Q", title="Datenpunkte"),
    ], zeilenhoehe=26)
    return _konfiguration(chart, f)


def diagramm_raster(roh: pd.DataFrame, dunkel: bool = False) -> alt.Chart | None:
    """Datei x Funktionsgruppe als Mengen-Raster (eine Farbe, hell -> dunkel)."""
    if roh.empty or roh["dokument_id"].nunique() < 2:
        return None
    f = farben(dunkel)
    daten = (roh.groupby(["datei", "gruppe"], dropna=False)["wert"].sum()
             .reset_index(name="menge"))
    daten["gruppe"] = daten["gruppe"].fillna("").replace("", "ohne Gruppe")
    daten = daten[daten["menge"] != 0]
    if daten.empty:
        return None
    gruppen_reihenfolge = (daten.groupby("gruppe")["menge"].sum()
                           .sort_values(ascending=False).index.tolist())
    grenze = daten["menge"].max()
    # Die Luecke zwischen den Feldern entsteht ueber den Bandabstand - kein
    # Rahmen um die Felder.
    chart = alt.Chart(daten).mark_rect(cornerRadius=3).encode(
        x=alt.X("gruppe:N", sort=gruppen_reihenfolge, title=None,
                scale=alt.Scale(paddingInner=0.08),
                axis=alt.Axis(labelAngle=-30, labelLimit=160, ticks=False, domain=False)),
        y=alt.Y("datei:N", title=None, scale=alt.Scale(paddingInner=0.08),
                axis=alt.Axis(ticks=False, domain=False, labelLimit=240)),
        color=alt.Color("menge:Q", title="Funktionen",
                        scale=alt.Scale(range=f["verlauf"], domain=[0, grenze]),
                        legend=alt.Legend(orient="right", gradientLength=120,
                                          labelExpr=_DEUTSCHE_ZAHL)),
        tooltip=[alt.Tooltip("datei:N", title="Datei"),
                 alt.Tooltip("gruppe:N", title="Funktionsgruppe"),
                 alt.Tooltip("menge:Q", title="Funktionen", format=",.0f")],
    ).properties(height=alt.Step(34))
    return _konfiguration(chart, f)


def raster_untertitel(dunkel: bool) -> str:
    """Im dunklen Schema steigt die Menge zur hellen Farbe hin an."""
    return ("Je heller, desto mehr Funktionen" if dunkel
            else "Je dunkler, desto mehr Funktionen")
