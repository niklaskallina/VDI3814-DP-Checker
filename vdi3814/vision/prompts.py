"""Prompts fuer das lokale Vision-Language-Modell.

Bewusst deutschsprachig und eng gefuehrt: es wird immer ein festes
JSON-Schema verlangt, damit die Antwort maschinell verwertbar ist.
"""

from __future__ import annotations

CLASSIFY = """Du analysierst eine Seite aus einer Planungsunterlage der Gebaeudeautomation.

Entscheide, was auf der Seite zu sehen ist:
- "funktionsliste": eine tabellarische GA-Funktionsliste / Datenpunktliste
  (viele schmale Spalten mit Zahlen, Kopfzeilen wie "Ein-/Ausgabefunktionen",
  "Verarbeitungsfunktionen", "Anwendungsfunktionen", "Abschnitt", "Spalte",
  meist mit einer Zeile "Summe Funktionen").
- "regelschema": ein Regelschema, Prinzipschaltbild, Anlagenschema oder
  Strangschema (Symbole, Rohrleitungen, Ventile, keine Tabelle).
- "sonstiges": Deckblatt, Textseite, Legende, etwas anderes.

Antworte NUR mit JSON:
{"typ": "funktionsliste|regelschema|sonstiges", "konfidenz": 0.0-1.0, "begruendung": "kurz"}
"""

HEADER = """Das Bild zeigt den Kopfbereich einer GA-Funktionsliste nach VDI 3814.

Lies die Spaltenstruktur aus. Wichtig sind zwei Nummernzeilen direkt ueber der
ersten Datenzeile:
- "Abschnitt": Gruppennummer, z. B. "1", "2.2." oder "0."
- "Spalte":    laufende Nummer innerhalb des Abschnitts, z. B. "1" ... "13"

Gib fuer JEDE Funktionsspalte von links nach rechts einen Eintrag zurueck.
Die Spaltenbezeichnungen sind meist senkrecht gedruckt (von unten nach oben lesen).
Hochgestellte Ziffern oder "3)" hinter einer Bezeichnung sind Fussnotenverweise.

Antworte NUR mit JSON:
{"spalten": [{"index": 0, "gruppe": "", "untergruppe": "", "bezeichnung": "",
              "abschnitt": "", "spalte_nr": "", "fussnoten": []}],
 "fussnoten": [{"marker": "1", "text": ""}]}
"""

ROWS = """Das Bild zeigt einen Ausschnitt einer GA-Funktionsliste nach VDI 3814:
oben der Tabellenkopf, darunter Datenpunktzeilen.

Lies jede Datenpunktzeile im unteren Bereich aus:
- "zeile_nr": die Nummer ganz links (falls vorhanden)
- "klartext": die Benennung des Datenpunkts, z. B. "KüW VL-Temperatur"
- "bas": die Benutzeradresse, z. B. "*ASE01*DEV" (falls vorhanden)
- "qualifikation": z. B. "komplex" (falls vorhanden)
- "werte": Objekt {Spaltenindex: Wert} NUR fuer Spalten mit Inhalt.
  Spaltenindex ist die Position der Funktionsspalte von links, beginnend bei 0.
  Werte sind Zahlen ("1", "2", "16") oder ein Kreuz ("x").
- "bemerkung": Freitext aus der rechten Bemerkungsspalte (kein Zaehlwert!)
- "summenzeile": true, wenn es die Zeile "Summe Funktionen" ist

Leere Zellen NICHT ausgeben. Nichts erfinden - im Zweifel weglassen.

Antworte NUR mit JSON:
{"zeilen": [{"zeile_nr": "", "klartext": "", "bas": "", "qualifikation": "",
             "werte": {}, "bemerkung": "", "summenzeile": false}]}
"""

METADATA = """Das Bild zeigt den Kopf- bzw. Fussbereich einer GA-Funktionsliste.

Lies die Projektangaben aus. Nicht vorhandene Angaben als leeren String liefern.

Antworte NUR mit JSON:
{"projekt": "", "auftraggeber": "", "anlage": "", "gewerk": "", "planersteller": "",
 "protokoll": "", "planstand": "", "blatt_nr": "", "blatt_von": "", "datum": "",
 "vdi_blatt": "", "informationsschwerpunkt": ""}
"""
