# VDI 3814 DP-Checker

Auswertung von **GA-Funktionslisten nach VDI 3814 Blatt 1** (alte und neue Fassung) –
Batch-Import vieler Dateien, automatische Spalten- und Zeilenerkennung, editierbare
Vorschau, projektübergreifende Aggregation, Kostenschätzung und Excel-Export.

**Vollständig offline.** Es wird keine Cloud-KI und keine externe Schnittstelle
verwendet; die Mustererkennung läuft über ein lokal installiertes
Vision-Language-Modell (Ollama auf `127.0.0.1`). Projektdaten verlassen das Gerät nicht.

---

## Was das Tool kann

| Anforderung | Umsetzung |
|---|---|
| Batch-Import | beliebig viele PDF/PNG/JPEG/TIFF **und** die Excel-Vorlagen (`.xlsx`, `.xls`), gemischt |
| Mehrseitige, gescannte PDFs | ja – jede Seite wird einzeln klassifiziert und verarbeitet |
| Alte **und** neue Fassung | automatische Erkennung anhand der gelesenen Kopfstruktur, keine Abfrage nötig |
| Regelschemata in derselben Datei | werden erkannt, übersprungen und **protokolliert** (Blatt „Dokumente“ im Export) |
| Spalten-/Zeilenerkennung | über das normative Adressschema `Abschnitt.Spalte`, nicht über feste Pixelpositionen |
| Fußnoten (Zählregeln) | werden ausgelesen, der referenzierten Spalte zugeordnet und exportiert |
| Freitext in Zellen | wird als Anmerkung geführt und **nie** als Zählwert addiert |
| Summenzeile | wird nicht übernommen, sondern gegen die eigene Aufsummierung geprüft |
| Vorschau & Korrektur | editierbare Tabelle je Datei, erst danach Speichern/Export |
| Persistenz | lokale SQLite-Datenbank, inkrementell erweiterbar, Re-Export ohne erneute Erkennung |
| Kostenschätzung | Einheitspreis je Funktionsspalte, jederzeit änderbar, sofortige Neuberechnung |
| Excel-Export | `.xlsx` mit Übersicht, Kostenschätzung (**mit Formeln**), Rohdaten, Dokumente, Fußnoten, Projekte |

---

## Drei Erkennungswege – automatisch gewählt

Das Tool nimmt je Datei/Seite den genauesten verfügbaren Weg:

| Quelle | Verfahren | Genauigkeit |
|---|---|---|
| `.xlsx` / `.xls` | direktes Auslesen der Zellen | exakt |
| PDF **mit Textebene** (aus Excel/CAD gedruckt) | Rekonstruktion des Rasters aus den Wortkoordinaten | exakt |
| Scan, Foto, PNG/JPEG, bildbasiertes PDF | lokales Vision-Language-Modell + optionales OCR | modellabhängig, in der Vorschau korrigierbar |

Der Anker ist in allen Fällen derselbe: die beiden Kopfzeilen **„Abschnitt“** und
**„Spalte“**, die in beiden VDI-Vorlagen enthalten sind. Aus ihnen ergibt sich für jede
Funktionsspalte eine eindeutige Adresse (z. B. `6.13` = *Tarifabhängiges Schalten* in
der alten, `2.2.10` = *Frostschutzsteuerung* in der neuen Fassung). Diese Adresse ist
robuster als die senkrecht gedruckten Spaltenüberschriften; die Bezeichnungstexte
dienen nur als Rückfallebene.

---

## Installation (Windows)

### 1. Python

Python 3.10 oder neuer von [python.org](https://www.python.org/downloads/windows/)
installieren (bei der Installation **„Add python.exe to PATH“** anhaken).

```bat
cd C:\Pfad\zu\VDI3814-DP-Checker
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Ollama + Vision-Modell (für Scans und Bilder)

Nur nötig, wenn gescannte PDFs, Fotos oder PNG/JPEG ausgewertet werden sollen.
Für Excel-Dateien und Text-PDFs läuft das Tool auch ohne.

1. Ollama von [ollama.com/download](https://ollama.com/download) installieren.
2. Einmalig – **hierfür wird Internet gebraucht, danach nie wieder**:

```bat
ollama pull qwen2.5vl:7b
```

`qwen2.5-VL` ist derzeit das offline verfügbare Modell mit dem besten
Dokument-/Tabellenverständnis in dieser Größenklasse. Alternativen, die sich mit
`--modell` bzw. über die Oberfläche einstellen lassen:

| Modell | RAM/VRAM | Hinweis |
|---|---|---|
| `qwen2.5vl:7b` | ~8 GB | Voreinstellung, bestes Tabellenverständnis |
| `qwen2.5vl:3b` | ~4 GB | für schwächere Rechner |
| `llama3.2-vision:11b` | ~10 GB | Alternative, etwas schwächer bei dichten Tabellen |
| `minicpm-v` | ~6 GB | kompakt, gute OCR-Eigenschaften |

Ollama läuft als lokaler Dienst auf `http://127.0.0.1:11434`. Prüfen:

```bat
python -m vdi3814.cli check
```

### 3. Tesseract-OCR (optional)

Verbessert die Fassungserkennung bei schlechten Scans.
[UB-Mannheim-Installer](https://github.com/UB-Mannheim/tesseract/wiki) inkl.
deutschem Sprachpaket installieren; falls nicht im PATH:

```bat
set VDI_TESSERACT_CMD=C:\Program Files\Tesseract-OCR\tesseract.exe
```

---

## Start

### Oberfläche (empfohlen)

```bat
.venv\Scripts\activate
streamlit run vdi3814\ui\app.py
```

Der Browser öffnet `http://localhost:8501`. Die Oberfläche läuft rein lokal – der
Browser ist nur die Anzeige, es geht nichts ins Netz.

> **Warum Streamlit und keine Desktop-GUI?** Der Arbeitsablauf ist tabellarisch:
> mehrere hundert Zeilen × 50–57 Spalten prüfen und korrigieren. `st.data_editor`
> liefert dafür eine Excel-ähnliche Bearbeitung inklusive Filtern und Sortieren, die
> mit Tkinter/Qt erst aufwendig nachgebaut werden müsste. Zusätzlich entfällt jede
> GUI-Paketierung – ein `pip install` genügt. Wer ein echtes Fenster statt eines
> Browser-Tabs möchte, startet Streamlit im App-Modus
> (`streamlit run … --server.headless true` + Browser-Verknüpfung) oder nutzt die CLI.

### Kommandozeile (Batch ohne Oberfläche)

```bat
python -m vdi3814.cli check                          :: Installation prüfen
python -m vdi3814.cli import C:\Projekte\Listen      :: Ordner importieren
python -m vdi3814.cli import liste.pdf --nur-pruefen :: nur erkennen, nichts speichern
python -m vdi3814.cli list                           :: importierte Dateien anzeigen
python -m vdi3814.cli preise --setzen A_1_1=45 A_1_5=80
python -m vdi3814.cli preise --vorlage vdi3814_alt   :: Preisvorlage als CSV
python -m vdi3814.cli preise --csv-import preise.csv
python -m vdi3814.cli export Auswertung.xlsx
```

---

## Beispiel: kompletter Ablauf in 3 Minuten

Der Beispieldatensatz enthält eine Liste der **alten** Fassung (mit einer zweiten
Seite „Regelschema“, die übersprungen werden muss) und eine der **neuen** Fassung.

```bat
python sample_data\make_samples.py            :: Beispieldateien erzeugen
python -m vdi3814.cli import sample_data --kein-modell
python -m vdi3814.cli preise --setzen A_1_5=80 A_1_1=45 N_1_1_1=75
python -m vdi3814.cli export out\Auswertung.xlsx
```

Ausgabe des Imports:

```
   beispiel_alte_fassung.pdf Seite 1: Tabelle ueber PDF-Textebene gelesen
   beispiel_alte_fassung.pdf Seite 2: uebersprungen (regelschema)
   beispiel_neue_fassung.pdf Seite 1: Tabelle ueber PDF-Textebene gelesen
   beispiel_alte_fassung.pdf: 8 Datenpunkte, Summe 31, Fassung alt
   beispiel_neue_fassung.pdf: 8 Datenpunkte, Summe 30, Fassung neu
```

Denselben Ablauf gibt es in der Oberfläche über die Reiter
**1 Import → 2 Vorschau & Korrektur → 3 Gesamtübersicht → 4 Kostenschätzung → 5 Export**.

---

## Aufbau des Excel-Exports

| Blatt | Inhalt |
|---|---|
| **Übersicht** | Summe je Funktionsspalte über alle importierten Listen, mit Abschnitt.Spalte, Gruppe, Anzahl Dateien; unten „Summe Funktionen“ als Summenformel |
| **Kostenschätzung** | je Spalte: Menge als **Formel** auf „Übersicht“, Einheitspreis, Kosten `= Menge × Einheitspreis`, Gesamtkosten als Summenformel – in Excel direkt weiterrechenbar |
| **Rohdaten** | jede einzelne Zelle mit Quelldatei, Seite, Zeilennummer, Datenpunkt, BAS, Spalte und Wert |
| **Dokumente** | importierte Dateien: Fassung, Erkennungsverfahren, Metadaten, ausgewertete und **übersprungene** Seiten, Hinweise |
| **Fußnoten** | Zählregeln aus der Kopfzeile inkl. der Spalten, auf die sie sich beziehen |
| **Projekte** | Kreuztabelle Projekt/Anlage × Spaltengruppe |

---

## Spalten-Mapping

Die vollständige Zuordnung beider Fassungen steht in
**[`docs/spalten-mapping.md`](docs/spalten-mapping.md)** (50 Spalten alt, 56 Spalten neu,
je mit Adresse, Schlüssel, Gruppe, Bezeichnung, Kürzel und Fußnotenbezug).
Die Datei wird aus den Profilen erzeugt:

```bat
python tools\gen_mapping_doc.py
```

### Eigene bzw. abweichende Layouts ergänzen

Die Profile sind reine YAML-Dateien in `vdi3814/profiles/`. Ein weiteres Layout
(anderes Gewerk, andere Planersteller-Vorlage, weiteres VDI-Blatt) wird ergänzt, **ohne
den Erkennungskern zu ändern**:

1. `vdi3814/profiles/vdi3814_alt.yaml` kopieren, `id` und `name` ändern.
2. Unter `detection.keywords` die Begriffe eintragen, die nur in diesem Layout vorkommen.
3. Unter `columns` die Spalten mit `address` (`Abschnitt.Spalte`), `label` und
   `aliases` pflegen. `aggregate_group` bestimmt, wie die Spalte in der
   Gesamtübersicht zusammengefasst wird.
4. Fertig – das Profil wird beim nächsten Start automatisch geladen
   (`python -m vdi3814.cli profile`).

---

## Konfiguration

Alle Einstellungen lassen sich über Umgebungsvariablen setzen (Präfix `VDI_`):

| Variable | Voreinstellung | Bedeutung |
|---|---|---|
| `VDI_DB_PATH` | `data/vdi3814.sqlite3` | Pfad der Datenbank |
| `VDI_VISION_MODEL` | `qwen2.5vl:7b` | Ollama-Modell |
| `VDI_OLLAMA_HOST` | `http://127.0.0.1:11434` | Ollama-Adresse |
| `VDI_RENDER_DPI` | `300` | Auflösung beim Rastern von PDF-Seiten |
| `VDI_BAND_COUNT` | `3` | Anzahl Zeilenstreifen je Seite beim Vision-Verfahren |
| `VDI_USE_OCR` | `1` | ergänzendes Tesseract-OCR |
| `VDI_TESSERACT_CMD` | – | Pfad zu `tesseract.exe`, falls nicht im PATH |
| `VDI_CURRENCY` | `EUR` | Währung der Kostenschätzung |

---

## Fehlerbehandlung

* Nicht lesbare Seiten führen **nie** zum Abbruch – die Datei wird mit einem Hinweis
  weiterverarbeitet, der Hinweis erscheint in der Oberfläche und im Blatt „Dokumente“.
* Wird die Fassung nicht sicher erkannt, meldet das Tool die Bewertungen aller Profile
  und die Zuordnung kann in der Vorschau manuell nachgezogen werden.
* Weicht die Zeile „Summe Funktionen“ von der eigenen Aufsummierung ab, wird die
  betroffene Spalte mit beiden Werten gemeldet.
* Fehlt Ollama oder das Modell, arbeitet das Tool ohne Bilderkennung weiter
  (Excel und Text-PDFs) und sagt das ausdrücklich.

---

## Entwicklung

```bat
.venv\Scripts\activate
pytest -q
```

Projektstruktur:

```
vdi3814/
  profiles/            Spalten-Profile beider Fassungen (YAML, erweiterbar)
  ingest/              Datei -> Seitenbild, Bildvorverarbeitung (Deskew, Kontrast)
  extract/             Excel-, PDF-Textebenen- und Vision-Extraktion
  vision/              Ollama-Client, Prompts, OCR, Test-Backend
  ui/app.py            Streamlit-Oberfläche
  pipeline.py          Ablaufsteuerung inkl. Fassungserkennung und Summenprüfung
  db.py                SQLite/SQLAlchemy
  aggregate.py         Auswertung und Drilldown
  costs.py             Einheitspreise und Kosten
  export_excel.py      Excel-Export mit Formeln
sample_data/           Beispieldatensatz (Generator + PDFs)
docs/                  Spalten-Mapping
tests/                 Pytest-Suite
```
