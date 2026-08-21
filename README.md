# VDI 3814 DP-Checker

Datenpunkte aus **GA-Funktionslisten nach VDI 3814** (alte und neue Fassung) automatisch
auszählen – für Kalkulation und Mengenermittlung. Viele Listen auf einmal einlesen,
Ergebnis nachvollziehbar prüfen, nach Projekten getrennt speichern, mit Einheitspreisen
bewerten und als Excel-Datei ausgeben.

**Ohne Zusatzsoftware und ohne KI-Dienst.** Alles läuft auf dem eigenen Rechner:
Excel-Vorlagen und PDFs werden geometrisch ausgewertet, gescannte Listen mit der
mitgelieferten Texterkennung. Projektdaten verlassen das Gerät nicht.

---

## Für Anwender: installieren und loslegen

1. `VDI3814-DP-Checker-Setup.exe` aus dem [Release](../../releases/latest) herunterladen
2. Doppelklick – das Setup installiert alles in einem Schritt (keine Administratorrechte nötig)
3. Programm über das Startmenü starten; die Oberfläche öffnet sich im Browser und
   läuft ausschließlich lokal

Die Setup-Datei kann unverändert an Kolleginnen und Kollegen weitergegeben werden –
sie enthält alles Nötige.

*Neue Fassung erzeugen:* Reiter **Actions → „Windows-EXE bauen" → „Run workflow"**,
dort den Haken „Ergebnis als Release veroeffentlichen" setzen und eine Versionsnummer
angeben. Der Build erstellt Setup und ZIP und veröffentlicht beides.

Es muss **nichts weiter installiert** werden – weder Python noch Texterkennung noch
ein KI-Modell. Wer nichts installieren möchte, nimmt stattdessen die ZIP-Datei,
entpackt sie und startet `VDI3814-DP-Checker.exe`.

### Der Ablauf in der Oberfläche

| Reiter | Wozu |
|---|---|
| **1 Import** | Dateien auswählen (mehrere gleichzeitig) und einlesen |
| **2 Prüfen & korrigieren** | Erkanntes Ergebnis ansehen, Werte bei Bedarf ändern, einzelne oder mehrere ganze Zeilen löschen, speichern |
| **3 Nachweis & Differenzen** | Jeden Befund anklicken – die Originalseite wird an der betreffenden Stelle markiert angezeigt |
| **4 Gesamtübersicht** | Summen über alle Listen, nach Schwerpunkt (ASP/ISP), Projekt/Anlage/Gewerk aufschlüsselbar |
| **5 Kostenschätzung** | Einheitspreis je Funktionsspalte eintragen, Kosten rechnen sofort mit |
| **6 Export** | Excel-Datei mit Funktionsliste, Summen, Kostenblatt und Rohdaten |
| **7 Projekt & Daten** | Projekte anlegen/löschen, einzelne Dateien aus der Datenbank entfernen |

Jedes **Projekt** hat eine eigene Datenbank und einen eigenen Ablageordner – Auswertungen
verschiedener Bauvorhaben vermischen sich also nicht. Projekte lassen sich jederzeit
anlegen, leeren oder vollständig löschen, ebenso einzelne importierte Dateien.

---

## Was das Tool kann

| Anforderung | Umsetzung |
|---|---|
| Batch-Import | beliebig viele PDF/PNG/JPEG/TIFF **und** die Excel-Vorlagen (.xlsx/.xls), gemischt |
| Mehrseitige und gescannte PDFs | ja – jede Seite wird einzeln klassifiziert und verarbeitet |
| Alte **und** neue Fassung | automatische Erkennung, keine Abfrage nötig |
| Regelschemata in derselben Datei | werden erkannt, übersprungen und protokolliert |
| Summen-, Zwischensummen- und **Übertragszeilen** | werden erkannt und **nie mitgezählt**, bleiben aber zur Kontrolle sichtbar |
| Leerzeilen und Fußbereich | zählen nicht mit; der Datenbereich wird aus den Tabellenlinien hart begrenzt |
| Freitext in Zellen | wird als Anmerkung geführt, nie als Zahl addiert |
| **Automations-/Informationsschwerpunkt** | ASP/ISP wird aus Kopf-/Fußbereich, Anlagenangabe oder der Zeile selbst gelesen; jede Auswertung weist aus, wie viele Datenpunkte und Funktionen auf welchen ASP bzw. ISP entfallen |
| Kontrolle | die eigene Summe wird gegen die Zeile „Summe Funktionen" der Liste geprüft |
| **Nachweis** | zu jedem Wert und jeder Abweichung zeigt das Programm die markierte Fundstelle im Original |
| Vorschau & Korrektur | editierbare Tabelle je Datei; ganze Zeilen lassen sich ankreuzen und gesammelt löschen, erst danach speichern |
| Projekte | getrennte Datenbanken, jederzeit löschbar |
| Kostenschätzung | Einheitspreis je Funktionsspalte, sofortige Neuberechnung |
| Excel-Export | Funktionsliste im Original-Layout, Summen, Kostenblatt **mit Formeln**, Rohdaten, Dokumente, Fußnoten |

---

## Wie die Erkennung arbeitet

Das Tool nimmt je Datei den genauesten verfügbaren Weg:

| Quelle | Verfahren | Genauigkeit |
|---|---|---|
| `.xlsx` / `.xls` | direktes Auslesen der Zellen | exakt |
| PDF **mit Textebene** (aus Excel/CAD gedruckt) | Tabellenraster aus den Wortkoordinaten | exakt |
| Scan, Foto, PNG/JPEG, bildbasiertes PDF | Tabellenraster aus den Linien, dann Zelle-für-Zelle-Texterkennung | sehr hoch, wird gegen die Summenzeile geprüft |

Anker ist in allen Fällen das Kopfzeilenpaar **„Abschnitt"** und **„Spalte"**, das beide
VDI-Vorlagen enthalten. Daraus ergibt sich für jede Funktionsspalte eine eindeutige
Adresse (z. B. `6.13` = *Tarifabhängiges Schalten* in der alten, `2.2.10` =
*Frostschutzsteuerung* in der neuen Fassung).

Bei Scans, in denen die winzigen Kopfziffern nicht sicher lesbar sind, wird die Fassung
zusätzlich über den **Fingerabdruck der Vorlage** bestimmt – die Größen der
Abschnittsblöcke (alte Fassung 5-5-6-5-8-13-4-4, neue Fassung 1-4-2-5-2-12-9-5-4-5-1-6).
Passt der Fingerabdruck nicht, wird bewusst nichts geraten, sondern gemeldet.

Es werden **keine festen Positionen** verwendet: Spaltenmitten, Tabellenober- und
-unterkante sowie alle Toleranzen werden für jede Seite neu gemessen.

### Automations- und Informationsschwerpunkt (ASP/ISP)

In den Funktionslisten steht praktisch immer, zu welchem Schwerpunkt eine Seite oder
eine Zeile gehört. Das Tool sucht die Angabe in dieser Reihenfolge:

1. **die Zeile selbst** – z. B. `ASP01 Heizungstechnik 2.UG` auf einem Übersichtsblatt
   oder als Zwischenüberschrift; ab dort gilt sie auch für die folgenden Zeilen der Seite.
   Steht die Bezeichnung eine Zeile unter der Kennung, wird sie von dort übernommen;
2. **Kopf-/Fußbereich der Seite** – ein Feld `ASP:` / `Informationsschwerpunkt:`;
3. **die Anlagenangabe** des Blattes, wenn der Schwerpunkt dort mit drin steht
   (`Anlage: ASP03 Lüftungstechnik`);
4. nennt eine Seite nur an einer Stelle einen Schwerpunkt, gilt dieser für die ganze Seite.

Erkannt werden `ASP01`, `ASP 1`, `ASP-Nr. 7`, `ISP 02` und die ausgeschriebenen Wörter;
die Kennung wird auf die Form `ASP01` vereinheitlicht. Wörter wie „Aspiration" lösen
bewusst **keinen** Treffer aus. Wird nichts gefunden, meldet das Programm dies als
Hinweis – die Angabe lässt sich im Reiter *Prüfen & korrigieren* je Zeile nachtragen.
Summen- und Übertragszeilen werden keinem Schwerpunkt zugeschlagen: sie fassen die
Zeilen darüber zusammen.

> Bestehende Projekte bleiben erhalten – die Datenbank wird beim ersten Start ergänzt.
> Listen, die **vor** dieser Erweiterung importiert wurden, tragen noch keinen
> Schwerpunkt; sie werden als *ohne Zuordnung* geführt, bis sie einmal neu
> eingelesen werden.

> Für besonders schlechte Scans lässt sich zusätzlich ein lokales Vision-Modell
> (Ollama) einschalten. Das ist optional und für den normalen Betrieb nicht nötig.

---

## Start aus dem Quellcode (für Entwicklung)

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

streamlit run vdi3814\ui\app.py
```

Für gescannte Listen wird dabei ein installiertes Tesseract benötigt
([UB-Mannheim-Installer](https://github.com/UB-Mannheim/tesseract/wiki), Sprachpaket
Deutsch). In der fertigen EXE ist es bereits enthalten.

### Kommandozeile

Dieselbe EXE arbeitet mit Argumenten als Kommandozeilenwerkzeug:

```bat
VDI3814-DP-Checker.exe check                                :: Einrichtung prüfen
VDI3814-DP-Checker.exe projekt --neu "Neubau Musterstadt"   :: Projekt anlegen
VDI3814-DP-Checker.exe --projekt "Neubau Musterstadt" import C:\Listen
VDI3814-DP-Checker.exe --projekt "Neubau Musterstadt" list
VDI3814-DP-Checker.exe --projekt "Neubau Musterstadt" schwerpunkte  :: je ASP/ISP
VDI3814-DP-Checker.exe --projekt "Neubau Musterstadt" pruefen
VDI3814-DP-Checker.exe --projekt "Neubau Musterstadt" entfernen 3
VDI3814-DP-Checker.exe preise --setzen A_1_1=45 A_1_5=80
VDI3814-DP-Checker.exe --projekt "Neubau Musterstadt" export Auswertung.xlsx
```

Aus dem Quellcode entsprechend mit `python -m vdi3814.cli ...`.

### Beispiel: kompletter Ablauf

```bat
python sample_data\make_samples.py
python -m vdi3814.cli --projekt Demo import sample_data
python -m vdi3814.cli --projekt Demo preise --setzen A_1_5=80 A_1_1=45
python -m vdi3814.cli --projekt Demo export Auswertung.xlsx
```

Der Beispieldatensatz enthält eine Liste der alten Fassung (mit einer zweiten Seite
„Regelschema", die übersprungen werden muss), eine der neuen Fassung und ein
Übersichtsblatt, auf dem je Zeile ein Automationsschwerpunkt steht (ASP01, ASP02,
ASP03) – so, wie Plansätze ihre Zusammenstellung führen.

---

## Aufbau des Excel-Exports

| Blatt | Inhalt |
|---|---|
| **Übersicht** | Das VDI-3814-Blatt für die Kalkulation: **eine Zeile je importierter Liste** mit ihren Mengen je Funktionsspalte, rechts die Gesamtzahl der Funktionen dieser Liste. Darunter „Summe Funktionen" (Formel), eine Zeile für **Einheitspreise** und die daraus berechneten Kosten. Liegen beide Fassungen vor, gibt es je Fassung ein Blatt |
| **Schwerpunkte** | Eine Zeile je ASP/ISP: Kennung, Bezeichnung, Anzahl der Datenpunkte und Summe der Funktionen – die Sicht, mit der sich eine Summe auf die einzelnen Automationsschwerpunkte aufteilen lässt |
| **Mengen je Schwerpunkt** | Dasselbe im VDI-Aufbau: eine Zeile je ASP/ISP mit den Mengen je Funktionsspalte |
| **GA-Funktionsliste** | Dieselbe Struktur, aber Zeile für Zeile jeder einzelne Datenpunkt – mit dem ASP/ISP, zu dem er gehört. Die Grundlage der Mengen |
| **Spaltensummen** | Flache Liste aller Funktionsspalten, nach Menge oder Gruppe sortier- und filterbar |
| **Kostenschätzung** | Je Spalte: Menge (**Formel** auf die Übersicht), Einheitspreis, Kosten, Gesamtkosten |
| **Prüfung** | Je Seite der Abgleich mit der Zeile „Summe Funktionen" und jede nicht gezählte Zeile mit Begründung |
| **Rohdaten** | Jede einzelne Zelle mit Quelldatei, Seite, Zeilennummer, Datenpunkt und Wert |
| **Dokumente** | Importierte Dateien: Fassung, Verfahren, Metadaten, übersprungene Seiten, Hinweise |
| **Fußnoten** | Zählregeln aus der Kopfzeile inkl. der Spalten, auf die sie sich beziehen |

Alle Mengen und Kosten sind **Formeln**, keine festen Zahlen: Einheitspreise lassen
sich direkt in der Übersicht eintragen, alles darunter rechnet sofort mit. Die
Kostenschätzung verweist auf dieselbe Summenzelle – es gibt nur eine Quelle.

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
| `VDI_DATA_DIR` | `data` | Ablage der Projekte (`data/projekte/<Name>/`) |
| `VDI_DB_PATH` | `data/vdi3814.sqlite3` | einzelne Datenbank statt Projektordner |
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
* Werte, die im Scan anders gelesen wurden als in der Liste, fallen über die Prüfung
  gegen die Summenzeile auf und erscheinen im Reiter „Nachweis & Differenzen" mit
  Seitenausschnitt.
* Passt die erkannte Spaltenstruktur zu keiner hinterlegten Vorlage, wird das gemeldet,
  statt Werte falsch zuzuordnen.

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
  extract/             Excel, PDF-Textebene, Scan-OCR; wordgrid.py = gemeinsamer Kern
                       rowfilter.py = Summen-/Übertrags-/Leerzeilen-Erkennung
  vision/              OCR-Anbindung; optionaler Ollama-Client
  ui/app.py            Streamlit-Oberfläche
  pipeline.py          Ablaufsteuerung inkl. Fassungserkennung und Summenprüfung
  projects.py          Projektverwaltung (eigene Datenbank je Projekt)
  evidence.py          Nachweis: Fundstellen im Original markieren
  db.py                SQLite/SQLAlchemy
  aggregate.py         Auswertung und Drilldown
  costs.py             Einheitspreise und Kosten
  export_excel.py      Excel-Export mit Formeln
sample_data/           Beispieldatensatz (Generator + PDFs)
docs/                  Spalten-Mapping
tests/                 Pytest-Suite
launcher.py            Einstiegspunkt der EXE (Oberfläche bzw. Kommandozeile)
vdi3814.spec           PyInstaller-Beschreibung des Windows-Builds
.github/workflows/     Build der Windows-EXE inkl. Prüfung des Ergebnisses
```

### Windows-EXE selbst bauen

```bat
pip install -r requirements.txt pyinstaller
pyinstaller vdi3814.spec --noconfirm
dist\VDI3814-DP-Checker\VDI3814-DP-Checker.exe selbsttest
```

Der GitHub-Workflow macht genau das und prüft das Ergebnis anschließend mit
einem vollständigen Durchlauf (Profile laden → Selbsttest der Oberfläche →
Import → Datenbank → Excel-Export → Serverstart), bevor das ZIP veröffentlicht wird.
