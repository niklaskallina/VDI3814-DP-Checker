# Spalten-Mapping der VDI-3814-Funktionslisten

Automatisch erzeugt aus `vdi3814/profiles/*.yaml` (`python tools/gen_mapping_doc.py`).

Der Schlüssel (`Key`) ist die interne, fassungsübergreifend stabile Kennung – er wird in der Datenbank, in der Kostenschätzung und im Excel-Export verwendet. Die Adresse `Abschnitt.Spalte` entspricht der Nummerierung im Tabellenkopf der Vorlage und ist der primäre Anker der automatischen Erkennung.

## VDI 3814 Blatt 1 (alte Fassung) / DIN EN ISO 16484-3

- Profil-ID: `vdi3814_alt`
- Fassung: `alt`
- Funktionsspalten: 50
- Quelle: `vdi3814_alt.yaml`

| Adresse | Key | Gruppe | Untergruppe | Bezeichnung | Kürzel | Aggregation | Fußnoten |
|---|---|---|---|---|---|---|---|
| 1.1 | `A_1_1` | Ein-/Ausgabefunktionen | 1 Physikalisch | Binäre Ausgabe Schalten/Stellen | BA | EIN_AUSGABE | 1 |
| 1.2 | `A_1_2` | Ein-/Ausgabefunktionen | 1 Physikalisch | Analoge Ausgabe Stellen | AA | EIN_AUSGABE |  |
| 1.3 | `A_1_3` | Ein-/Ausgabefunktionen | 1 Physikalisch | Binäre Eingabe Melden | BE-M | EIN_AUSGABE |  |
| 1.4 | `A_1_4` | Ein-/Ausgabefunktionen | 1 Physikalisch | Binäre Eingabe Zählen | BE-Z | EIN_AUSGABE |  |
| 1.5 | `A_1_5` | Ein-/Ausgabefunktionen | 1 Physikalisch | Analoge Eingabe Messen | AE | EIN_AUSGABE | 2 |
| 2.1 | `A_2_1` | Ein-/Ausgabefunktionen | 2 Gemeinsam (kommunikativ) | Binärer Ausgabewert, Schalten | BAW | EIN_AUSGABE |  |
| 2.2 | `A_2_2` | Ein-/Ausgabefunktionen | 2 Gemeinsam (kommunikativ) | Analoger Ausgabewert, Stellen/Sollwert | AAW | EIN_AUSGABE |  |
| 2.3 | `A_2_3` | Ein-/Ausgabefunktionen | 2 Gemeinsam (kommunikativ) | Binärer Eingabewert, Zustand | BEW | EIN_AUSGABE |  |
| 2.4 | `A_2_4` | Ein-/Ausgabefunktionen | 2 Gemeinsam (kommunikativ) | Zählwerteingabe | ZWE | EIN_AUSGABE |  |
| 2.5 | `A_2_5` | Ein-/Ausgabefunktionen | 2 Gemeinsam (kommunikativ) | Analoger Eingabewert, Messen | AEW | EIN_AUSGABE |  |
| 3.1 | `A_3_1` | Verarbeitungsfunktionen | 3 Überwachen | Grenzwert fest | GWF | VERARBEITUNG |  |
| 3.2 | `A_3_2` | Verarbeitungsfunktionen | 3 Überwachen | Grenzwert gleitend | GWG | VERARBEITUNG |  |
| 3.3 | `A_3_3` | Verarbeitungsfunktionen | 3 Überwachen | Betriebsstunden-Erfassung | BSE | VERARBEITUNG |  |
| 3.4 | `A_3_4` | Verarbeitungsfunktionen | 3 Überwachen | Ereigniszählung | EZ | VERARBEITUNG |  |
| 3.5 | `A_3_5` | Verarbeitungsfunktionen | 3 Überwachen | Befehlsausführkontrolle | BAK | VERARBEITUNG |  |
| 3.6 | `A_3_6` | Verarbeitungsfunktionen | 3 Überwachen | Meldungsbearbeitung | MB | VERARBEITUNG | 4 |
| 4.1 | `A_4_1` | Verarbeitungsfunktionen | 4 Steuern | Anlagensteuerung | ANS | VERARBEITUNG |  |
| 4.2 | `A_4_2` | Verarbeitungsfunktionen | 4 Steuern | Motorsteuerung | MOS | VERARBEITUNG |  |
| 4.3 | `A_4_3` | Verarbeitungsfunktionen | 4 Steuern | Umschaltung | UMS | VERARBEITUNG | 5 |
| 4.4 | `A_4_4` | Verarbeitungsfunktionen | 4 Steuern | Folgesteuerung | FST | VERARBEITUNG | 5 |
| 4.5 | `A_4_5` | Verarbeitungsfunktionen | 4 Steuern | Sicherheits-/Frostschutzsteuerung | SFS | VERARBEITUNG |  |
| 5.1 | `A_5_1` | Verarbeitungsfunktionen | 5 Regeln | P-Regelung | P | VERARBEITUNG |  |
| 5.2 | `A_5_2` | Verarbeitungsfunktionen | 5 Regeln | PI-/PID-Regelung | PID | VERARBEITUNG |  |
| 5.3 | `A_5_3` | Verarbeitungsfunktionen | 5 Regeln | Sollwertführung/-kennlinie | SWF | VERARBEITUNG |  |
| 5.4 | `A_5_4` | Verarbeitungsfunktionen | 5 Regeln | Stellausgabe stetig | SAS | VERARBEITUNG |  |
| 5.5 | `A_5_5` | Verarbeitungsfunktionen | 5 Regeln | Stellausgabe 2-Punkt | SA2 | VERARBEITUNG | 6 |
| 5.6 | `A_5_6` | Verarbeitungsfunktionen | 5 Regeln | Stellausgabe Pulsweitenmodulation | PWM | VERARBEITUNG |  |
| 5.7 | `A_5_7` | Verarbeitungsfunktionen | 5 Regeln | Begrenzung Sollwert/Stellgröße | BEG | VERARBEITUNG |  |
| 5.8 | `A_5_8` | Verarbeitungsfunktionen | 5 Regeln | Parameterumschaltung | PUM | VERARBEITUNG |  |
| 6.1 | `A_6_1` | Verarbeitungsfunktionen | 6 Rechnen/Optimieren | h,x geführte Strategie | HXS | VERARBEITUNG | 7 |
| 6.2 | `A_6_2` | Verarbeitungsfunktionen | 6 Rechnen/Optimieren | Arithmetische Berechnung | ARB | VERARBEITUNG | 7 |
| 6.3 | `A_6_3` | Verarbeitungsfunktionen | 6 Rechnen/Optimieren | Ereignisabhängiges Schalten | EAS | VERARBEITUNG |  |
| 6.4 | `A_6_4` | Verarbeitungsfunktionen | 6 Rechnen/Optimieren | Zeitabhängiges Schalten | ZAS | VERARBEITUNG |  |
| 6.5 | `A_6_5` | Verarbeitungsfunktionen | 6 Rechnen/Optimieren | Gleitendes Ein-/Ausschalten | GEA | VERARBEITUNG |  |
| 6.6 | `A_6_6` | Verarbeitungsfunktionen | 6 Rechnen/Optimieren | Zyklisches Schalten | ZYS | VERARBEITUNG |  |
| 6.7 | `A_6_7` | Verarbeitungsfunktionen | 6 Rechnen/Optimieren | Nachtkühlbetrieb | NKB | VERARBEITUNG |  |
| 6.8 | `A_6_8` | Verarbeitungsfunktionen | 6 Rechnen/Optimieren | Raumtemperaturbegrenzung | RTB | VERARBEITUNG |  |
| 6.9 | `A_6_9` | Verarbeitungsfunktionen | 6 Rechnen/Optimieren | Energierückgewinnung | ERG | VERARBEITUNG | 7 |
| 6.10 | `A_6_10` | Verarbeitungsfunktionen | 6 Rechnen/Optimieren | Netzersatzbetrieb | NEB | VERARBEITUNG |  |
| 6.11 | `A_6_11` | Verarbeitungsfunktionen | 6 Rechnen/Optimieren | Netzwiederkehrprogramm | NWP | VERARBEITUNG |  |
| 6.12 | `A_6_12` | Verarbeitungsfunktionen | 6 Rechnen/Optimieren | Höchstlastbegrenzung | HLB | VERARBEITUNG |  |
| 6.13 | `A_6_13` | Verarbeitungsfunktionen | 6 Rechnen/Optimieren | Tarifabhängiges Schalten | TAS | VERARBEITUNG |  |
| 7.1 | `A_7_1` | Managementfunktionen | 7 Managementfunktionen | Ein-/Ausgabe Objekttyp | EAO | MANAGEMENT | 9 |
| 7.2 | `A_7_2` | Managementfunktionen | 7 Managementfunktionen | Komplexer Objekttyp | KOT | MANAGEMENT | 8, 9 |
| 7.3 | `A_7_3` | Managementfunktionen | 7 Managementfunktionen | Ereignis-Langzeitspeicherung | ELS | MANAGEMENT |  |
| 7.4 | `A_7_4` | Managementfunktionen | 7 Managementfunktionen | Historisierung in Datenbank | HDB | MANAGEMENT |  |
| 8.1 | `A_8_1` | Bedienfunktionen | 8 Bedienfunktionen | Grafik/Anlagenbild | GAB | BEDIENUNG |  |
| 8.2 | `A_8_2` | Bedienfunktionen | 8 Bedienfunktionen | Dynamische Einblendung | DYE | BEDIENUNG |  |
| 8.3 | `A_8_3` | Bedienfunktionen | 8 Bedienfunktionen | Ereignis-Anweisungstext | EAT | BEDIENUNG |  |
| 8.4 | `A_8_4` | Bedienfunktionen | 8 Bedienfunktionen | Nachricht an externe Stelle | NES | BEDIENUNG |  |

**Freitextspalten (werden nie gezählt):**

- 9 `NOTE_BEM` – Bemerkungen

**Fußnoten der Kopfzeile (Zählregeln):**

- **1)** Dauerbefehl: z. B. 0,I,II = 2 BA / Impulsbefehl: z. B. 0,I,II = 3 BA / Stellbefehl: z. B. Zu-0-Auf = 2 BA / Pulsweitenmod. = 1 BA
- **2)** aktiv oder passiv
- **3)** nur gemeinsame, kommunikative Datenpunkte von Fremdsystemen für interoperable Funktionen
- **4)** pro Eingangs-Benutzeradresse zum a) Zusammenfassen, b) Verzögern und c) Unterdrücken von Meldungen
- **5)** pro Ausgangs-Benutzeradresse
- **6)** Stellausgabe: z. B. 3-Punkt = 2 x 2-Punkt, Pulsweitenmod. = 1 BA
- **7)** pro Eingangs-Benutzeradresse
- **8)** z. B. Gerätestatus, Zeitschalttab., Sicherheitspkt., Regler, Datei (DIN EN ISO 16484-5)
- **9)** Falls erforderlich sind bei gemeinsamen (shared) Datenpunkten die Funktionen im Client mit "A" und die im Server mit "B" zu kennzeichnen (siehe BIBBs).

## VDI 3814 (neue Fassung, Vorlage 2022-07 / Blatt 4.3)

- Profil-ID: `vdi3814_neu`
- Fassung: `neu`
- Funktionsspalten: 56
- Quelle: `vdi3814_neu.yaml`

| Adresse | Key | Gruppe | Untergruppe | Bezeichnung | Kürzel | Aggregation | Fußnoten |
|---|---|---|---|---|---|---|---|
| 0.1 | `N_0_1` | 0. Integration |  | Integrationsart der Datenpunkte oder Grafiktyp | I | INTEGRATION |  |
| 1.1.1 | `N_1_1_1` | 1. Ein-/Ausgabefunktionen | 1.1 Physikalische | Analoge Eingabe (AI) | AI | EIN_AUSGABE |  |
| 1.1.2 | `N_1_1_2` | 1. Ein-/Ausgabefunktionen | 1.1 Physikalische | Binäre Eingabe (BI) | BI | EIN_AUSGABE |  |
| 1.1.3 | `N_1_1_3` | 1. Ein-/Ausgabefunktionen | 1.1 Physikalische | Analoge Ausgabe (AO) | AO | EIN_AUSGABE |  |
| 1.1.4 | `N_1_1_4` | 1. Ein-/Ausgabefunktionen | 1.1 Physikalische | Binäre Ausgabe (BO) | BO | EIN_AUSGABE |  |
| 1.2.1 | `N_1_2_1` | 1. Ein-/Ausgabefunktionen | 1.2 Werte | Analogwert (AV) | AV | EIN_AUSGABE |  |
| 1.2.2 | `N_1_2_2` | 1. Ein-/Ausgabefunktionen | 1.2 Werte | Binärwert(e) (BV/MV) | BV | EIN_AUSGABE |  |
| 1.3.1 | `N_1_3_1` | 1. Ein-/Ausgabefunktionen | 1.3 Komplexe | Zeitplan (SCH) | SCH | EIN_AUSGABE |  |
| 1.3.2 | `N_1_3_2` | 1. Ein-/Ausgabefunktionen | 1.3 Komplexe | Kalender (CAL) | CAL | EIN_AUSGABE |  |
| 1.3.3 | `N_1_3_3` | 1. Ein-/Ausgabefunktionen | 1.3 Komplexe | Alarm-/Ereignismeldung (NC) | NC | EIN_AUSGABE |  |
| 1.3.4 | `N_1_3_4` | 1. Ein-/Ausgabefunktionen | 1.3 Komplexe | Datenaufzeichnung (LOG) | LOG | EIN_AUSGABE |  |
| 1.3.5 | `N_1_3_5` | 1. Ein-/Ausgabefunktionen | 1.3 Komplexe | Sonstige komplexe Objekte (KO) | KO | EIN_AUSGABE |  |
| 2.1.1 | `N_2_1_1` | 2. Anwendungsfunktionen | 2.1 Logik | Zeiten (TP, TON, TOFF) | ZEIT | ANWENDUNG |  |
| 2.1.2 | `N_2_1_2` | 2. Anwendungsfunktionen | 2.1 Logik | Arithmetische Berechnung | ARB | ANWENDUNG |  |
| 2.2.1 | `N_2_2_1` | 2. Anwendungsfunktionen | 2.2 Überwachung und Steuerung | Grenzwertüberwachung | GWU | ANWENDUNG |  |
| 2.2.2 | `N_2_2_2` | 2. Anwendungsfunktionen | 2.2 Überwachung und Steuerung | Betriebsstundenüberwachung | BSU | ANWENDUNG |  |
| 2.2.3 | `N_2_2_3` | 2. Anwendungsfunktionen | 2.2 Überwachung und Steuerung | Ereignisüberwachung (Zählung) | EZU | ANWENDUNG |  |
| 2.2.4 | `N_2_2_4` | 2. Anwendungsfunktionen | 2.2 Überwachung und Steuerung | Befehlsausführüberwachung | BAU | ANWENDUNG |  |
| 2.2.5 | `N_2_2_5` | 2. Anwendungsfunktionen | 2.2 Überwachung und Steuerung | Anlagen-/Gerätestatus | AGS | ANWENDUNG |  |
| 2.2.6 | `N_2_2_6` | 2. Anwendungsfunktionen | 2.2 Überwachung und Steuerung | Motorsteuerung | MOS | ANWENDUNG |  |
| 2.2.7 | `N_2_2_7` | 2. Anwendungsfunktionen | 2.2 Überwachung und Steuerung | Blockierschutz | BLS | ANWENDUNG |  |
| 2.2.8 | `N_2_2_8` | 2. Anwendungsfunktionen | 2.2 Überwachung und Steuerung | Umschaltung (Analogwert) | UMS | ANWENDUNG |  |
| 2.2.9 | `N_2_2_9` | 2. Anwendungsfunktionen | 2.2 Überwachung und Steuerung | Folgesteuerung | FST | ANWENDUNG |  |
| 2.2.10 | `N_2_2_10` | 2. Anwendungsfunktionen | 2.2 Überwachung und Steuerung | Frostschutzsteuerung | FRS | ANWENDUNG |  |
| 2.2.11 | `N_2_2_11` | 2. Anwendungsfunktionen | 2.2 Überwachung und Steuerung | Sicherheitssteuerung | SIS | ANWENDUNG |  |
| 2.2.12 | `N_2_2_12` | 2. Anwendungsfunktionen | 2.2 Überwachung und Steuerung | Prioritätssteuerung | PRS | ANWENDUNG |  |
| 2.3.1 | `N_2_3_1` | 2. Anwendungsfunktionen | 2.3 Regelung | P-/ PI- / PID-Regler | REG | ANWENDUNG |  |
| 2.3.2 | `N_2_3_2` | 2. Anwendungsfunktionen | 2.3 Regelung | Sollwertführung / -kennlinie | SWF | ANWENDUNG |  |
| 2.3.3 | `N_2_3_3` | 2. Anwendungsfunktionen | 2.3 Regelung | Stellausgabe stetig | SAS | ANWENDUNG |  |
| 2.3.4 | `N_2_3_4` | 2. Anwendungsfunktionen | 2.3 Regelung | Stellausgabe 2-Punkt | SA2 | ANWENDUNG |  |
| 2.3.5 | `N_2_3_5` | 2. Anwendungsfunktionen | 2.3 Regelung | Stellausgabe 3-Punkt | SA3 | ANWENDUNG |  |
| 2.3.6 | `N_2_3_6` | 2. Anwendungsfunktionen | 2.3 Regelung | Stellausgabe 3-Punkt mit variabler Impulslänge | SA3V | ANWENDUNG |  |
| 2.3.7 | `N_2_3_7` | 2. Anwendungsfunktionen | 2.3 Regelung | Stellausgabe Pulsweitenmodulation | PWM | ANWENDUNG |  |
| 2.3.8 | `N_2_3_8` | 2. Anwendungsfunktionen | 2.3 Regelung | Begrenzung Sollwert/Stellwert | BEG | ANWENDUNG |  |
| 2.3.9 | `N_2_3_9` | 2. Anwendungsfunktionen | 2.3 Regelung | Parameterumschaltung | PUM | ANWENDUNG |  |
| 2.4.1 | `N_2_4_1` | 2. Anwendungsfunktionen | 2.4 Optimierung | Energieniveau | ENI | ANWENDUNG |  |
| 2.4.2 | `N_2_4_2` | 2. Anwendungsfunktionen | 2.4 Optimierung | Energierückgewinnung | ERG | ANWENDUNG |  |
| 2.4.3 | `N_2_4_3` | 2. Anwendungsfunktionen | 2.4 Optimierung | h,x-Sollwertführung | HXS | ANWENDUNG |  |
| 2.4.4 | `N_2_4_4` | 2. Anwendungsfunktionen | 2.4 Optimierung | Schaltzeitpunktoptimierung | SZO | ANWENDUNG |  |
| 2.4.5 | `N_2_4_5` | 2. Anwendungsfunktionen | 2.4 Optimierung | Nachtkühlung | NKB | ANWENDUNG |  |
| 2.5.1 | `N_2_5_1` | 2. Anwendungsfunktionen | 2.5 Beleuchtung | Lichtsteuerung | LIS | ANWENDUNG |  |
| 2.5.2 | `N_2_5_2` | 2. Anwendungsfunktionen | 2.5 Beleuchtung | Treppenlichtschaltung | TLS | ANWENDUNG |  |
| 2.5.3 | `N_2_5_3` | 2. Anwendungsfunktionen | 2.5 Beleuchtung | Tageslichtschaltung | TLA | ANWENDUNG |  |
| 2.5.4 | `N_2_5_4` | 2. Anwendungsfunktionen | 2.5 Beleuchtung | Konstantlichtregelung | KLR | ANWENDUNG |  |
| 2.6.1 | `N_2_6_1` | 2. Anwendungsfunktionen | 2.6 Sonnenschutz | Sonnenschutz stellen | SST | ANWENDUNG |  |
| 2.6.2 | `N_2_6_2` | 2. Anwendungsfunktionen | 2.6 Sonnenschutz | Sonnen-/Dämmerungsautomatik | SDA | ANWENDUNG |  |
| 2.6.3 | `N_2_6_3` | 2. Anwendungsfunktionen | 2.6 Sonnenschutz | Thermoautomatik | THA | ANWENDUNG |  |
| 2.6.4 | `N_2_6_4` | 2. Anwendungsfunktionen | 2.6 Sonnenschutz | Witterungsschutz | WIS | ANWENDUNG |  |
| 2.6.5 | `N_2_6_5` | 2. Anwendungsfunktionen | 2.6 Sonnenschutz | Sonnenstandsberechnung | SSB | ANWENDUNG |  |
| 2.7.1 | `N_2_7_1` | 2. Anwendungsfunktionen | 2.7 Sonstige | Sonstige Anwendungsfunktionen | SAF | ANWENDUNG |  |
| 3.1.1 | `N_3_1_1` | 3. B-/A-Funktionen | 3.1 Bedienung/Anzeige | Automationsfunktionen und Parameter | AFP | BEDIENUNG |  |
| 3.1.2 | `N_3_1_2` | 3. B-/A-Funktionen | 3.1 Bedienung/Anzeige | Grafik | GRA | BEDIENUNG |  |
| 3.1.3 | `N_3_1_3` | 3. B-/A-Funktionen | 3.1 Bedienung/Anzeige | Dynamisierung | DYN | BEDIENUNG |  |
| 3.1.4 | `N_3_1_4` | 3. B-/A-Funktionen | 3.1 Bedienung/Anzeige | Handlungsanweisung | HAW | BEDIENUNG |  |
| 3.1.5 | `N_3_1_5` | 3. B-/A-Funktionen | 3.1 Bedienung/Anzeige | Nachricht an externe Stellen | NES | BEDIENUNG |  |
| 3.1.6 | `N_3_1_6` | 3. B-/A-Funktionen | 3.1 Bedienung/Anzeige | Historisierung in Datenbank | HDB | BEDIENUNG |  |

**Freitextspalten (werden nie gezählt):**

- 4.1 `NOTE_BEM` – Bemerkungen und Referenzierungen
