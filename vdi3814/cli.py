"""Kommandozeile - fuer Batch-Laeufe ohne Oberflaeche.

Beispiele:
    python -m vdi3814.cli check
    python -m vdi3814.cli import ./listen --db data/vdi3814.sqlite3
    python -m vdi3814.cli preise --setzen A_1_1=45 A_1_5=80
    python -m vdi3814.cli export auswertung.xlsx
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from sqlalchemy.orm import Session

from . import aggregate, db, evidence, export_excel, projects
from .config import SETTINGS, bundled_tesseract
from .costs import price_template
from .ingest.loader import collect_files
from .pipeline import process_files
from .profiles_loader import load_profiles
from .vision import ocr
from .vision.ollama_client import OllamaUnavailable, OllamaVisionBackend


def _engine(args):
    """Datenbank des gewaehlten Projekts (oder ein direkt angegebener Pfad)."""
    if getattr(args, "db", None):
        return db.make_engine(args.db)
    projekt = projects.resolve(getattr(args, "projekt", None))
    return db.make_engine(projekt.db_path)


def _backend(args):
    """Vision-Backend nur aufbauen, wenn es gebraucht werden koennte."""
    if getattr(args, "kein_modell", False):
        return None
    backend = OllamaVisionBackend(model=getattr(args, "modell", None) or SETTINGS.vision_model)
    try:
        backend.ensure_ready()
    except OllamaUnavailable as exc:
        print(f"[Hinweis] {exc}\n          Bilder/Scans koennen dadurch nicht ausgewertet werden; "
              f"Excel- und Text-PDFs funktionieren weiterhin.", file=sys.stderr)
        return None
    return backend


def cmd_check(args) -> int:
    print("VDI 3814 DP-Checker - Einrichtung")
    print()
    print("Ohne Zusatzsoftware auswertbar:")
    print("  Excel-Vorlagen (.xlsx/.xls)          : ja")
    print("  PDF mit Textebene (aus Excel/CAD)    : ja")
    print(f"  Gescannte Listen und Bilder          : "
          f"{'ja' if ocr.available() else 'NEIN - Texterkennung fehlt'}")
    print()
    print("Einzelheiten")
    print(f"  Projektordner    : {projects.projects_root()}")
    print(f"  Profile          : {', '.join(p.id for p in load_profiles())}")
    tesseract = bundled_tesseract()
    print(f"  Tesseract-OCR    : {'ja' if ocr.available() else 'nein'}"
          + (f" (mitgeliefert: {tesseract})" if tesseract else ""))
    if ocr.available():
        print(f"  OCR-Sprachen     : {', '.join(ocr.sprachen()) or 'unbekannt'}")
    print()
    print("Optional (nur fuer besonders schlechte Scans):")
    backend = OllamaVisionBackend()
    try:
        models = backend.list_models()
        print(f"  Lokales Vision-Modell: Ollama erreichbar ({len(models)} Modelle), "
              f"{SETTINGS.vision_model} {'vorhanden' if backend.available() else 'fehlt'}")
    except OllamaUnavailable:
        print("  Lokales Vision-Modell: nicht eingerichtet (wird nicht benoetigt)")
    return 0


def ui_app_path() -> Path:
    """Pfad zum Streamlit-Skript - auch im gepackten Zustand (.exe)."""
    if getattr(sys, "frozen", False):
        base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    else:
        base = Path(__file__).resolve().parents[1]
    return base / "vdi3814" / "ui" / "app.py"


def cmd_selftest(args) -> int:
    """Fuehrt das Oberflaechen-Skript einmal komplett aus.

    Wichtig fuer den EXE-Build: er prueft, dass Streamlit, die Profile und alle
    verzoegert importierten Module im gepackten Zustand tatsaechlich vorhanden
    sind - ein blosser HTTP-Test wuerde das nicht zeigen, weil das Skript erst
    beim Verbinden eines Browsers laeuft.
    """
    app_path = ui_app_path()
    if not app_path.exists():
        print(f"Oberflaechen-Skript nicht gefunden: {app_path}", file=sys.stderr)
        return 1
    from streamlit.testing.v1 import AppTest

    print(f"Oberflaeche wird geprueft: {app_path}")
    app = AppTest.from_file(str(app_path), default_timeout=180)
    app.run()
    if app.exception:
        for exception in app.exception:
            print(f"  FEHLER: {exception.value}", file=sys.stderr)
        return 1
    print(f"  ok - {len(app.tabs)} Reiter, {len(app.dataframe)} Tabellen aufgebaut")
    return 0


def cmd_import(args) -> int:
    files = collect_files(args.pfade)
    if not files:
        print("Keine unterstuetzten Dateien gefunden.", file=sys.stderr)
        return 1
    print(f"{len(files)} Datei(en) werden verarbeitet ...")
    backend = _backend(args)
    results = process_files(files, backend=backend, progress=lambda m: print("  ", m))

    engine = _engine(args)
    saved = 0
    with Session(engine) as session:
        for result in results:
            status = f"{result.file_name}: {len(result.data_rows())} Datenpunkte, " \
                     f"Summe {result.grand_total():g}, Fassung {result.fassung or '?'}"
            for warning in result.warnings:
                status += f"\n     ! {warning}"
            print("  ", status)
            if result.columns and not args.nur_pruefen:
                db.save_document(session, result)
                saved += 1
    if not args.nur_pruefen:
        ziel = args.db or projects.resolve(getattr(args, "projekt", None)).db_path
        print(f"{saved} Datei(en) gespeichert in {ziel}")
    return 0


def cmd_list(args) -> int:
    engine = _engine(args)
    with Session(engine) as session:
        frame = aggregate.documents_frame(session)
    if frame.empty:
        print("Datenbank ist leer.")
        return 0
    print(frame[["dokument_id", "datei", "fassung", "verfahren", "datenpunkte",
                 "summe_funktionen", "seiten_uebersprungen"]].to_string(index=False))
    return 0


def cmd_prices(args) -> int:
    engine = _engine(args)
    with Session(engine) as session:
        if args.setzen:
            updates = {}
            for item in args.setzen:
                key, _, value = item.partition("=")
                updates[key.strip()] = float(value.replace(",", "."))
            db.set_unit_prices(session, updates)
            print(f"{len(updates)} Einheitspreis(e) aktualisiert.")
        if args.vorlage:
            frame = price_template(args.vorlage, db.get_unit_prices(session))
            frame.to_csv(args.csv or f"einheitspreise_{args.vorlage}.csv", index=False, sep=";")
            print(f"Preisvorlage geschrieben: {args.csv or f'einheitspreise_{args.vorlage}.csv'}")
        if args.csv_import:
            import pandas as pd

            frame = pd.read_csv(args.csv_import, sep=";")
            db.set_unit_prices(session, dict(zip(frame["spalte_key"], frame["einheitspreis"])))
            print(f"{len(frame)} Einheitspreis(e) aus {args.csv_import} uebernommen.")
        prices = db.get_unit_prices(session)
    if prices and not args.vorlage:
        for key, value in sorted(prices.items()):
            print(f"  {key:10s} {value:10.2f} {SETTINGS.currency}")
    return 0


def cmd_export(args) -> int:
    engine = _engine(args)
    with Session(engine) as session:
        raw = aggregate.raw_dataframe(session)
        if raw.empty:
            print("Keine Daten zum Exportieren - bitte zuerst importieren.", file=sys.stderr)
            return 1
        summary = aggregate.column_summary(raw)
        path = export_excel.export_workbook(
            args.datei,
            summary=summary,
            raw=raw,
            documents=aggregate.documents_frame(session),
            footnotes=aggregate.footnotes_frame(session),
            prices=db.get_unit_prices(session),
            projects=aggregate.pivot_projects(raw),
            layouts={profile_id: aggregate.vdi_layout_frame(session, profile_id)
                     for profile_id in aggregate.profiles_in_use(session)},
        )
    print(f"Export geschrieben: {path}")
    return 0


def cmd_projects(args) -> int:
    if args.neu:
        projekt = projects.create(args.neu)
        print(f"Projekt angelegt: {projekt.name} ({projekt.path})")
    if args.loeschen:
        if projects.delete(args.loeschen):
            print(f"Projekt geloescht: {args.loeschen}")
        else:
            print(f"Projekt nicht gefunden: {args.loeschen}", file=sys.stderr)
            return 1
    if args.leeren:
        if projects.clear(args.leeren):
            print(f"Projekt geleert: {args.leeren}")
        else:
            print(f"Projekt nicht gefunden: {args.leeren}", file=sys.stderr)
            return 1
    for projekt in projects.list_projects():
        stand = projekt.changed_at.strftime("%d.%m.%Y %H:%M") if projekt.changed_at else "-"
        print(f"  {projekt.name:30s} {projekt.size_mb:8.2f} MB   {stand}   {projekt.path}")
    return 0


def cmd_delete(args) -> int:
    engine = _engine(args)
    with Session(engine) as session:
        for dokument_id in args.ids:
            db.delete_document(session, int(dokument_id))
    print(f"{len(args.ids)} Datei(en) aus der Datenbank entfernt.")
    return 0


def cmd_check_data(args) -> int:
    """Pruefbericht: Abweichungen und nicht gezaehlte Zeilen."""
    engine = _engine(args)
    with Session(engine) as session:
        befunde = evidence.all_findings(session)
    if not befunde:
        print("Keine Befunde - die Zaehlung deckt sich mit den Listen.")
        return 0
    for befund in befunde:
        werte = ", ".join(f"{k}: {v:g}" for k, v in befund.werte.items())
        print(f"[{befund.art:14s}] {befund.datei} S.{befund.anzeige_seite}: {befund.titel}"
              + (f"  ({werte})" if werte else ""))
    print(f"\n{len(befunde)} Befund(e). Details mit Seitenansicht in der Oberflaeche.")
    return 0


def cmd_profiles(args) -> int:
    for profile in load_profiles():
        print(f"{profile.id}  ({profile.fassung})  {profile.name}")
        print(f"   Spalten: {len(profile.columns)}   Fussnoten: {len(profile.footnotes)}")
        if args.ausfuehrlich:
            for column in profile.columns:
                print(f"     {column.address:8s} {column.key:10s} {column.group} / "
                      f"{column.subgroup} / {column.label}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="vdi3814", description="VDI 3814 DP-Checker")
    parser.add_argument("--projekt", help="Projektname (eigene Datenbank je Projekt)", default=None)
    parser.add_argument("--db", help="Pfad zur SQLite-Datenbank (statt --projekt)", default=None)
    sub = parser.add_subparsers(dest="befehl", required=True)

    check = sub.add_parser("check", help="Installation und Modellverfuegbarkeit pruefen")
    check.set_defaults(func=cmd_check)

    selftest = sub.add_parser("selbsttest", help="Oberflaeche einmal komplett ausfuehren")
    selftest.set_defaults(func=cmd_selftest)

    imp = sub.add_parser("import", help="Dateien/Ordner importieren")
    imp.add_argument("pfade", nargs="+", help="Dateien oder Ordner (PDF/PNG/JPEG/XLSX/XLS)")
    imp.add_argument("--modell", help="Ollama-Modell abweichend von der Voreinstellung")
    imp.add_argument("--kein-modell", action="store_true", dest="kein_modell",
                     help="ohne Vision-Modell arbeiten (nur Excel/Text-PDF)")
    imp.add_argument("--nur-pruefen", action="store_true", dest="nur_pruefen",
                     help="nur erkennen und anzeigen, nichts speichern")
    imp.set_defaults(func=cmd_import)

    listing = sub.add_parser("list", help="importierte Dateien anzeigen")
    listing.set_defaults(func=cmd_list)

    prices = sub.add_parser("preise", help="Einheitspreise pflegen")
    prices.add_argument("--setzen", nargs="*", default=[], metavar="KEY=WERT")
    prices.add_argument("--vorlage", help="CSV-Preisvorlage fuer ein Profil erzeugen")
    prices.add_argument("--csv", help="Zieldatei fuer --vorlage")
    prices.add_argument("--csv-import", dest="csv_import", help="Einheitspreise aus CSV uebernehmen")
    prices.set_defaults(func=cmd_prices)

    export = sub.add_parser("export", help="Excel-Auswertung schreiben")
    export.add_argument("datei", help="Zieldatei (.xlsx)")
    export.set_defaults(func=cmd_export)

    projekte = sub.add_parser("projekt", help="Projekte anlegen, leeren, loeschen, anzeigen")
    projekte.add_argument("--neu", help="neues Projekt anlegen")
    projekte.add_argument("--leeren", help="Daten eines Projekts entfernen, Projekt behalten")
    projekte.add_argument("--loeschen", help="Projekt samt Daten entfernen")
    projekte.set_defaults(func=cmd_projects)

    entfernen = sub.add_parser("entfernen", help="einzelne Dateien aus der Datenbank loeschen")
    entfernen.add_argument("ids", nargs="+", help="Dokument-IDs (siehe 'list')")
    entfernen.set_defaults(func=cmd_delete)

    pruefen = sub.add_parser("pruefen", help="Abweichungen und nicht gezaehlte Zeilen auflisten")
    pruefen.set_defaults(func=cmd_check_data)

    profiles = sub.add_parser("profile", help="hinterlegte Spalten-Profile anzeigen")
    profiles.add_argument("--ausfuehrlich", action="store_true")
    profiles.set_defaults(func=cmd_profiles)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    SETTINGS.ensure_dirs()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
