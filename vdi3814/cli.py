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

from . import aggregate, db, export_excel
from .config import SETTINGS
from .costs import price_template
from .ingest.loader import collect_files
from .pipeline import process_files
from .profiles_loader import load_profiles
from .vision import ocr
from .vision.ollama_client import OllamaUnavailable, OllamaVisionBackend


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
    print("Konfiguration")
    print(f"  Datenbank        : {SETTINGS.db_path}")
    print(f"  Profile          : {', '.join(p.id for p in load_profiles())}")
    print(f"  Ollama-Host      : {SETTINGS.ollama_host}")
    print(f"  Vision-Modell    : {SETTINGS.vision_model}")
    backend = OllamaVisionBackend()
    try:
        models = backend.list_models()
        ready = backend.available()
        print(f"  Ollama erreichbar: ja ({len(models)} Modelle installiert)")
        print(f"  Modell vorhanden : {'ja' if ready else 'NEIN - bitte `ollama pull ' + SETTINGS.vision_model + '`'}")
    except OllamaUnavailable as exc:
        print(f"  Ollama erreichbar: NEIN ({exc})")
    print(f"  Tesseract-OCR    : {'ja' if ocr.available() else 'nein (optional)'}")
    return 0


def cmd_import(args) -> int:
    files = collect_files(args.pfade)
    if not files:
        print("Keine unterstuetzten Dateien gefunden.", file=sys.stderr)
        return 1
    print(f"{len(files)} Datei(en) werden verarbeitet ...")
    backend = _backend(args)
    results = process_files(files, backend=backend, progress=lambda m: print("  ", m))

    engine = db.make_engine(args.db)
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
        print(f"{saved} Datei(en) in {args.db or SETTINGS.db_path} gespeichert.")
    return 0


def cmd_list(args) -> int:
    engine = db.make_engine(args.db)
    with Session(engine) as session:
        frame = aggregate.documents_frame(session)
    if frame.empty:
        print("Datenbank ist leer.")
        return 0
    print(frame[["dokument_id", "datei", "fassung", "verfahren", "datenpunkte",
                 "summe_funktionen", "seiten_uebersprungen"]].to_string(index=False))
    return 0


def cmd_prices(args) -> int:
    engine = db.make_engine(args.db)
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
    engine = db.make_engine(args.db)
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
        )
    print(f"Export geschrieben: {path}")
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
    parser.add_argument("--db", help="Pfad zur SQLite-Datenbank", default=None)
    sub = parser.add_subparsers(dest="befehl", required=True)

    check = sub.add_parser("check", help="Installation und Modellverfuegbarkeit pruefen")
    check.set_defaults(func=cmd_check)

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
