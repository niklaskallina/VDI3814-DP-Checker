"""Erzeugt docs/spalten-mapping.md aus den Profil-YAMLs.

So bleibt die Dokumentation der Spalten-Mappings automatisch synchron zu den
tatsaechlich verwendeten Profilen.

Aufruf:  python tools/gen_mapping_doc.py
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from vdi3814.profiles_loader import load_profiles      # noqa: E402

OUT = Path(__file__).resolve().parents[1] / "docs" / "spalten-mapping.md"


def main() -> None:
    lines = [
        "# Spalten-Mapping der VDI-3814-Funktionslisten",
        "",
        "Automatisch erzeugt aus `vdi3814/profiles/*.yaml` "
        "(`python tools/gen_mapping_doc.py`).",
        "",
        "Der Schlüssel (`Key`) ist die interne, fassungsübergreifend stabile Kennung – "
        "er wird in der Datenbank, in der Kostenschätzung und im Excel-Export verwendet. "
        "Die Adresse `Abschnitt.Spalte` entspricht der Nummerierung im Tabellenkopf der "
        "Vorlage und ist der primäre Anker der automatischen Erkennung.",
        "",
    ]
    for profile in sorted(load_profiles(), key=lambda p: p.fassung):
        lines += [
            f"## {profile.name}",
            "",
            f"- Profil-ID: `{profile.id}`",
            f"- Fassung: `{profile.fassung}`",
            f"- Funktionsspalten: {len(profile.columns)}",
            f"- Quelle: `{Path(profile.source_file).name}`",
            "",
            "| Adresse | Key | Gruppe | Untergruppe | Bezeichnung | Kürzel | Aggregation | Fußnoten |",
            "|---|---|---|---|---|---|---|---|",
        ]
        for column in profile.columns:
            lines.append(
                f"| {column.address} | `{column.key}` | {column.group} | {column.subgroup} | "
                f"{column.label} | {column.abbrev} | {column.aggregate_group} | "
                f"{', '.join(column.footnotes)} |"
            )
        lines.append("")
        if profile.note_columns:
            lines += ["**Freitextspalten (werden nie gezählt):**", ""]
            for column in profile.note_columns:
                lines.append(f"- {column.address} `{column.key}` – {column.label}")
            lines.append("")
        if profile.footnotes:
            lines += ["**Fußnoten der Kopfzeile (Zählregeln):**", ""]
            for marker, text in profile.footnotes:
                lines.append(f"- **{marker})** {text}")
            lines.append("")
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text("\n".join(lines), encoding="utf-8")
    print(f"geschrieben: {OUT}")


if __name__ == "__main__":
    main()
