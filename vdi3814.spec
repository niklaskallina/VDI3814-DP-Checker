# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller-Beschreibung fuer die ausfuehrbare Datei.

Build (Windows):
    pip install -r requirements.txt pyinstaller
    pyinstaller vdi3814.spec --noconfirm

Ergebnis: dist/VDI3814-DP-Checker/VDI3814-DP-Checker.exe (Ordner-Build).
Ordner-Build statt Einzeldatei, weil eine Einzeldatei bei jedem Start
mehrere hundert MB entpacken muesste - der Ordner startet in Sekunden und
wird als ZIP weitergegeben.
"""

from PyInstaller.utils.hooks import collect_data_files, collect_submodules, copy_metadata

APP_NAME = "VDI3814-DP-Checker"

# Streamlit prueft zur Laufzeit die Paket-Metadaten seiner Abhaengigkeiten -
# ohne diese Kopien bricht der Start mit PackageNotFoundError ab.
METADATA_PACKAGES = [
    "streamlit", "altair", "pandas", "numpy", "pyarrow", "pydeck", "tornado",
    "click", "blinker", "cachetools", "protobuf", "packaging", "tenacity",
    "toml", "watchdog", "gitpython", "requests", "rich", "typing_extensions",
    "pillow", "openpyxl", "sqlalchemy", "pymupdf", "pyyaml", "xlrd",
]

datas = []
for package in METADATA_PACKAGES:
    try:
        datas += copy_metadata(package)
    except Exception:          # Paket nicht installiert -> ueberspringen
        pass

# Streamlit bringt Frontend-Dateien und zur Laufzeit gelesene .py-Dateien mit
datas += collect_data_files("streamlit", include_py_files=True)
datas += collect_data_files("altair")

# Eigene Dateien: Spaltenprofile, die als Skript ausgefuehrte Oberflaeche,
# Beispieldaten und Dokumentation
# Mitgelieferte Texterkennung (wird im CI-Lauf bereitgestellt), damit
# gescannte Listen ohne Fremdinstallation gelesen werden koennen.
import os
if os.path.isdir("tesseract"):
    datas += [("tesseract", "tesseract")]

datas += [
    ("vdi3814/profiles", "vdi3814/profiles"),
    ("vdi3814/ui/app.py", "vdi3814/ui"),
    ("sample_data/make_samples.py", "sample_data"),
    ("docs/spalten-mapping.md", "docs"),
    ("README.md", "."),
]

hiddenimports = collect_submodules("streamlit") + [
    "vdi3814.ui.app",
    "xlrd",              # nur fuer .xls, wird erst zur Laufzeit importiert
    "pytesseract",       # optionales OCR, ebenfalls verzoegert importiert
    "pymupdf",
    "PIL._tkinter_finder",
]

a = Analysis(
    ["launcher.py"],
    pathex=["."],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pytest", "formulas", "schedula", "matplotlib", "tkinter", "IPython", "notebook"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,          # Konsolenfenster zeigt die lokale Adresse und Fehler
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name=APP_NAME,
)
