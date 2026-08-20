"""Startprogramm der ausfuehrbaren Windows-Datei.

Doppelklick  -> lokale Oberflaeche im Browser (Streamlit auf 127.0.0.1)
Mit Argumenten -> Kommandozeile, z. B.:

    VDI3814-DP-Checker.exe import C:\\Projekte\\Listen
    VDI3814-DP-Checker.exe export Auswertung.xlsx

Die Oberflaeche ist ausschliesslich lokal erreichbar; es werden keinerlei
Daten nach aussen gesendet.
"""

from __future__ import annotations

import os
import socket
import sys
import threading
import time
import webbrowser
from pathlib import Path

APP_RELATIVE = Path("vdi3814") / "ui" / "app.py"


def base_dir() -> Path:
    """Verzeichnis mit den mitgelieferten Programmdateien."""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    return Path(__file__).resolve().parent


def free_port(preferred: int = 8501) -> int:
    for port in range(preferred, preferred + 50):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
            if probe.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return 0          # 0 = beliebiger freier Port


def open_browser_later(url: str, delay: float = 3.0) -> None:
    def worker() -> None:
        time.sleep(delay)
        webbrowser.open(url)

    threading.Thread(target=worker, daemon=True).start()


def run_ui() -> int:
    app_path = base_dir() / APP_RELATIVE
    if not app_path.exists():                       # pragma: no cover - nur im Build
        print(f"Oberflaeche nicht gefunden: {app_path}", file=sys.stderr)
        return 1

    port = free_port()
    url = f"http://127.0.0.1:{port}"
    print("VDI 3814 DP-Checker")
    print(f"  Oberflaeche: {url}")
    print("  Dieses Fenster bitte geoeffnet lassen - Beenden mit Strg+C.")
    open_browser_later(url)

    # Streamlit-Voreinstellungen: lokal, ohne Telemetrie, ohne eigenen Browserstart
    os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")
    os.environ.setdefault("STREAMLIT_SERVER_HEADLESS", "true")
    os.environ.setdefault("STREAMLIT_GLOBAL_DEVELOPMENT_MODE", "false")

    from streamlit.web import cli as streamlit_cli

    sys.argv = [
        "streamlit", "run", str(app_path),
        "--server.address=127.0.0.1",
        f"--server.port={port}",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
        "--global.developmentMode=false",
    ]
    return int(streamlit_cli.main() or 0)


def main() -> int:
    if len(sys.argv) > 1:
        from vdi3814.cli import main as cli_main

        return cli_main(sys.argv[1:])
    return run_ui()


if __name__ == "__main__":
    raise SystemExit(main())
