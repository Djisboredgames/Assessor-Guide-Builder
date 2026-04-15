#!/usr/bin/env python3
"""
Assessor Guide Builder — Auto-installer and launcher.
Usage:  python run.py
        python3 run.py
"""
import sys
import os
import subprocess
import platform
import threading
import webbrowser
import time
from pathlib import Path

MIN_PYTHON = (3, 9)
PROJECT_DIR = Path(__file__).parent.resolve()
VENV_DIR = PROJECT_DIR / ".venv"
REQUIREMENTS = PROJECT_DIR / "requirements.txt"
APP = PROJECT_DIR / "app.py"
PORT = 8501

BANNER = """
╔══════════════════════════════════════════════╗
║       CDU Assessor Guide Builder  v1.0       ║
╚══════════════════════════════════════════════╝
"""


def log(msg: str, end: str = "\n"):
    print(msg, end=end, flush=True)


# ── Python version check ──────────────────────────────────────────────────────

def check_python_version():
    if sys.version_info < MIN_PYTHON:
        log(f"\nERROR: Python {MIN_PYTHON[0]}.{MIN_PYTHON[1]}+ is required.")
        log(f"       You have Python {sys.version_info.major}.{sys.version_info.minor}.")
        log("       Download from: https://www.python.org/downloads/")
        sys.exit(1)


# ── Venv helpers ──────────────────────────────────────────────────────────────

def _venv_python() -> Path:
    if platform.system() == "Windows":
        return VENV_DIR / "Scripts" / "python.exe"
    return VENV_DIR / "bin" / "python"


def _venv_streamlit() -> Path:
    if platform.system() == "Windows":
        return VENV_DIR / "Scripts" / "streamlit.exe"
    return VENV_DIR / "bin" / "streamlit"


def _deps_ok() -> bool:
    """Return True if the core packages are already importable from the venv."""
    python = _venv_python()
    if not python.exists():
        return False
    result = subprocess.run(
        [str(python), "-c", "import streamlit, docx, bs4, requests, lxml"],
        capture_output=True,
        timeout=15,
    )
    return result.returncode == 0


def setup():
    """Create venv and install requirements, only if needed."""
    if not VENV_DIR.exists():
        log("  Creating virtual environment... ", end="")
        subprocess.run(
            [sys.executable, "-m", "venv", str(VENV_DIR)],
            check=True,
            capture_output=True,
        )
        log("done.")

    if not _deps_ok():
        log("  Installing dependencies (first-run, ~1–2 minutes)...")
        python = _venv_python()
        subprocess.run(
            [str(python), "-m", "pip", "install", "--quiet", "--upgrade", "pip"],
            check=True,
        )
        subprocess.run(
            [str(python), "-m", "pip", "install", "--quiet", "-r", str(REQUIREMENTS)],
            check=True,
        )
        log("  Dependencies installed successfully.")
    else:
        log("  Dependencies OK.")


# ── Launch ────────────────────────────────────────────────────────────────────

def _open_browser():
    time.sleep(2.5)
    webbrowser.open(f"http://localhost:{PORT}")


def launch():
    streamlit = _venv_streamlit()
    log(f"\n  Starting at http://localhost:{PORT}")
    log("  Press Ctrl+C to stop.\n")
    threading.Thread(target=_open_browser, daemon=True).start()
    subprocess.run([
        str(streamlit), "run", str(APP),
        f"--server.port={PORT}",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
    ])


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print(BANNER)
    check_python_version()
    try:
        setup()
        launch()
    except subprocess.CalledProcessError as e:
        log(f"\nSetup failed: {e}")
        log("Try manually: pip install -r requirements.txt && streamlit run app.py")
        sys.exit(1)
    except KeyboardInterrupt:
        log("\n\nStopped.")
