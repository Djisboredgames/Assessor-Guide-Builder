#!/usr/bin/env bash
# Assessor Guide Builder — launcher for macOS / Linux
# Usage: ./start.sh   (or double-click in file manager)
set -e
cd "$(dirname "$0")"

if ! command -v python3 &>/dev/null; then
    echo "ERROR: python3 not found."
    echo "Install Python 3.9+ from https://www.python.org/downloads/"
    exit 1
fi

python3 run.py
