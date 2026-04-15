#!/usr/bin/env bash
# Assessor Guide Builder — macOS launcher
# Double-click this file to start the app.
cd "$(dirname "$0")"

if ! command -v python3 &>/dev/null; then
    echo "ERROR: python3 not found."
    echo "Install Python 3.9+ from https://www.python.org/downloads/"
    read -p "Press Enter to close..."
    exit 1
fi

python3 run.py
