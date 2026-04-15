@echo off
title CDU Assessor Guide Builder
cd /d "%~dp0"

where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found.
    echo Install Python 3.9+ from https://www.python.org/downloads/
    echo Make sure to tick "Add Python to PATH" during installation.
    pause
    exit /b 1
)

python run.py
if errorlevel 1 pause
