@echo off
cd /d "%~dp0"
set "VENV_PY=%~dp0..\.venv\Scripts\python.exe"

if exist "%VENV_PY%" (
    "%VENV_PY%" gui_reports.py
) else (
    echo Virtual environment not found: %VENV_PY%
    echo Create it in the PFE_Project folder first.
)
if errorlevel 1 pause