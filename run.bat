@echo off
cd /d "%~dp0"

if exist "%~dp0.venv\Scripts\python.exe" (
    "%~dp0.venv\Scripts\python.exe" Ui_PyQt.py
) else (
    python Ui_PyQt.py
)

if errorlevel 1 (
    echo.
    echo Application exited with an error.
    pause
)
