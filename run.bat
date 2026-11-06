@echo off
cd /d "%~dp0"

if exist "%~dp0.venv\Scripts\python.exe" (
    "%~dp0.venv\Scripts\python.exe" ui_pyqt.py
) else (
    python ui_pyqt.py
)

if errorlevel 1 (
    echo.
    echo Application exited with an error.
    pause
)
