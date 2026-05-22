@echo off
setlocal

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment...
    python -m venv .venv
)

echo Installing dependencies...
".venv\Scripts\python.exe" -m pip install -r requirements.txt

echo.
echo PDF Office 2007 Converter is starting...
echo Config file: %cd%\.env
echo Default URL: http://127.0.0.1:7633
echo.

".venv\Scripts\python.exe" app.py
