@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Creating virtual environment...
  py -m venv .venv 2>nul || python -m venv .venv
)

if not exist ".venv\Scripts\python.exe" (
  echo Python not found. Please install Python 3.10+ first.
  pause
  exit /b 1
)

echo Installing dependencies...
".venv\Scripts\python.exe" -m pip install -r requirements.txt

if not exist ".domains_initialized" (
  echo First run detected. Importing domains_full.txt...
  ".venv\Scripts\python.exe" init_domains.py
  if %errorlevel% neq 0 (
    echo Domain import failed.
    pause
    exit /b 1
  )
  type nul > ".domains_initialized"
)

echo Starting app on http://127.0.0.1:5000
".venv\Scripts\python.exe" app.py
