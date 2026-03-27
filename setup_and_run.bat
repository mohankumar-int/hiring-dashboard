@echo off
:: ─────────────────────────────────────────────────────────────
:: Intuit Hiring Dashboard — Windows setup & launcher
:: Double-click or run from Command Prompt to start the app
:: ─────────────────────────────────────────────────────────────

echo.
echo   Intuit Hiring Dashboard
echo ────────────────────────────

:: ── Check Python ─────────────────────────────────────────────
python --version >nul 2>&1
IF ERRORLEVEL 1 (
  echo [ERROR] Python 3 not found.
  echo         Install from https://www.python.org/downloads/
  echo         Make sure to check "Add Python to PATH" during install.
  pause
  exit /b 1
)
echo [OK] Python found

:: ── Create virtual environment if needed ─────────────────────
IF NOT EXIST ".venv\" (
  echo [INFO] Creating virtual environment...
  python -m venv .venv
)

:: ── Install / update dependencies ────────────────────────────
echo [INFO] Installing dependencies...
.venv\Scripts\pip install -q --upgrade pip
.venv\Scripts\pip install -q -r requirements.txt
echo [OK] Dependencies ready

:: ── Launch ───────────────────────────────────────────────────
echo.
echo [INFO] Starting dashboard at http://localhost:8501
echo        Press Ctrl+C to stop
echo.
.venv\Scripts\streamlit run app.py --server.headless true --server.port 8501 --server.fileWatcherType none
pause
