#!/bin/bash
# ─────────────────────────────────────────────────────────────
# Intuit Hiring Dashboard — Mac/Linux setup & launcher
# Run once to install, then run again any time to start the app
# ─────────────────────────────────────────────────────────────

set -e
VENV=".venv"

echo ""
echo "  Intuit Hiring Dashboard"
echo "────────────────────────────"

# ── Check Python ─────────────────────────────────────────────
if ! command -v python3 &>/dev/null; then
  echo "❌  Python 3 not found."
  echo "    Install from https://www.python.org/downloads/ then re-run this script."
  exit 1
fi
PY=$(python3 --version)
echo "✅  $PY found"

# ── Create virtual environment if needed ─────────────────────
if [ ! -d "$VENV" ]; then
  echo "📦  Creating virtual environment..."
  python3 -m venv "$VENV"
fi

# ── Install / update dependencies ────────────────────────────
echo "📥  Installing dependencies..."
"$VENV/bin/pip" install -q --upgrade pip
"$VENV/bin/pip" install -q -r requirements.txt
echo "✅  Dependencies ready"

# ── Launch ───────────────────────────────────────────────────
echo ""
echo "🚀  Starting dashboard at http://localhost:8501"
echo "    Press Ctrl+C to stop"
echo ""
"$VENV/bin/streamlit" run app.py \
  --server.headless true \
  --server.port 8501 \
  --server.fileWatcherType none
