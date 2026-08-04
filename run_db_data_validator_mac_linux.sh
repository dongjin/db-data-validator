#!/usr/bin/env bash
set -euo pipefail

VENV_DIR=".venv"
VENV_PYTHON="$VENV_DIR/bin/python"

recreate_venv() {
  echo "Creating virtual environment at $VENV_DIR"
  rm -rf "$VENV_DIR"
  python3 -m venv "$VENV_DIR"
}

# Recreate the venv if missing or broken (common after folder moves/renames).
if [[ ! -x "$VENV_PYTHON" ]] || ! "$VENV_PYTHON" -c "import sys" >/dev/null 2>&1; then
  recreate_venv
fi

"$VENV_PYTHON" -m pip install --upgrade pip
"$VENV_PYTHON" -m pip install -r requirements.txt
"$VENV_PYTHON" -m pip install -e .
"$VENV_PYTHON" run_app.py
