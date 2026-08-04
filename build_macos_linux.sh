#!/usr/bin/env bash
set -euo pipefail

VENV_DIR=".venv_build"
VENV_PYTHON="$VENV_DIR/bin/python"

recreate_venv() {
  echo "Creating virtual environment at $VENV_DIR"
  rm -rf "$VENV_DIR" || true
  if [[ -e "$VENV_DIR" ]]; then
    echo "Could not remove $VENV_DIR. Please delete it manually or fix permissions."
    exit 1
  fi
  python3 -m venv "$VENV_DIR"
}

venv_is_usable() {
  [[ -x "$VENV_PYTHON" ]] || return 1
  "$VENV_PYTHON" -c "import pathlib, sys; expected = pathlib.Path('$PWD/$VENV_DIR').resolve(); actual = pathlib.Path(sys.prefix).resolve(); raise SystemExit(0 if expected == actual else 1)" >/dev/null 2>&1
}

install_build_dependencies() {
  "$VENV_PYTHON" -m pip install --upgrade pip
  "$VENV_PYTHON" -m pip install -r requirements-dev.txt
  "$VENV_PYTHON" -m pip install -e .
}

# Recreate venv if missing/broken (for moved/renamed project folders).
if ! venv_is_usable; then
  recreate_venv
fi

if ! install_build_dependencies; then
  echo "Dependency install failed. Recreating virtual environment and retrying once."
  recreate_venv
  install_build_dependencies
fi

# Keep PyInstaller cache/config inside the project to avoid macOS permission issues.
export PYINSTALLER_CONFIG_DIR="$PWD/.pyinstaller"

"$VENV_PYTHON" -m PyInstaller --noconfirm --clean --windowed --name db_data_validator --paths src run_app.py

echo "Build complete. Find the application under dist/db_data_validator/"
