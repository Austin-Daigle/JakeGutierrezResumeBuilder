#!/bin/bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

PYTHON_BIN="${PYTHON_BIN:-python3}"
VENV_DIR=".venv_macos_build"

"$PYTHON_BIN" -m venv "$VENV_DIR"
source "$VENV_DIR/bin/activate"

python -m pip install --upgrade pip setuptools wheel
python -m pip install --upgrade pyinstaller pyinstaller-hooks-contrib
python -m pip install --upgrade reportlab python-docx pyspellchecker

pyinstaller --noconfirm --clean ResumeTemplateMaker_macOS.spec

echo ""
echo "Build complete. App is in: dist/JakeGResumeBuilder v.1.0.2.app"
