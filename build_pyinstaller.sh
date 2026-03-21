#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"

PYTHON_BIN="${PYTHON_BIN:-./venv/bin/python}"

"$PYTHON_BIN" -m PyInstaller --noconfirm --clean pymapcal.spec

echo
echo "Build finished:"
echo "  $ROOT_DIR/dist/pymapcal"
