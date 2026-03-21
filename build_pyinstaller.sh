#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"

PYTHON_BIN="${PYTHON_BIN:-./venv/bin/python}"
BUILD_SUPPORT_DIR="${BUILD_SUPPORT_DIR:-$ROOT_DIR/build_support}"
IMGKAP_SRC_DIR="${IMGKAP_SRC_DIR:-$HOME/imgkap}"
IMGKAP_OUT="$BUILD_SUPPORT_DIR/imgkap"

mkdir -p "$BUILD_SUPPORT_DIR"

if [[ -f "$IMGKAP_SRC_DIR/imgkap.c" ]]; then
  echo "Building bundled imgkap from: $IMGKAP_SRC_DIR"
  gcc "$IMGKAP_SRC_DIR/imgkap.c" -O3 -s -lm -lfreeimage -o "$IMGKAP_OUT"
fi

"$PYTHON_BIN" -m PyInstaller --noconfirm --clean pymapcal.spec

echo
echo "Build finished:"
echo "  $ROOT_DIR/dist/pymapcal"
