#!/usr/bin/env bash
# Render the demo deck to PNG slide screenshots for the README.
#
# Pipeline: examples/demo.ts -> .pptx -> (LibreOffice headless) .pdf -> (poppler) .png
#
# Requirements (not needed to use the library — only to regenerate screenshots):
#   - LibreOffice  (provides `soffice`):  brew install --cask libreoffice
#   - poppler      (provides `pdftoppm`): brew install poppler
#
# Usage: ./scripts/screenshots.sh
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUT_IMG="$ROOT/docs/images"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

SOFFICE="${SOFFICE:-/Applications/LibreOffice.app/Contents/MacOS/soffice}"
command -v "$SOFFICE" >/dev/null 2>&1 || [ -x "$SOFFICE" ] || SOFFICE="$(command -v soffice || true)"
[ -n "$SOFFICE" ] && { [ -x "$SOFFICE" ] || command -v "$SOFFICE" >/dev/null; } || {
  echo "error: soffice (LibreOffice) not found. Install: brew install --cask libreoffice" >&2; exit 1; }
command -v pdftoppm >/dev/null || { echo "error: pdftoppm not found. Install: brew install poppler" >&2; exit 1; }

echo "==> Generating demo deck"
( cd "$ROOT" && npm run --silent demo >/dev/null )
PPTX="$(ls -t "$ROOT"/output/*.pptx | head -1)"
echo "    deck: $PPTX"

echo "==> Converting PPTX -> PDF (LibreOffice headless)"
"$SOFFICE" --headless --convert-to pdf --outdir "$TMP" "$PPTX" >/dev/null
PDF="$(ls "$TMP"/*.pdf | head -1)"

echo "==> Rendering PDF pages -> PNG (poppler)"
mkdir -p "$OUT_IMG"
pdftoppm -png -r 110 "$PDF" "$OUT_IMG/slide"

echo "==> Done. Wrote:"
ls -1 "$OUT_IMG"/slide-*.png
