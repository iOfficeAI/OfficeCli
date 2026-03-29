#!/usr/bin/env bash
set -euo pipefail

# Fallback builder for this imported template.
# Usage: ./build.sh [output.pptx]
OUT="${1:-$PWD/budget-review-classic.pptx}"
SRC_DIR="$(cd "$(dirname "$0")" && pwd)"
cp "$SRC_DIR/template.pptx" "$OUT"
echo "Generated: $OUT"
