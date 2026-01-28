#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
ENTRY_POINT="$ROOT_DIR/arc_poc.py"

WINDOW_ICON="$ROOT_DIR/assets/app.ico"
EXE_ICON="$ROOT_DIR/assets/icons/app.ico"

ICON_PATH=""
if [[ -f "$EXE_ICON" ]]; then
  ICON_PATH="$EXE_ICON"
elif [[ -f "$WINDOW_ICON" ]]; then
  ICON_PATH="$WINDOW_ICON"
fi

DATA_ARGS=()
if [[ -f "$WINDOW_ICON" ]]; then
  DATA_ARGS+=(--add-data "${WINDOW_ICON};assets")
fi
if [[ -f "$EXE_ICON" && "$EXE_ICON" != "$WINDOW_ICON" ]]; then
  DATA_ARGS+=(--add-data "${EXE_ICON};assets/icons")
fi

NAME="ARC App Scanner"
PYI_CMD=(pyinstaller --noconfirm --clean --onefile --windowed --name "$NAME" "${DATA_ARGS[@]}" "$ENTRY_POINT")
if [[ -n "$ICON_PATH" ]]; then
  PYI_CMD+=(--icon "$ICON_PATH")
fi

"${PYI_CMD[@]}"
echo "Build complete: $ROOT_DIR/dist/${NAME}.exe"
