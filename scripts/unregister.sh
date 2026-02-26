#!/usr/bin/env bash
set -euo pipefail

MANIFEST_NAME="manifest.dev.xml"
if [[ $# -gt 0 ]]; then
  MANIFEST_NAME="$(basename "$1")"
fi

DRY_RUN="${DRY_RUN:-0}"

WORD_WEF_DIR="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
POWERPOINT_WEF_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
EXCEL_WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"

echo "Unregistering Office add-in (Word, PowerPoint, Excel only)..."
echo "Manifest filename: $MANIFEST_NAME"

for dir in "$WORD_WEF_DIR" "$POWERPOINT_WEF_DIR" "$EXCEL_WEF_DIR"; do
  target="$dir/$MANIFEST_NAME"
  if [[ -f "$target" ]]; then
    if [[ "$DRY_RUN" == "1" ]]; then
      echo "[DryRun] Would remove: $target"
    else
      rm "$target"
      echo "Removed: $target"
    fi
  else
    echo "Not found: $target"
  fi
done

echo "Done."
