#!/usr/bin/env bash
set -euo pipefail

MANIFEST_PATH="${1:-manifests/manifest.dev.xml}"
SKIP_CERT_TRUST="${SKIP_CERT_TRUST:-0}"
DRY_RUN="${DRY_RUN:-0}"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

if [[ "$MANIFEST_PATH" != /* ]]; then
  MANIFEST_PATH="$REPO_ROOT/$MANIFEST_PATH"
fi

if [[ ! -f "$MANIFEST_PATH" ]]; then
  echo "Manifest not found: $MANIFEST_PATH" >&2
  exit 1
fi

WORD_WEF_DIR="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
POWERPOINT_WEF_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
EXCEL_WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"

CERT_PATH="$HOME/.office-addin-dev-certs/localhost.pem"

echo "Registering Office add-in (Word, PowerPoint, Excel only)..."
echo "Manifest: $MANIFEST_PATH"

if [[ "$SKIP_CERT_TRUST" != "1" && -f "$CERT_PATH" ]]; then
  if [[ "$DRY_RUN" == "1" ]]; then
    echo "[DryRun] Would trust cert: $CERT_PATH"
  else
    echo "Trusting development certificate (may prompt for sudo password)..."
    sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain "$CERT_PATH" || true
  fi
else
  echo "Skipping certificate trust step."
fi

for dir in "$WORD_WEF_DIR" "$POWERPOINT_WEF_DIR" "$EXCEL_WEF_DIR"; do
  if [[ "$DRY_RUN" == "1" ]]; then
    echo "[DryRun] Would create: $dir"
    echo "[DryRun] Would copy manifest to: $dir"
  else
    mkdir -p "$dir"
    cp "$MANIFEST_PATH" "$dir/"
  fi
done

echo "Done. Open Word, PowerPoint, or Excel and add via Insert > Add-ins > My Add-ins."
