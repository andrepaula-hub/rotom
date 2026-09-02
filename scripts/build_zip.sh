#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
COMMIT="$(git -C "$ROOT" rev-parse --short=12 HEAD)"
OUT_DIR="$ROOT/dist"
STAGE="$(mktemp -d)"

mkdir -p "$STAGE/ifood" "$OUT_DIR"
cp "$ROOT/ifood/print_bridge.py" "$STAGE/ifood/print_bridge.py"
cp "$ROOT/ifood/rotom_lite.py" "$STAGE/ifood/rotom_lite.py"
cat > "$STAGE/ifood/rotom_version.json" <<JSON
{
  "version": "$COMMIT",
  "commit": "$COMMIT"
}
JSON

(cd "$STAGE" && zip -qr "$OUT_DIR/rotom-$COMMIT.zip" ifood)
rm -rf "$STAGE"

echo "$OUT_DIR/rotom-$COMMIT.zip"

