#!/usr/bin/env bash
# Count MCP tools registered in the server.
# Usage: bash scripts/count_tools.sh

set -euo pipefail

SRC="$(cd "$(dirname "$0")/../src" && pwd)"

echo "=== Tool count by module ==="
for f in "$SRC/server.py" "$SRC/ppt_com/"*.py; do
  count=$(grep -c 'name="ppt_' "$f" 2>/dev/null || true)
  if [ "$count" -gt 0 ]; then
    printf "%4d  %s\n" "$count" "$(basename "$f")"
  fi
done

echo ""
echo "=== Total ==="
total=$(grep -rh 'name="ppt_' "$SRC/" | grep -oP 'ppt_[^"]+' | sort -u | wc -l)
printf "%4d  tools\n" "$total"

echo ""
echo "=== README states ==="
readme="$(cd "$(dirname "$0")/.." && pwd)/README.md"
stated=$(grep -oP '(?<=\*\*)\d+(?= tools\b)' "$readme" | head -1 || echo "?")
printf "%4s  tools (README.md)\n" "$stated"

if [ "$total" -ne "$stated" ] 2>/dev/null; then
  echo ""
  echo "WARNING: Mismatch! Update README.md (and README_ja.md) to reflect the actual count."
fi
