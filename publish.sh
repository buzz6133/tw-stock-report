#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")" && pwd)"

python3 "$ROOT/stock_report.py" report

cp "$ROOT/reports/latest.html" "$ROOT/docs/index.html"
cp "$ROOT/reports/latest.md" "$ROOT/docs/latest.md"

git add "$ROOT/docs/index.html" "$ROOT/docs/latest.md"

echo "Updated docs/index.html and docs/latest.md"
