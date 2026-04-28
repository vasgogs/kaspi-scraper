#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$ROOT_DIR"

PHARMACIES_CSV="${1:-}"
if [[ -n "$PHARMACIES_CSV" && "$PHARMACIES_CSV" != --* ]]; then
  shift
else
  PHARMACIES_CSV=""
fi

CMD=(./venv/bin/python wolt_project/scripts/run_portfolio_brand_monitor.py --only-brand "Vitrum")
if [[ -n "$PHARMACIES_CSV" ]]; then
  CMD+=(--pharmacies-csv "$PHARMACIES_CSV")
fi
CMD+=("$@")
"${CMD[@]}"
