#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$ROOT_DIR"

./venv/bin/python wolt_stock_monitor.py \
  --config wolt_project/config/wolt_positions.csv \
  --output-prefix wolt_almaty_vitrum \
  "$@"
