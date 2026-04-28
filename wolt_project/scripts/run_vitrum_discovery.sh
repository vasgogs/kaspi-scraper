#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$ROOT_DIR"

./venv/bin/python wolt_discover_pharmacies.py \
  --city-slug almaty \
  --country-alpha2 KZ \
  --language en \
  --output-prefix wolt_almaty_pharmacies \
  "$@"
