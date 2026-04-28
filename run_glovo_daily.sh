#!/bin/bash
# Daily Glovo run: refresh pharmacies and collect brand price/stock snapshots.

set -euo pipefail

cd /home/vas/kaspi-scraper || exit 1

LOCK_FILE="state/glovo_daily.lock"
mkdir -p state logs
exec 9>"$LOCK_FILE"
if ! flock -n 9; then
  echo "$(date -u '+%Y-%m-%d %H:%M:%S') [glovo_daily] another run is active, skipping"
  exit 0
fi

if [ -f .env ]; then
  set -a
  . ./.env
  set +a
fi

source venv/bin/activate

DISCOVERY_CMD=(python glovo_address_grid_discover.py)

if [ -n "${GLOVO_GRID_ROWS:-}" ]; then
  DISCOVERY_CMD+=(--grid-rows "${GLOVO_GRID_ROWS}")
fi

if [ -n "${GLOVO_GRID_COLS:-}" ]; then
  DISCOVERY_CMD+=(--grid-cols "${GLOVO_GRID_COLS}")
fi

if [ -n "${GLOVO_GRID_PADDING:-}" ]; then
  DISCOVERY_CMD+=(--grid-padding "${GLOVO_GRID_PADDING}")
fi

if [ -n "${GLOVO_GRID_SLEEP_MS:-}" ]; then
  DISCOVERY_CMD+=(--sleep-ms "${GLOVO_GRID_SLEEP_MS}")
fi

"${DISCOVERY_CMD[@]}" >> logs/glovo_daily.log 2>&1

CMD=(python glovo_project/scripts/run_portfolio_brand_monitor.py --continue-on-error)

if [ "${GLOVO_DAILY_SEND_TELEGRAM:-0}" = "1" ]; then
  CMD+=(--send-telegram)
fi

if [ -n "${GLOVO_DAILY_ONLY_BRAND:-}" ]; then
  CMD+=(--only-brand "${GLOVO_DAILY_ONLY_BRAND}")
fi

if [ -n "${GLOVO_DAILY_STAGE:-}" ]; then
  CMD+=(--stage "${GLOVO_DAILY_STAGE}")
fi

"${CMD[@]}" >> logs/glovo_daily.log 2>&1
