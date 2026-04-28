#!/bin/bash
# Ежедневный запуск Wolt-портфеля (обновляет срезы для миниаппа)

set -euo pipefail

cd /home/vas/kaspi-scraper || exit 1

LOCK_FILE="state/wolt_daily.lock"
mkdir -p state
exec 9>"$LOCK_FILE"
if ! flock -n 9; then
  echo "$(date -u '+%Y-%m-%d %H:%M:%S') [wolt_daily] another run is active, skipping"
  exit 0
fi

if [ -f .env ]; then
  set -a
  . ./.env
  set +a
fi

source venv/bin/activate

if [ "${WOLT_DAILY_REFRESH_PHARMACIES:-1}" = "1" ]; then
  python wolt_discover_pharmacies.py --city-slug almaty || echo "$(date -u '+%Y-%m-%d %H:%M:%S') [wolt_daily] pharmacy discovery failed, keeping previous catalog"
fi

CMD=(python wolt_project/scripts/run_portfolio_brand_monitor.py --continue-on-error)

if [ "${WOLT_DAILY_SEND_TELEGRAM:-0}" = "1" ]; then
  CMD+=(--send-telegram)
fi

if [ -n "${WOLT_DAILY_ONLY_BRAND:-}" ]; then
  CMD+=(--only-brand "${WOLT_DAILY_ONLY_BRAND}")
fi

if [ -n "${WOLT_DAILY_STAGE:-}" ]; then
  CMD+=(--stage "${WOLT_DAILY_STAGE}")
fi

if [ "${WOLT_DAILY_DRY_RUN:-0}" = "1" ]; then
  CMD+=(--dry-run)
fi

"${CMD[@]}"

python wolt_project/scripts/build_wolt_assortment_gap.py --city-slug almaty || echo "$(date -u '+%Y-%m-%d %H:%M:%S') [wolt_daily] assortment gap build failed"
