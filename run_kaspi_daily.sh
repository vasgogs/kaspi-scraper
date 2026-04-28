#!/bin/bash
# Ежедневный тихий запуск Kaspi скрейпера (только файл + email)

cd /home/vas/kaspi-scraper || exit 1

if [ -f .env ]; then
  set -a
  . ./.env
  set +a
fi

export TELEGRAM_FILE_ONLY=1
export REVIEWS_CITY="${REVIEWS_CITY:-Алматы}"
export COMPETITOR_CITY="${COMPETITOR_CITY:-Алматы}"
export COMPETITOR_FILE_PREFIX="${COMPETITOR_FILE_PREFIX:-kaspi_competitors_almaty}"

source venv/bin/activate

python Scraper_Kaspi.py >> logs/kaspi_scraper_daily.log 2>&1
python Scraper_Kaspi.py competitors "${COMPETITOR_PRODUCTS_CSV:-competitor_products.csv}" "${COMPETITOR_CITY}" "${COMPETITOR_FILE_PREFIX}" >> logs/kaspi_competitors_daily.log 2>&1 || true
python kaspi_price_leaders_live.py --region "${PRICE_LEADERS_REGION:-Алматы}" --send-telegram >> logs/kaspi_price_leaders_daily.log 2>&1 || true
