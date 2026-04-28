#!/bin/bash
# Автоматический запуск Kaspi скрейпера

# 1) Путь к проекту
cd /home/vas/kaspi-scraper || exit 1

# 1.5) Загружаем секреты, если есть .env (не коммитить его)
if [ -f .env ]; then
  set -a
  . ./.env
  set +a
fi

# Reviews scraping is pinned to Almaty unless explicitly overridden.
export REVIEWS_CITY="${REVIEWS_CITY:-Алматы}"
export COMPETITOR_CITY="${COMPETITOR_CITY:-Алматы}"
export COMPETITOR_FILE_PREFIX="${COMPETITOR_FILE_PREFIX:-kaspi_competitors_almaty}"

# 2) Активируем виртуальное окружение
source venv/bin/activate

# 3) Запускаем скрипт, всё пишем в лог
python Scraper_Kaspi.py >> logs/kaspi_scraper.log 2>&1
python Scraper_Kaspi.py competitors "${COMPETITOR_PRODUCTS_CSV:-competitor_products.csv}" "${COMPETITOR_CITY}" "${COMPETITOR_FILE_PREFIX}" >> logs/kaspi_competitors.log 2>&1 || true
python kaspi_price_leaders_live.py --region "${PRICE_LEADERS_REGION:-Алматы}" --send-telegram >> logs/kaspi_price_leaders.log 2>&1 || true
