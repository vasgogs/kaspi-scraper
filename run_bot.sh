#!/bin/bash
# Запуск Telegram-бота для управления Kaspi скрейпером

cd /home/vas/kaspi-scraper || exit 1

# Загружаем секреты, если есть .env
if [ -f .env ]; then
  set -a
  . ./.env
  set +a
fi

source venv/bin/activate

python kaspi_bot.py >> logs/kaspi_bot.log 2>&1
