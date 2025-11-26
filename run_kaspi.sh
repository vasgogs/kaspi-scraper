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

# 2) Активируем виртуальное окружение
source venv/bin/activate

# 3) Запускаем скрипт, всё пишем в лог
python Scraper_Kaspi.py >> logs/kaspi_scraper.log 2>&1
