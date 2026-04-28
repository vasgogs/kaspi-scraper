#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
WEBAPP_DIR="${ROOT_DIR}/telegram_webapp"

cd "${WEBAPP_DIR}"
./venv/bin/uvicorn main:app --host 127.0.0.1 --port 8001 --reload
