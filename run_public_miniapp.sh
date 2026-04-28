#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "${ROOT_DIR}"

mkdir -p logs

# Prefer stable domain URL; fallback keeps manager away from temporary tunnels.
: "${MINIAPP_PUBLIC_URL:=https://app.newnordicshop.kz/}"
export MINIAPP_PUBLIC_URL

# Stop previous manager instance (if any)
pkill -f "python3 .*maintain_public_miniapp.py" || true
sleep 0.4

nohup python3 maintain_public_miniapp.py > logs/miniapp_manager.log 2>&1 &
echo "miniapp manager started: pid=$! (url=${MINIAPP_PUBLIC_URL})"
