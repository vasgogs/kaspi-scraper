#!/usr/bin/env python3
"""Постоянный запуск списка mission."""

from __future__ import annotations

import os
import sys
import time
from datetime import datetime
from pathlib import Path
import re
import subprocess

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
RUN_KASPI_SCRAPE = None


def _mission_scraper_defaults() -> dict[str, str]:
    return {
        "SCRAPER_JOB_TIMEOUT": _env_first("MISSION_SCRAPER_JOB_TIMEOUT_SEC", default="180") or "180",
        "SCRAPER_BROWSER_RETRIES": _env_first("MISSION_SCRAPER_BROWSER_RETRIES", default="1") or "1",
        "SCRAPER_SELLER_DISCOVERY_ATTEMPTS": _env_first("MISSION_SCRAPER_SELLER_ATTEMPTS", default="6") or "6",
        "SCRAPER_SELLER_WAIT_TIMEOUT_MS": _env_first("MISSION_SCRAPER_SELLER_WAIT_MS", default="9000") or "9000",
        "SCRAPER_SELLER_DISCOVERY_RELOAD_EVERY": _env_first("MISSION_SCRAPER_RELOAD_EVERY", default="2") or "2",
    }


def _prepare_mission_scraper_env():
    # Mission should fail fast on problematic SKUs and never block the whole cycle.
    for key, value in _mission_scraper_defaults().items():
        os.environ[key] = value


def _get_run_kaspi_scrape():
    global RUN_KASPI_SCRAPE
    if RUN_KASPI_SCRAPE is None:
        from Scraper_Kaspi import run_kaspi_scrape as _run_kaspi_scrape

        RUN_KASPI_SCRAPE = _run_kaspi_scrape
    return RUN_KASPI_SCRAPE


def load_env_file():
    env_path = BASE_DIR / ".env"
    if not env_path.exists():
        return
    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip("'\"")
        if key and key not in os.environ:
            os.environ[key] = value


def discover_cities(csv_path: Path) -> list[str] | None:
    try:
        df = pd.read_csv(csv_path)
    except Exception:
        return None
    if "region" not in df.columns:
        return None
    cities: set[str] = set()
    for raw in df["region"].dropna().astype(str):
        for part in re.split(r"[;,/|]", raw):
            part = part.strip()
            if part:
                cities.add(part)
    return sorted(cities) if cities else None


def _env_first(*names: str, default: str | None = None) -> str | None:
    for name in names:
        value = os.environ.get(name)
        if value is not None and str(value).strip():
            return value
    return default


def _latest_mission_file(prefix: str) -> Path | None:
    results_dir = BASE_DIR / "RESULTS"
    if not results_dir.exists():
        return None
    candidates = sorted(
        results_dir.glob(f"{prefix}_*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return candidates[0] if candidates else None


def run_once(csv_path: Path, chat_id: str | None):
    if not csv_path.exists():
        print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] ⚠️ CSV не найден: {csv_path}")
        return
    cities = discover_cities(csv_path)
    output_prefix = _env_first(
        "MISSION_APRIL_PREFIX",
        "MISSION_FEBRUARY_PREFIX",
        default="mission_april",
    )
    run_kaspi_scrape = _get_run_kaspi_scrape()
    started_at = time.time()
    run_kaspi_scrape(
        cities=cities,
        extra_csv_paths=[csv_path],
        include_base=False,
        chat_id=chat_id,
        alert_only=True,
        change_alerts=False,
        output_prefix=output_prefix,
        priority="secondary",
    )
    latest = _latest_mission_file(output_prefix or "mission_april")
    if not latest or latest.stat().st_mtime < started_at:
        print(
            f"[{datetime.now():%Y-%m-%d %H:%M:%S}] ⚠️ Mission файл не появился после прогона."
        )
        retry_flag = os.environ.get("MISSION_EXPORT_RETRY", "1").strip().lower()
        if retry_flag in {"1", "true", "yes"}:
            delay = int(os.environ.get("MISSION_EXPORT_RETRY_DELAY_SEC", "20"))
            print(
                f"[{datetime.now():%Y-%m-%d %H:%M:%S}] 🔁 Перезапуск миссии через {delay}s..."
            )
            time.sleep(max(5, delay))
            run_kaspi_scrape(
                cities=cities,
                extra_csv_paths=[csv_path],
                include_base=False,
                chat_id=chat_id,
                alert_only=True,
                change_alerts=False,
                output_prefix=output_prefix,
                priority="secondary",
            )
            latest = _latest_mission_file(output_prefix or "mission_april")
    if os.environ.get("MISSION_SUPPRESS_TELEGRAM", "0").strip().lower() in {"1", "true", "yes"}:
        if latest and latest.exists():
            _run_campaign_analytics(output_prefix or "mission_april")
        return
    if latest and latest.exists():
        try:
            env = os.environ.copy()
            env.setdefault("MISSION_FILE_PREFIX", output_prefix or "mission_april")
            broadcast_cmd = [sys.executable, str(BASE_DIR / "mission_broadcast.py")]
            subprocess.Popen(broadcast_cmd, env=env)
        except Exception as exc:
            print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] ⚠️ Mission broadcast failed to start: {exc}")
        _run_campaign_analytics(output_prefix or "mission_april")
    else:
        print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] ⚠️ Mission файл не найден для отправки.")


def _run_campaign_analytics(prefix: str):
    enabled = os.environ.get("MISSION_CAMPAIGN_ANALYTICS_ENABLED", "0").strip().lower()
    if enabled not in {"1", "true", "yes"}:
        return
    campaign_start = _env_first("MISSION_CAMPAIGN_START")
    campaign_end = _env_first("MISSION_CAMPAIGN_END")
    if not campaign_start or not campaign_end:
        print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] ⚠️ Mission campaign analytics skipped: start/end not set.")
        return
    try:
        cmd = [
            sys.executable,
            str(BASE_DIR / "analyze_mission_campaign_quality.py"),
            "--prefix",
            prefix,
            "--campaign-start",
            campaign_start,
            "--campaign-end",
            campaign_end,
        ]
        subprocess.run(cmd, cwd=BASE_DIR, check=True, timeout=180)
    except Exception as exc:
        print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] ⚠️ Mission campaign analytics failed: {exc}")


def main():
    load_env_file()
    _prepare_mission_scraper_env()
    print(
        "⚙️ Mission scrape limits: "
        f"job_timeout={os.environ.get('SCRAPER_JOB_TIMEOUT')}s, "
        f"browser_retries={os.environ.get('SCRAPER_BROWSER_RETRIES')}, "
        f"seller_attempts={os.environ.get('SCRAPER_SELLER_DISCOVERY_ATTEMPTS')}, "
        f"seller_wait={os.environ.get('SCRAPER_SELLER_WAIT_TIMEOUT_MS')}ms"
    )
    csv_path = _env_first(
        "MISSION_APRIL_CSV",
        "MISSION_FEBRUARY_CSV",
        default="миссия апрель.csv",
    )
    csv_path = Path(csv_path)
    if not csv_path.is_absolute():
        csv_path = BASE_DIR / csv_path
    chat_id = _env_first("MISSION_APRIL_CHAT_ID", "MISSION_FEBRUARY_CHAT_ID")
    interval = int(_env_first("MISSION_APRIL_INTERVAL_SEC", "MISSION_FEBRUARY_INTERVAL_SEC", default="1800"))
    print(
        f"🚀 Mission runner запущен: файл {csv_path}, интервал {interval} сек, чат {chat_id or 'не задан'}"
    )
    while True:
        started = datetime.now()
        try:
            run_once(csv_path, chat_id)
        except KeyboardInterrupt:
            print("⏹ Остановлено пользователем.")
            return
        except Exception as exc:
            print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] ❌ Ошибка запуска: {exc}")
        elapsed = (datetime.now() - started).total_seconds()
        sleep_for = max(5, interval - int(elapsed))
        time.sleep(sleep_for)


if __name__ == "__main__":
    main()
