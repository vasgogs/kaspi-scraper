#!/usr/bin/env python3
import json
import os
import time
import subprocess
import fcntl
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

BASE_DIR = Path(__file__).resolve().parent
STATE_DIR = BASE_DIR / "state"
PROGRESS_PATH = STATE_DIR / "progress.json"
LOCK_PATH = Path(os.environ.get("SCRAPER_LOCK_FILE", "/tmp/kaspi_scraper.lock"))
LOG_PATH = Path(os.environ.get("SCRAPER_WATCHDOG_LOG", str(BASE_DIR / "logs" / "scraper_watchdog.log")))
STALE_SEC = int(os.environ.get("SCRAPER_WATCHDOG_STALE_SEC", "900"))
CHECK_SEC = int(os.environ.get("SCRAPER_WATCHDOG_CHECK_SEC", "60"))
COOLDOWN_SEC = int(os.environ.get("SCRAPER_WATCHDOG_COOLDOWN_SEC", "600"))
SERVICE_NAME = os.environ.get("SCRAPER_WATCHDOG_SERVICE", "kaspi-bot.service")
ALMATY_TZ = ZoneInfo("Asia/Almaty")


def _log(msg: str) -> None:
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    ts = datetime.now(ALMATY_TZ).strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_PATH, "a", encoding="utf-8") as fh:
        fh.write(f"[{ts}] {msg}\n")


def _scrape_active() -> bool:
    try:
        LOCK_PATH.parent.mkdir(parents=True, exist_ok=True)
        fd = os.open(LOCK_PATH, os.O_CREAT | os.O_RDWR)
    except Exception:
        return False
    try:
        try:
            fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
            fcntl.flock(fd, fcntl.LOCK_UN)
            return False
        except BlockingIOError:
            return True
    finally:
        os.close(fd)


def _progress_incomplete(data: dict) -> bool:
    cities = data.get("cities") or {}
    for city_state in cities.values():
        if not city_state.get("completed", False):
            return True
    return False


def main() -> None:
    last_restart = 0.0
    _log("watchdog started")
    while True:
        try:
            if PROGRESS_PATH.exists():
                data = json.loads(PROGRESS_PATH.read_text(encoding="utf-8"))
                if _progress_incomplete(data) and _scrape_active():
                    mtime = PROGRESS_PATH.stat().st_mtime
                    age = time.time() - mtime
                    if age >= STALE_SEC and (time.time() - last_restart) >= COOLDOWN_SEC:
                        _log(f"stale progress for {int(age)}s, restarting {SERVICE_NAME}")
                        subprocess.run(["systemctl", "restart", SERVICE_NAME], check=False)
                        last_restart = time.time()
            time.sleep(CHECK_SEC)
        except Exception as exc:
            _log(f"watchdog error: {exc}")
            time.sleep(CHECK_SEC)


if __name__ == "__main__":
    main()
