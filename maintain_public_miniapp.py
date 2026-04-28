#!/usr/bin/env python3
import os
import re
import signal
import subprocess
import sys
import time
from pathlib import Path

import requests


ROOT_DIR = Path(__file__).resolve().parent
WEBAPP_DIR = ROOT_DIR / "telegram_webapp"
ENV_PATH = ROOT_DIR / ".env"
LOG_DIR = ROOT_DIR / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

_ENV_CACHE: dict[str, str] | None = None
PORT = int(os.environ.get("MINIAPP_PORT", "8002"))
HOST = os.environ.get("MINIAPP_HOST", "127.0.0.1")
HEALTH_URL = f"http://{HOST}:{PORT}/health"
UVICORN_BIN = WEBAPP_DIR / "venv" / "bin" / "uvicorn"

URL_RE = re.compile(r"your url is:\s*(https?://\S+)", re.IGNORECASE)
TOKEN_RE = re.compile(r"^TELEGRAM_BOT_TOKEN=(.+)$", re.MULTILINE)


def _env_map() -> dict[str, str]:
    global _ENV_CACHE
    if _ENV_CACHE is not None:
        return _ENV_CACHE
    values: dict[str, str] = {}
    if ENV_PATH.exists():
        for raw_line in ENV_PATH.read_text(encoding="utf-8", errors="ignore").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            values[key.strip()] = value.strip().strip("'\"")
    _ENV_CACHE = values
    return values


def load_env_value(*keys: str, default: str = "") -> str:
    env_values = _env_map()
    for key in keys:
        value = os.environ.get(key)
        if value is None:
            value = env_values.get(key)
        if value is not None and str(value).strip():
            return str(value).strip()
    return default


MENU_TEXT = load_env_value("MINIAPP_MENU_TEXT", default="ПООООНК")
PUBLIC_URL = load_env_value("MINIAPP_PUBLIC_URL")
TARGET_CHAT_ID = load_env_value("MINIAPP_CHAT_ID", "MISSION_APRIL_CHAT_ID", "MISSION_FEBRUARY_CHAT_ID")
TUNNEL_SUBDOMAIN = load_env_value("MINIAPP_TUNNEL_SUBDOMAIN")


def load_bot_token() -> str:
    token = load_env_value("TELEGRAM_BOT_TOKEN")
    if not token:
        raise RuntimeError("TELEGRAM_BOT_TOKEN not found in .env")
    return token


def set_menu_button(token: str, url: str, chat_id: str | None = None) -> None:
    payload = {
        "menu_button": {
            "type": "web_app",
            "text": MENU_TEXT,
            "web_app": {"url": url},
        }
    }
    if chat_id:
        payload["chat_id"] = str(chat_id)
    resp = requests.post(
        f"https://api.telegram.org/bot{token}/setChatMenuButton",
        json=payload,
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    if not data.get("ok"):
        raise RuntimeError(f"setChatMenuButton failed: {data}")
    scope = f"chat {chat_id}" if chat_id else "default"
    print(f"[miniapp] menu button updated ({scope}) -> {url}", flush=True)


def set_menu_buttons(token: str, url: str) -> None:
    set_menu_button(token, url)
    if TARGET_CHAT_ID:
        set_menu_button(token, url, TARGET_CHAT_ID)


def webapp_is_healthy() -> bool:
    try:
        resp = requests.get(HEALTH_URL, timeout=3)
        return resp.ok and "ok" in resp.text.lower()
    except Exception:
        return False


def start_webapp() -> subprocess.Popen | None:
    if webapp_is_healthy():
        print(f"[miniapp] webapp already healthy on {HEALTH_URL}", flush=True)
        return None
    if not UVICORN_BIN.exists():
        raise RuntimeError(f"uvicorn not found: {UVICORN_BIN}")
    cmd = [
        str(UVICORN_BIN),
        "main:app",
        "--host",
        HOST,
        "--port",
        str(PORT),
        "--reload",
    ]
    out = open(LOG_DIR / "miniapp_uvicorn.log", "a", encoding="utf-8")
    proc = subprocess.Popen(
        cmd,
        cwd=str(WEBAPP_DIR),
        stdout=out,
        stderr=subprocess.STDOUT,
        text=True,
    )
    print(f"[miniapp] started webapp pid={proc.pid} on {HOST}:{PORT}", flush=True)

    deadline = time.time() + 30
    while time.time() < deadline:
        if webapp_is_healthy():
            return proc
        if proc.poll() is not None:
            raise RuntimeError("webapp process exited early; check logs/miniapp_uvicorn.log")
        time.sleep(1)
    raise RuntimeError("webapp did not become healthy in time")


def start_localtunnel_and_get_url() -> tuple[subprocess.Popen, str]:
    cmd = ["npx", "--yes", "localtunnel", "--port", str(PORT)]
    if TUNNEL_SUBDOMAIN:
        cmd.extend(["--subdomain", TUNNEL_SUBDOMAIN])
    proc = subprocess.Popen(
        cmd,
        cwd=str(ROOT_DIR),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
    )
    deadline = time.time() + 60
    lines: list[str] = []
    assert proc.stdout is not None
    while time.time() < deadline:
        line = proc.stdout.readline()
        if not line:
            if proc.poll() is not None:
                break
            time.sleep(0.2)
            continue
        line = line.strip()
        lines.append(line)
        print(f"[tunnel] {line}", flush=True)
        m = URL_RE.search(line)
        if m:
            return proc, m.group(1).rstrip("/")
    raise RuntimeError(f"failed to get localtunnel url; output={lines[-10:]}")


def terminate(proc: subprocess.Popen | None, name: str) -> None:
    if not proc:
        return
    if proc.poll() is not None:
        return
    try:
        proc.terminate()
        proc.wait(timeout=8)
    except Exception:
        try:
            proc.kill()
        except Exception:
            pass
    print(f"[miniapp] stopped {name}", flush=True)


def normalize_public_url(value: str) -> str:
    value = (value or "").strip()
    if not value:
        return ""
    if not value.startswith(("http://", "https://")):
        value = "https://" + value
    return value.rstrip("/") + "/"


def main() -> int:
    token = load_bot_token()
    webapp_proc: subprocess.Popen | None = None
    tunnel_proc: subprocess.Popen | None = None
    stopping = False
    fixed_public_url = normalize_public_url(PUBLIC_URL)

    def _handle_signal(sig, _frame):
        nonlocal stopping
        stopping = True
        print(f"[miniapp] signal={sig}, stopping", flush=True)
        terminate(tunnel_proc, "tunnel")
        terminate(webapp_proc, "webapp")
        sys.exit(0)

    signal.signal(signal.SIGINT, _handle_signal)
    signal.signal(signal.SIGTERM, _handle_signal)

    webapp_proc = start_webapp()

    if fixed_public_url:
        set_menu_buttons(token, fixed_public_url)
        print(f"[miniapp] fixed public url mode -> {fixed_public_url}", flush=True)
        while not stopping:
            if not webapp_is_healthy():
                raise RuntimeError("local webapp healthcheck failed")
            time.sleep(5)
        return 0

    while not stopping:
        try:
            tunnel_proc, public_url = start_localtunnel_and_get_url()
            set_menu_buttons(token, public_url + "/")
            while tunnel_proc.poll() is None and not stopping:
                if not webapp_is_healthy():
                    raise RuntimeError("local webapp healthcheck failed")
                time.sleep(5)
            if not stopping:
                raise RuntimeError("localtunnel exited")
        except Exception as exc:
            print(f"[miniapp] {exc}", flush=True)
            terminate(tunnel_proc, "tunnel")
            tunnel_proc = None
            time.sleep(3)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
