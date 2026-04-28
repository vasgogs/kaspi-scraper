import os
import time
from datetime import datetime
import re
import json
import threading
import uuid
import sys
from pathlib import Path
from html import unescape
import subprocess
import tempfile
import shutil
import pandas as pd
import requests

from Scraper_Kaspi import (
    run_kaspi_scrape,
    run_keyword_search_report,
    bot_help_text,
    resolve_kaspi_link,
    analyze_product_card,
    format_product_analysis,
    gpt_product_opinion,
    _split_telegram_text,
    render_mission_slice_images,
)


BASE_DIR = Path(__file__).resolve().parent
USER_DATA_DIR = BASE_DIR / "user_data"
MISSION_DECEMBER_DEFAULT = (
    os.environ.get("MISSION_APRIL_CSV")
    or os.environ.get("MISSION_FEBRUARY_CSV")
    or "миссия апрель.csv"
)
MISSION_FILE_PREFIX = os.environ.get("MISSION_FILE_PREFIX") or "mission_april"
MAIN_FILE_PREFIX = os.environ.get("MAIN_FILE_PREFIX", "kaspi_prices")
MSP_SELLER_PATTERN = re.compile(r"^аптека\s*msp(?:\s+(?:алматы|астана|шымкент))?$", re.IGNORECASE)
MISSION_SUPPRESS_TELEGRAM = os.environ.get("MISSION_SUPPRESS_TELEGRAM", "0").strip().lower() in {"1", "true", "yes"}
MISSION_BOT_AUTO_ENABLED = os.environ.get("MISSION_BOT_AUTO_ENABLED", "0").strip().lower() in {"1", "true", "yes"}
BROADCAST_IDS = {
    chat.strip()
    for chat in re.split(r"[;,]", os.environ.get("BROADCAST_CHAT_IDS", ""))
    if chat.strip()
}
MISSION_ADMIN_CHATS = {
    chat.strip()
    for chat in re.split(r"[;,]", os.environ.get("MISSION_ADMIN_CHAT_IDS", "")) if chat.strip()
} or BROADCAST_IDS
MAIN_ADMIN_CHATS = {
    chat.strip()
    for chat in re.split(r"[;,]", os.environ.get("MAIN_ADMIN_CHAT_IDS", "")) if chat.strip()
} or BROADCAST_IDS
NATASHA_PRICE_ADMIN_IDS = {
    chat.strip()
    for chat in re.split(r"[;,]", os.environ.get("NATASHA_PRICE_ADMIN_IDS", "")) if chat.strip()
} or MAIN_ADMIN_CHATS or BROADCAST_IDS
NATASHA_PRICE_CSV = os.environ.get("NATASHA_PRICE_CSV") or str(BASE_DIR / "prices_natasha.csv")
NATASHA_PRICE_OUTPUT_PREFIX = os.environ.get("NATASHA_PRICE_OUTPUT_PREFIX", "prices_natasha_scraped")
NATASHA_PRICE_WORKERS = max(1, int(os.environ.get("NATASHA_PRICE_WORKERS", "14")))
NATASHA_PRICE_ITEKA_MAX_PAGES = max(1, int(os.environ.get("NATASHA_PRICE_ITEKA_MAX_PAGES", "50")))
MISSION_SUBSCRIBERS: set[str] = set()
MISSION_PUSH_STOP: threading.Event | None = None
MISSION_PUSH_THREAD: threading.Thread | None = None
MISSION_LAST_SENT: dict[str, float] = {}

SCRAPER_THREAD: threading.Thread | None = None
SCRAPER_STOP_EVENT: threading.Event | None = None
PROGRESS = {"status": "idle", "city": "", "done": 0, "note": ""}
PROGRESS_LOCK = threading.Lock()
PENDING_EXPECTED: dict[str, dict] = {}
PENDING_LOCK = threading.Lock()
EXTRA_PREVIEW_LIMIT = 15
CLEAR_CONFIRM_TOKEN = "confirm_clear"
PENDING_CITY: dict[str, dict] = {}
PENDING_CITY_LOCK = threading.Lock()
AUTO_THREADS: dict[str, threading.Thread] = {}
AUTO_STOPS: dict[str, threading.Event] = {}
AUTO_CONTEXT: dict[str, dict] = {}
MISSION_AUTO_THREADS: dict[str, threading.Thread] = {}
MISSION_AUTO_STOPS: dict[str, threading.Event] = {}
PENDING_ADD: dict[str, dict] = {}
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-3.5-turbo")
MEMORY_LOCK = threading.Lock()
TAVILY_API_KEY = os.environ.get("TAVILY_API_KEY")
ALLOWED_CHATS: set[str] | None = None
ANALYSIS_REQUESTS: dict[str, dict] = {}
ANALYSIS_IN_PROGRESS: set[str] = set()
PRICE_ANALYSIS_IN_PROGRESS: set[str] = set()
ANALYSIS_LOCK = threading.Lock()
LINK_ACTIONS: dict[str, dict] = {}
RESEARCH_CACHE: dict[str, dict] = {}
RESEARCH_CONTEXT: dict[str, dict] = {}
RESEARCH_USAGE: dict[str, list[float]] = {}
RESEARCH_IN_PROGRESS: set[str] = set()
RESEARCH_LOCK = threading.Lock()
RESEARCH_CACHE_TTL = int(os.environ.get("RESEARCH_CACHE_TTL", "3600"))
RESEARCH_LIMIT_MAX = int(os.environ.get("RESEARCH_LIMIT_MAX", "5"))
RESEARCH_LIMIT_WINDOW = int(os.environ.get("RESEARCH_LIMIT_WINDOW_SEC", "1800"))
RESEARCH_FAST_RESULTS = int(os.environ.get("RESEARCH_FAST_RESULTS", "5"))
RESEARCH_DEEP_RESULTS = int(os.environ.get("RESEARCH_DEEP_RESULTS", "10"))
RESEARCH_MODE_PRESETS = {
    "fast": {"label": "⚡ Быстрый поиск", "description": "до 5 источников, короткий вывод", "max_results": RESEARCH_FAST_RESULTS},
    "deep": {"label": "🔍 Глубокий поиск", "description": "до 10 источников, подробный анализ", "max_results": RESEARCH_DEEP_RESULTS},
}
PENDING_RESEARCH: dict[str, dict] = {}
_ENV_LOADED = False
BOT_LOG_PATH = Path(os.environ.get("BOT_LOG_PATH", str(BASE_DIR / "logs" / "kaspi_bot_events.log")))
AGENT_WORKDIR = Path(os.environ.get("AGENT_WORKDIR", str(BASE_DIR.parent))).resolve()
AGENT_MODEL = os.environ.get("OPENAI_AGENT_MODEL", "gpt-4.1-mini")
AGENT_MAX_STEPS = max(1, int(os.environ.get("AGENT_MAX_STEPS", "6")))
AGENT_CMD_TIMEOUT_SEC = max(10, int(os.environ.get("AGENT_CMD_TIMEOUT_SEC", "120")))
AGENT_OUTPUT_LIMIT = max(500, int(os.environ.get("AGENT_OUTPUT_LIMIT", "2500")))
AGENT_AUTO_MODE_DEFAULT = os.environ.get("AGENT_AUTO_MODE_DEFAULT", "0").strip().lower() in {"1", "true", "yes"}
AGENT_MODE_CHATS: set[str] = set()
AGENT_THREADS: dict[str, threading.Thread] = {}
AGENT_STATE_LOCK = threading.Lock()
AGENT_BLOCK_PATTERNS: list[tuple[str, str]] = [
    (r"(^|[;&|])\s*sudo(\s|$)", "sudo запрещён"),
    (r"(^|[;&|])\s*rm(\s|$)", "удаление через rm запрещено в agent-режиме"),
    (r"\bxargs\b[^\n]*\brm\b", "удаление через xargs rm запрещено в agent-режиме"),
    (r"\bfind\b[^\n]*\s-delete\b", "удаление через find -delete запрещено в agent-режиме"),
    (r"(^|[;&|])\s*rm\s+-rf(\s|$)", "опасное удаление запрещено"),
    (r"\bgit\s+reset\s+--hard\b", "git reset --hard запрещён"),
    (r"\bgit\s+checkout\s+--\b", "git checkout -- запрещён"),
    (r"\bshutdown\b|\breboot\b|\bpoweroff\b", "команды выключения запрещены"),
    (r":\(\)\s*\{\s*:\|\:&\s*\};:", "fork bomb запрещён"),
    (r"\bmkfs\b|\bfdisk\b|\bparted\b", "диск-утилиты запрещены"),
    (r"\bdd\s+if=", "dd if= запрещён"),
]


def _log_event(kind: str, payload: dict | None = None):
    try:
        BOT_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().isoformat(timespec="seconds")
        entry = {"ts": stamp, "kind": kind, "payload": payload or {}}
        with open(BOT_LOG_PATH, "a", encoding="utf-8") as fh:
            fh.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except Exception:
        pass


def load_env_from_file():
    """Load .env vars so the bot still works if script started without run_bot.sh."""
    global _ENV_LOADED
    if _ENV_LOADED:
        return
    env_path = BASE_DIR / ".env"
    if not env_path.exists():
        _ENV_LOADED = True
        return
    for raw_line in env_path.read_text().splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip("'\"")
        if key and key not in os.environ:
            os.environ[key] = value
    _ENV_LOADED = True


def user_extra_csv(chat_id: str) -> Path:
    USER_DATA_DIR.mkdir(parents=True, exist_ok=True)
    return USER_DATA_DIR / f"{chat_id}_extra.csv"


def mission_december_csv() -> Path:
    path = Path(MISSION_DECEMBER_DEFAULT)
    if not path.is_absolute():
        path = BASE_DIR / path
    return path


def start_mission_auto(token: str, chat_id: str, interval_sec: int = 3600):
    path = mission_december_csv()
    if not path.exists():
        send_bot_message(token, chat_id, f"Не нашёл файл для миссии: {path}", reply_markup=build_keyboard())
        return
    stop_event = MISSION_AUTO_STOPS.get(chat_id)
    if stop_event:
        stop_event.set()
    stop_event = threading.Event()
    MISSION_AUTO_STOPS[chat_id] = stop_event

    def _loop():
        send_bot_message(
            token,
            chat_id,
            f"🎯 Миссия {path.name}: автозапуск каждые {interval_sec // 60} мин",
            reply_markup=build_keyboard(),
        )
        while not stop_event.is_set():
            cycle_started_at = time.time()
            try:
                run_kaspi_scrape(
                    include_base=False,
                    chat_id=chat_id,
                    extra_csv_paths=[path],
                    cities=None,
                    alert_only=True,
                    change_alerts=False,
                    output_prefix=MISSION_FILE_PREFIX,
                    priority="secondary",
                )
                # После успешного прогона пробуем разослать отчёты (если настроены TELETHON_* и цели)
                if not MISSION_SUPPRESS_TELEGRAM:
                    try:
                        env = os.environ.copy()
                        env.setdefault("MISSION_FILE_PREFIX", MISSION_FILE_PREFIX)
                        broadcast_cmd = [sys.executable, str(BASE_DIR / "mission_broadcast.py")]
                        subprocess.Popen(broadcast_cmd, env=env)
                    except Exception as exc:
                        print(f"⚠️ Mission auto: broadcast failed to start: {exc}")
            except Exception as exc:
                send_bot_message(token, chat_id, f"❌ Авто-миссия: ошибка {exc}", reply_markup=build_keyboard())
            sleep_for = max(5, int(cycle_started_at + interval_sec - time.time()))
            for _ in range(sleep_for):
                if stop_event.is_set():
                    break
                time.sleep(1)
        send_bot_message(token, chat_id, "🎯 Авто-миссия остановлена.", reply_markup=build_keyboard())
        MISSION_AUTO_STOPS.pop(chat_id, None)
        MISSION_AUTO_THREADS.pop(chat_id, None)

    th = threading.Thread(target=_loop, daemon=True)
    MISSION_AUTO_THREADS[chat_id] = th
    th.start()


def competitor_csv_paths() -> tuple[list[Path], list[Path]]:
    """Возвращает существующие/отсутствующие CSV со списком конкурентов."""
    env_value = os.environ.get("COMPETITOR_PRODUCTS_CSV")
    candidates: list[Path] = []
    if env_value:
        for part in re.split(r"[;,]", env_value):
            part = part.strip()
            if part:
                candidates.append(Path(part))
    else:
        candidates.append(Path("competitor_products.csv"))
    normalized: list[Path] = []
    for path in candidates:
        normalized_path = path if path.is_absolute() else BASE_DIR / path
        normalized.append(normalized_path)
    existing = [p for p in normalized if p.exists()]
    missing = [p for p in normalized if not p.exists()]
    return existing, missing


def keyword_csv_path() -> Path:
    raw = os.environ.get("SEARCH_KEYWORDS_CSV", "search_keywords.csv")
    path = Path(raw)
    if not path.is_absolute():
        path = BASE_DIR / path
    return path


def _research_cache_key(query: str, max_results: int | None = None) -> str:
    base = re.sub(r"\s+", " ", (query or "").strip().lower())
    extra = max_results or RESEARCH_FAST_RESULTS
    return f"{base}|{extra}"


def research_cache_available(query: str, max_results: int | None = None) -> bool:
    key = _research_cache_key(query, max_results)
    if not key:
        return False
    now = time.time()
    with RESEARCH_LOCK:
        entry = RESEARCH_CACHE.get(key)
        if not entry:
            return False
        if now - entry["ts"] > RESEARCH_CACHE_TTL:
            RESEARCH_CACHE.pop(key, None)
            return False
        return True


def get_cached_research(query: str, max_results: int | None = None):
    key = _research_cache_key(query, max_results)
    if not key:
        return None
    now = time.time()
    with RESEARCH_LOCK:
        entry = RESEARCH_CACHE.get(key)
        if not entry:
            return None
        if now - entry["ts"] > RESEARCH_CACHE_TTL:
            RESEARCH_CACHE.pop(key, None)
            return None
        return entry["results"]


def set_research_cache(query: str, results: list[dict], max_results: int | None = None):
    key = _research_cache_key(query, max_results)
    if not key:
        return
    with RESEARCH_LOCK:
        RESEARCH_CACHE[key] = {"results": results, "ts": time.time()}


def research_quota_status(chat_id: str) -> tuple[bool, int | None]:
    now = time.time()
    with RESEARCH_LOCK:
        usage = [ts for ts in RESEARCH_USAGE.get(chat_id, []) if now - ts < RESEARCH_LIMIT_WINDOW]
        RESEARCH_USAGE[chat_id] = usage
        if len(usage) >= RESEARCH_LIMIT_MAX:
            wait = RESEARCH_LIMIT_WINDOW - int(now - min(usage))
            return False, max(1, wait)
        return True, None


def reserve_research_slot(chat_id: str) -> tuple[bool, int | None]:
    now = time.time()
    with RESEARCH_LOCK:
        usage = [ts for ts in RESEARCH_USAGE.get(chat_id, []) if now - ts < RESEARCH_LIMIT_WINDOW]
        if len(usage) >= RESEARCH_LIMIT_MAX:
            RESEARCH_USAGE[chat_id] = usage
            wait = RESEARCH_LIMIT_WINDOW - int(now - min(usage))
            return False, max(1, wait)
        usage.append(now)
        RESEARCH_USAGE[chat_id] = usage
        return True, None


def format_wait_time(seconds: int | None) -> str:
    if not seconds:
        return "несколько секунд"
    minutes, secs = divmod(max(0, seconds), 60)
    if minutes and secs:
        return f"{minutes} мин {secs} сек"
    if minutes:
        return f"{minutes} мин"
    return f"{secs} сек"


def tavily_search(query: str, max_results: int | None = None) -> list[dict]:
    if not TAVILY_API_KEY:
        raise RuntimeError("TAVILY_API_KEY не задан")
    payload = {
        "api_key": TAVILY_API_KEY,
        "query": query,
        "search_depth": "advanced",
        "max_results": max_results or RESEARCH_FAST_RESULTS,
        "include_answer": False,
    }
    resp = requests.post("https://api.tavily.com/search", json=payload, timeout=30)
    if resp.status_code != 200:
        raise RuntimeError(f"Tavily вернул {resp.status_code}: {resp.text[:200]}")
    data = resp.json()
    raw_results = data.get("results") or []
    normalized: list[dict] = []
    for item in raw_results:
        url = (item.get("url") or "").strip()
        if not url:
            continue
        scraped_text = scrape_page_text(url)
        base_content = (
            scraped_text
            or (item.get("content") or item.get("snippet") or item.get("raw_content") or "").strip()
        )
        normalized.append({
            "title": item.get("title") or url,
            "url": url,
            "content": base_content,
            "scraped": bool(scraped_text),
        })
        if max_results and len(normalized) >= max_results:
            break
    return normalized


def format_research_sources(results: list[dict]) -> str:
    lines = []
    for idx, entry in enumerate(results, start=1):
        title = (entry.get("title") or "Источник").strip()
        url = (entry.get("url") or "").strip()
        lines.append(f"[{idx}] {title} — {url}")
    return "\n".join(lines)


def research_mode_info(mode: str | None):
    key = (mode or "fast").lower()
    preset = RESEARCH_MODE_PRESETS.get(key)
    if not preset:
        key = "fast"
        preset = RESEARCH_MODE_PRESETS["fast"]
    return key, preset


def build_research_mode_keyboard():
    rows = []
    for key, info in RESEARCH_MODE_PRESETS.items():
        label = info.get("label", key.title())
        desc = info.get("description", "")
        text = f"{label} — {desc}" if desc else label
        rows.append([{"text": text, "callback_data": f"research_mode:{key}"}])
    rows.append([{"text": "Отмена", "callback_data": "research_cancel"}])
    return {"inline_keyboard": rows}


def scrape_page_text(url: str, max_chars: int = 6000) -> str:
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0 Safari/537.36"
    }
    try:
        resp = requests.get(url, headers=headers, timeout=20)
    except Exception:
        return ""
    if resp.status_code != 200 or not resp.text:
        return ""
    html = resp.text
    html = re.sub(r"(?is)<(script|style|noscript|template)[^>]*>.*?</\1>", " ", html)
    text = re.sub(r"(?is)<[^>]+>", " ", html)
    text = unescape(text)
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) > max_chars:
        text = text[:max_chars].rstrip() + "…"
    return text


def research_summary_via_gpt(question: str, results: list[dict], base_query: str | None = None) -> str:
    if not results:
        return "Источники пусты, нечего анализировать."
    context_lines = []
    for idx, entry in enumerate(results, start=1):
        snippet = re.sub(r"\s+", " ", (entry.get("content") or "").strip())
        if len(snippet) > 700:
            snippet = snippet[:700] + "…"
        context_lines.append(f"[{idx}] {entry.get('title') or 'Источник'}\nURL: {entry.get('url')}\n{snippet}")
    payload = (
        f"Исходный запрос: {base_query or question}\n"
        f"Текущий вопрос: {question}\n"
        "Ты действуешь как эксперт-аналитик, который уже прочитал все источники ниже. "
        "Сформируй единый связный отчёт (2–4 абзаца) с ключевыми находками, трендами и рекомендациями, по-русски. "
        "Не вставляй ссылки, не перечисляй источники. Просто дай осмысленный пересказ и выводы с полезными действиями.\n\n"
        + "\n\n".join(context_lines)
    )
    messages = [
        {
            "role": "system",
            "content": (
                "Ты аналитик по рынку e-commerce. "
                "Делай краткие выводы и обязательно ссылайся на источники формата [1], [2]."
            ),
        },
        {"role": "user", "content": payload},
    ]
    answer, error = _call_openai_chat(messages, temperature=0.2, max_tokens=400, timeout=25)
    return answer or (error or "Не удалось получить ответ от GPT.")


def research_help_text() -> str:
    window_text = format_wait_time(RESEARCH_LIMIT_WINDOW)
    return (
        "🔎 Кнопка «Ресёрч» предлагает два режима:\n"
        f" - {RESEARCH_MODE_PRESETS['fast']['label']}: {RESEARCH_MODE_PRESETS['fast']['description']}\n"
        f" - {RESEARCH_MODE_PRESETS['deep']['label']}: {RESEARCH_MODE_PRESETS['deep']['description']}\n"
        "После выбора просто напиши тему сообщением, и я запущу поиск.\n"
        "⚡ Команды: /research <тема> — быстрый режим, /research_deep <тема> — глубокий.\n"
        "✏️ /research_more <вопрос> — уточнить последний ресёрч из сохранённых источников.\n"
        f"⚖️ Лимит: до {RESEARCH_LIMIT_MAX} запусков каждые {window_text}. Кэш действует ~1 час."
    )


def _normalize_field(val: str) -> str:
    text = (val or "").strip()
    if not text:
        return ""
    if text.lower() in {"nan", "none", "null"}:
        return ""
    return text


def fetch_product_name(url: str) -> str:
    """Пытаемся получить название товара с карточки Kaspi (og:title)."""
    headers = {
        "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    }
    try:
        resp = requests.get(url, headers=headers, timeout=12)
        if resp.status_code != 200 or not resp.text:
            return ""
        html = resp.text
        match = re.search(r'<meta\\s+property=["\\\']og:title["\\\']\\s+content=["\\\']([^\\\']+)["\\\']', html, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        match_h1 = re.search(r"<h1[^>]*>(.*?)</h1>", html, re.IGNORECASE | re.DOTALL)
        if match_h1:
            name = re.sub(r"<.*?>", "", match_h1.group(1))
            return name.strip()
    except Exception:
        return ""
    return ""


def append_product_row(chat_id: str, link: str, name: str = "", expected: str = ""):
    """Добавляем товар во временный список, без дубликатов по ссылке, для конкретного пользователя."""
    csv_path = user_extra_csv(chat_id)
    link = resolve_kaspi_link(_normalize_field(link))
    name = _normalize_field(name)
    expected = _normalize_field(expected)
    if not name:
        name = fetch_product_name(link)
    df = pd.read_csv(csv_path) if csv_path.exists() else pd.DataFrame(
        columns=["product_name", "product_link", "expected_sellers"]
    )
    new_row = pd.DataFrame([{
        "product_name": name.strip(),
        "product_link": link.strip(),
        "expected_sellers": expected.strip(),
    }])
    df = pd.concat([df, new_row], ignore_index=True)
    df = df.drop_duplicates(subset=["product_link"]).reset_index(drop=True)
    df.to_csv(csv_path, index=False)


def update_expected_for_link(chat_id: str, link: str, expected: str):
    """Обновляем expected_sellers для конкретной ссылки и пользователя."""
    csv_path = user_extra_csv(chat_id)
    if not csv_path.exists():
        return
    df = pd.read_csv(csv_path)
    mask = df["product_link"] == link
    if mask.any():
        df.loc[mask, "expected_sellers"] = expected
        df.to_csv(csv_path, index=False)


def clear_extra_list(chat_id: str):
    """Очищаем временный список пользователя, оставляя заголовки."""
    csv_path = user_extra_csv(chat_id)
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(columns=["product_name", "product_link", "expected_sellers"])
    df.to_csv(csv_path, index=False)


def _latest_report_any(prefixes: list[str]) -> Path | None:
    results_dir = BASE_DIR / "RESULTS"
    if not results_dir.exists():
        return None
    seen: set[str] = set()
    candidates: list[Path] = []
    for prefix in prefixes:
        norm = prefix.strip()
        if not norm or norm in seen:
            continue
        seen.add(norm)
        candidates.extend(results_dir.glob(f"{norm}_*.xlsx"))
    if not candidates:
        return None
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0]


def latest_report_file(prefix: str) -> Path | None:
    return _latest_report_any([prefix])


def natasha_price_csv_path() -> Path:
    path = Path(NATASHA_PRICE_CSV)
    return path if path.is_absolute() else BASE_DIR / path


def latest_natasha_price_file() -> Path | None:
    results_dir = BASE_DIR / "RESULTS"
    candidates: list[Path] = []
    if results_dir.exists():
        candidates.extend(results_dir.glob(f"{NATASHA_PRICE_OUTPUT_PREFIX}_*.xlsx"))
    default_output = BASE_DIR / f"{NATASHA_PRICE_OUTPUT_PREFIX}.xlsx"
    if default_output.exists():
        candidates.append(default_output)
    if not candidates:
        return None
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0]


def _natasha_price_admin_allowed(chat_id: str, requester_id: str | None = None) -> bool:
    admins = {str(item).strip() for item in NATASHA_PRICE_ADMIN_IDS if str(item).strip()}
    if not admins:
        return False
    # For callbacks in groups, the user id is authoritative; chat id is only a fallback for plain commands.
    if requester_id:
        return str(requester_id) in admins
    return str(chat_id) in admins


def _natasha_price_time_label(path: Path) -> str:
    return datetime.fromtimestamp(path.stat().st_mtime).strftime("%d.%m %H:%M")


def _send_latest_natasha_price(token: str, chat_id: str, prefix: str = "Свежий скрейп prices_natasha") -> bool:
    latest = latest_natasha_price_file()
    if not latest or not latest.exists():
        return False
    caption = f"{prefix} — {_natasha_price_time_label(latest)}"
    send_bot_file(token, chat_id, latest, caption=caption)
    return True


def latest_mission_file() -> Path | None:
    prefixes = [
        MISSION_FILE_PREFIX,
        "mission",
        "mission_april",
        "mission_april_temp",
        "mission_march",
        "mission_march_temp",
        "mission_february",
        "mission_february_temp",
        "mission_january",
        "mission_january_temp",
    ]
    return _latest_report_any(prefixes)


def latest_mission_preview() -> Path | None:
    results_dir = BASE_DIR / "RESULTS"
    if not results_dir.exists():
        return None
    files = sorted(results_dir.glob("mission_preview_*.png"), key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0] if files else None


def _normalize_status_text(value) -> str:
    if value is None:
        return ""
    text = str(value).replace("\u00a0", " ").strip()
    return " ".join(text.split()).lower()


def _canonical_mission_seller(value) -> str:
    text = " ".join(str(value or "").replace("\u00a0", " ").split()).strip()
    if not text:
        return "Без аптеки"
    if MSP_SELLER_PATTERN.match(text):
        return "Аптека MSP"
    return text


def _format_kzt_short(value) -> str:
    try:
        amount = int(round(float(value)))
    except Exception:
        return "0 ₸"
    return f"{amount:,}".replace(",", " ") + " ₸"


def _mission_report_time_label(path: Path) -> str:
    match = re.search(r"(\d{4}-\d{2}-\d{2})_(\d{2})-(\d{2})-(\d{2})$", path.stem)
    if match:
        date_part, hour, minute, _second = match.groups()
        try:
            dt = datetime.strptime(f"{date_part} {hour}:{minute}", "%Y-%m-%d %H:%M")
            return dt.strftime("%d.%m %H:%M")
        except Exception:
            pass
    return datetime.fromtimestamp(path.stat().st_mtime).strftime("%d.%m %H:%M")


def _build_mission_breakdown_text(df: pd.DataFrame, group_col: str, title: str, group_order: list[str] | None = None) -> str:
    work = df.copy()
    if group_col not in work.columns or work.empty:
        return f"{title}\nНет данных."

    if group_col == "seller":
        work[group_col] = work[group_col].apply(_canonical_mission_seller)
    else:
        work[group_col] = work[group_col].fillna("Без региона").astype(str).str.replace("\u00a0", " ", regex=False)
        work[group_col] = work[group_col].map(lambda x: " ".join(x.split()) or "Без региона")

    work["status_norm"] = work.get("status", "").map(_normalize_status_text)
    work["is_ok"] = work["status_norm"].map(lambda text: text.startswith("ok") or text.startswith("ок"))
    work["is_attention"] = work["status_norm"].map(
        lambda text: text.startswith("внимание") or "не самое выгодное предложение" in text or "не лучшее" in text
    )
    work["is_price_problem"] = work["status_norm"].map(
        lambda text: text.startswith("проблема") or text.startswith("дороже") or text.startswith("дешевле")
    )
    work["is_seller_missing"] = work["status_norm"].eq("продавец отсутствует")
    work["is_no_offers"] = work["status_norm"].eq("нет продавцов на карточке")
    work["is_scrape_error"] = work["status_norm"].eq("ошибка скрейпа")
    work["is_no_price"] = work["status_norm"].eq("нет цены")

    work["actual_price_kzt"] = pd.to_numeric(work.get("actual_price_kzt"), errors="coerce")
    work["best_price_kzt"] = pd.to_numeric(work.get("best_price_kzt"), errors="coerce")
    work["market_gap_kzt"] = work["actual_price_kzt"] - work["best_price_kzt"]

    grouped = work.groupby(group_col, dropna=False)
    metrics: list[dict] = []
    for name, bucket in grouped:
        total = int(len(bucket))
        ok_count = int(bucket["is_ok"].sum())
        attention_count = int(bucket["is_attention"].sum())
        price_problem_count = int(bucket["is_price_problem"].sum())
        seller_missing_count = int(bucket["is_seller_missing"].sum())
        no_offers_count = int(bucket["is_no_offers"].sum())
        scrape_error_count = int(bucket["is_scrape_error"].sum())
        no_price_count = int(bucket["is_no_price"].sum())
        positive_gap = bucket.loc[bucket["market_gap_kzt"] > 0, "market_gap_kzt"]
        metrics.append(
            {
                "name": str(name),
                "total": total,
                "ok": ok_count,
                "attention": attention_count,
                "price_problem": price_problem_count,
                "seller_missing": seller_missing_count,
                "no_offers": no_offers_count,
                "scrape_error": scrape_error_count,
                "no_price": no_price_count,
                "avg_positive_gap": int(round(float(positive_gap.mean()))) if not positive_gap.empty else None,
            }
        )

    if group_order:
        rank = {value: idx for idx, value in enumerate(group_order)}
        metrics.sort(key=lambda item: (rank.get(item["name"], len(rank)), -item["total"], item["name"].lower()))
    else:
        metrics.sort(key=lambda item: (-item["total"], item["name"].lower()))

    lines = [title]
    for item in metrics:
        parts = [f"SKU {item['total']}", f"OK {item['ok']}"]
        if item["price_problem"]:
            parts.append(f"проблема цены {item['price_problem']}")
        if item["attention"]:
            parts.append(f"внимание {item['attention']}")
        if item["seller_missing"]:
            parts.append(f"нет продавца {item['seller_missing']}")
        if item["no_offers"]:
            parts.append(f"нет офферов {item['no_offers']}")
        if item["no_price"]:
            parts.append(f"нет цены {item['no_price']}")
        if item["scrape_error"]:
            parts.append(f"ошибка {item['scrape_error']}")
        if item["avg_positive_gap"] is not None:
            parts.append(f"ср. gap { _format_kzt_short(item['avg_positive_gap']) }")
        lines.append(f"• {item['name']}: " + " | ".join(parts))
    return "\n".join(lines)


def _build_latest_mission_summaries(path: Path) -> list[str]:
    try:
        df = pd.read_excel(path)
    except Exception as exc:
        return [f"Не смог собрать сводку миссии: {exc}"]
    required = {"seller", "region", "status"}
    if df.empty or not required.issubset(df.columns):
        return ["Сводка миссии недоступна: в файле нет нужных колонок."]

    ts = _mission_report_time_label(path)
    header = f"Сводка миссии ({ts})"
    seller_summary = _build_mission_breakdown_text(df, "seller", f"{header}\nПо аптекам")
    city_summary = _build_mission_breakdown_text(
        df,
        "region",
        f"{header}\nПо городам",
        group_order=["Алматы", "Астана", "Шымкент"],
    )
    return [seller_summary, city_summary]


def _send_latest_mission(token: str, chat_id: str) -> bool:
    """Возвращает True, если удалось что-то отправить."""
    if MISSION_SUPPRESS_TELEGRAM:
        return True
    latest = latest_mission_file()
    if latest and latest.exists():
        ts = _mission_report_time_label(latest)
        caption = f"Свежий отчёт миссии ({ts})"
        try:
            send_bot_file(token, chat_id, latest, caption=caption)
        except Exception as exc:
            send_bot_message(
                token,
                chat_id,
                f"Не смог отправить файл миссии: {exc}. Попробуй позже или попроси админа обновить.",
                reply_markup=build_keyboard(),
            )
            return False
        try:
            MISSION_LAST_SENT[chat_id] = latest.stat().st_mtime
        except Exception:
            pass
        try:
            df = pd.read_excel(latest)
        except Exception:
            df = None
        if df is not None and not df.empty:
            try:
                with tempfile.TemporaryDirectory() as tmpdir:
                    previews = render_mission_slice_images(df, Path(tmpdir), timestamp="latest")
                    for preview in previews:
                        preview_path = preview.get("path")
                        preview_caption = preview.get("caption") or "Сводка миссии"
                        if not preview_path or not Path(preview_path).exists():
                            continue
                        try:
                            send_bot_photo(token, chat_id, Path(preview_path), caption=preview_caption)
                        except Exception:
                            pass
            except Exception:
                pass
        else:
            preview = latest_mission_preview()
            if preview and preview.exists():
                try:
                    send_bot_photo(token, chat_id, preview, caption="Сводка миссии")
                except Exception:
                    pass
        return True
    return False


def ensure_mission_push_loop(token: str, interval_sec: int = 1800):
    global MISSION_PUSH_THREAD, MISSION_PUSH_STOP
    if MISSION_PUSH_THREAD and MISSION_PUSH_THREAD.is_alive():
        return
    stop_event = threading.Event()
    MISSION_PUSH_STOP = stop_event

    def _loop():
        while not stop_event.is_set():
            try:
                latest = latest_mission_file()
                mtime = latest.stat().st_mtime if latest and latest.exists() else 0
                for chat in list(MISSION_SUBSCRIBERS):
                    last_sent = MISSION_LAST_SENT.get(chat, 0)
                    if mtime and mtime > last_sent:
                        _send_latest_mission(token, chat)
            except Exception:
                pass
            for _ in range(interval_sec):
                if stop_event.is_set():
                    break
                time.sleep(1)

    th = threading.Thread(target=_loop, daemon=True)
    MISSION_PUSH_THREAD = th
    th.start()


def _telegram_post(url: str, data: dict, files=None, timeout: int = 30, retries: int = 2):
    for attempt in range(retries + 1):
        resp = requests.post(url, data=data, files=files, timeout=timeout)
        if resp.status_code == 200:
            return resp
        if resp.status_code == 429:
            try:
                payload = resp.json()
                retry_after = int(payload.get("parameters", {}).get("retry_after", 2))
            except Exception:
                retry_after = 2
            time.sleep(max(1, retry_after))
            continue
        if resp.status_code >= 500 and attempt < retries:
            time.sleep(1 + attempt)
            continue
        return resp
    return resp


def send_bot_message(token: str, chat_id: str, text: str, reply_markup=None):
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    data = {"chat_id": chat_id, "text": text}
    if reply_markup:
        data["reply_markup"] = json.dumps(reply_markup)
    resp = _telegram_post(url, data=data)
    if resp.status_code != 200:
        _log_event("telegram_send_error", {"chat_id": chat_id, "status": resp.status_code, "body": resp.text[:500]})
        raise RuntimeError(f"Telegram send failed: {resp.status_code} {resp.text}")
    _log_event("telegram_send_ok", {"chat_id": chat_id, "text": text[:500]})


def send_bot_file(token: str, chat_id: str, file_path: Path, caption: str | None = None):
    url = f"https://api.telegram.org/bot{token}/sendDocument"
    with open(file_path, "rb") as fh:
        files = {"document": fh}
        data = {"chat_id": chat_id}
        if caption:
            data["caption"] = caption
        resp = _telegram_post(url, data=data, files=files, timeout=60)
    if resp.status_code != 200:
        _log_event(
            "telegram_file_error",
            {"chat_id": chat_id, "file": str(file_path), "status": resp.status_code, "body": resp.text[:500]},
        )
        raise RuntimeError(f"Telegram file send failed: {resp.status_code} {resp.text}")
    _log_event("telegram_file_ok", {"chat_id": chat_id, "file": str(file_path), "caption": caption or ""})


def send_chunked_message(token: str, chat_id: str, text: str):
    chunks = _split_telegram_text(text) or [text]
    for chunk in chunks:
        if not chunk:
            continue
        send_bot_message(token, chat_id, chunk, reply_markup=build_keyboard())


def send_bot_photo(token: str, chat_id: str, file_path: Path, caption: str | None = None):
    url = f"https://api.telegram.org/bot{token}/sendPhoto"
    with open(file_path, "rb") as fh:
        files = {"photo": fh}
        data = {"chat_id": chat_id}
        if caption:
            data["caption"] = caption
        resp = _telegram_post(url, data=data, files=files, timeout=60)
    if resp.status_code != 200:
        _log_event(
            "telegram_photo_error",
            {"chat_id": chat_id, "file": str(file_path), "status": resp.status_code, "body": resp.text[:500]},
        )
        raise RuntimeError(f"Telegram photo send failed: {resp.status_code} {resp.text}")
    _log_event("telegram_photo_ok", {"chat_id": chat_id, "file": str(file_path), "caption": caption or ""})


def handle_callback(token: str, chat_id: str, data_cb: str, requester_id: str | None = None):
    try:
        if data_cb == "run_full":
            start_scrape(token, chat_id, include_base=True)
        elif data_cb == "run_extra":
            tok = uuid.uuid4().hex
            with PENDING_CITY_LOCK:
                PENDING_CITY[tok] = {"chat_id": chat_id, "mode": "run_extra"}
            send_bot_message(token, chat_id, "Выбери город:", reply_markup=build_city_keyboard(tok))
        elif data_cb in {"run_natasha_prices", "run_competitors"}:
            start_natasha_price_scrape(token, chat_id, requester_id=requester_id)
        elif data_cb == "run_keywords":
            start_keyword_search(token, chat_id)
        elif data_cb == "run_mission":
            start_mission_december_scrape(token, chat_id)
        elif data_cb == "run_alerts":
            tok = uuid.uuid4().hex
            with PENDING_CITY_LOCK:
                PENDING_CITY[tok] = {"chat_id": chat_id, "mode": "run_alerts"}
            send_bot_message(token, chat_id, "Выбери город:", reply_markup=build_city_keyboard(tok))
        elif data_cb == "auto_on":
            tok = uuid.uuid4().hex
            with PENDING_CITY_LOCK:
                PENDING_CITY[tok] = {"chat_id": chat_id, "mode": "auto_alerts", "interval_sec": 300}
            send_bot_message(token, chat_id, "Выбери город для авто-алертов (каждые 5 мин):", reply_markup=build_city_keyboard(tok))
        elif data_cb == "auto_off":
            stop_auto_alerts(token, chat_id)
        elif data_cb == "auto_status":
            send_bot_message(token, chat_id, auto_alert_status_text(chat_id), reply_markup=build_keyboard())
        elif data_cb == "analysis_start":
            request_analysis_link(token, chat_id)
        elif data_cb == "research_info":
            PENDING_RESEARCH.pop(chat_id, None)
            instructions = (
                "Выбери тип ресёрча:\n"
                f"• {RESEARCH_MODE_PRESETS['fast']['label']} — {RESEARCH_MODE_PRESETS['fast']['description']}\n"
                f"• {RESEARCH_MODE_PRESETS['deep']['label']} — {RESEARCH_MODE_PRESETS['deep']['description']}\n"
                "После выбора просто ответь сообщением с темой, и я запущу поиск."
            )
            send_bot_message(token, chat_id, instructions, reply_markup=build_research_mode_keyboard())
        elif data_cb.startswith("research_mode:"):
            _, _, mode_choice = data_cb.partition(":")
            mode_key, preset = research_mode_info(mode_choice)
            PENDING_RESEARCH[chat_id] = {"mode": mode_key, "ts": time.time()}
            send_bot_message(
                token,
                chat_id,
                f"{preset.get('label', 'Ресёрч')}: напиши тему одним сообщением, и я сразу начну поиск.",
                reply_markup=build_keyboard(),
            )
        elif data_cb == "research_cancel":
            if chat_id in PENDING_RESEARCH:
                PENDING_RESEARCH.pop(chat_id, None)
                send_bot_message(token, chat_id, "Режим ожидания темы отменён.", reply_markup=build_keyboard())
            else:
                send_bot_message(token, chat_id, "Нечего отменять, выбери режим ещё раз.", reply_markup=build_keyboard())
        elif data_cb.startswith("linkact:"):
            parts = data_cb.split(":")
            if len(parts) == 3:
                _prefix, action, tok = parts
                ctx = LINK_ACTIONS.pop(tok, None)
                if not ctx or str(ctx.get("chat_id")) != str(chat_id):
                    send_bot_message(token, chat_id, "Эта ссылка больше не доступна.", reply_markup=build_keyboard())
                    return
                link = ctx.get("link")
                if action == "add":
                    append_product_row(chat_id, link, "", "")
                    send_bot_message(token, chat_id, "✅ Добавил в твой временный список.", reply_markup=build_keyboard())
                elif action == "card":
                    start_analysis_worker(token, chat_id, link)
                elif action == "price":
                    start_price_analysis_worker(token, chat_id, link)
                elif action == "research":
                    query = ctx.get("title") or ctx.get("link")
                    if not query:
                        send_bot_message(token, chat_id, "Нечего исследовать, попробуй снова отправить ссылку.", reply_markup=build_keyboard())
                    else:
                        request_research(token, chat_id, query, mode="fast")
                elif action == "cancel":
                    send_bot_message(token, chat_id, "Отменил действие.", reply_markup=build_keyboard())
                else:
                    send_bot_message(token, chat_id, "Неизвестная кнопка.", reply_markup=build_keyboard())
            else:
                send_bot_message(token, chat_id, "Неизвестная кнопка.", reply_markup=build_keyboard())
        elif data_cb == "status":
            send_bot_message(token, chat_id, status_text(), reply_markup=build_keyboard())
        elif data_cb == "stop":
            request_stop(token, chat_id)
        elif data_cb == "list_extra":
            csv_path = user_extra_csv(chat_id)
            if csv_path.exists():
                try:
                    df = pd.read_csv(csv_path)
                except Exception as exc:
                    send_bot_message(token, chat_id, f"Не удалось прочитать список: {exc}", reply_markup=build_keyboard())
                else:
                    if df.empty:
                        send_bot_message(token, chat_id, "Временный список пуст.", reply_markup=build_keyboard())
                    elif len(df) > EXTRA_PREVIEW_LIMIT:
                        send_bot_message(token, chat_id, f"Записей: {len(df)}. Отправляю файл...", reply_markup=build_keyboard())
                        try:
                            send_bot_file(token, chat_id, csv_path, caption=f"Временный список ({len(df)} строк)")
                        except Exception as exc:
                            send_bot_message(token, chat_id, f"Не удалось отправить файл: {exc}", reply_markup=build_keyboard())
                    else:
                        send_bot_message(token, chat_id, extra_list_text(chat_id), reply_markup=build_keyboard())
            else:
                send_bot_message(token, chat_id, "Временный список пуст.", reply_markup=build_keyboard())
        elif data_cb == "last_file":
            fpath = latest_result_file()
            if fpath:
                try:
                    send_bot_file(token, chat_id, fpath, caption=f"Последний файл: {fpath.name}")
                except Exception as exc:
                    send_bot_message(token, chat_id, f"Не удалось отправить файл: {exc}", reply_markup=build_keyboard())
            else:
                send_bot_message(token, chat_id, "Нет сохранённых файлов в RESULTS.", reply_markup=build_keyboard())
        elif data_cb == "clear_extra":
            confirm_kb = {
                "inline_keyboard": [
                    [
                        {"text": "❗ Да, очистить", "callback_data": "clear_extra_confirm"},
                        {"text": "Отмена", "callback_data": "clear_extra_cancel"},
                    ]
                ]
            }
            send_bot_message(token, chat_id, "Точно очистить временный список?", reply_markup=confirm_kb)
        elif data_cb == "clear_extra_confirm":
            clear_extra_list(chat_id)
            send_bot_message(token, chat_id, "Временный список очищен.", reply_markup=build_keyboard())
        elif data_cb == "clear_extra_cancel":
            send_bot_message(token, chat_id, "Отменил очистку.", reply_markup=build_keyboard())
        elif data_cb == "confirm_add_links":
            pending = PENDING_ADD.pop(chat_id, None)
            if not pending or not pending.get("links"):
                send_bot_message(token, chat_id, "Нет ссылок для добавления.", reply_markup=build_keyboard())
            else:
                added = 0
                for link in pending["links"]:
                    append_product_row(chat_id, link, "", "")
                    added += 1
                send_bot_message(
                    token,
                    chat_id,
                    f"✅ Добавил {added} ссылок в твой временный список. Продавцы не заданы, при необходимости добавь через /add.",
                    reply_markup=build_keyboard(),
                )
        elif data_cb == "cancel_add_links":
            PENDING_ADD.pop(chat_id, None)
            send_bot_message(token, chat_id, "Отменил добавление ссылок.", reply_markup=build_keyboard())
        elif data_cb == "ask_gpt_info":
            send_bot_message(
                token,
                chat_id,
                "Используй /ask <вопрос> для обычного ответа GPT.\n"
                "Для выполнения задач с прогрессом: /agent <задача>.\n"
                "Включить авто-режим: /agent_on.",
                reply_markup=build_keyboard(),
            )
        elif data_cb.startswith("setexp:"):
            parts = data_cb.split(":")
            if len(parts) == 3:
                _prefix, tok, choice = parts
                with PENDING_LOCK:
                    info = PENDING_EXPECTED.get(tok)
                    if info is None:
                        send_bot_message(token, chat_id, "Не нашёл, попробуй ещё раз /add", reply_markup=build_keyboard())
                        return
                    if str(info.get("chat_id")) != str(chat_id):
                        send_bot_message(token, chat_id, "Эта выборка продавцов принадлежит другому чату.", reply_markup=build_keyboard())
                        return
                    selected: set[str] = info.get("selected", set())
                    if choice == "save":
                        selected_names: list[str] = []
                        for key in selected:
                            _, names = EXPECTED_OPTIONS[key]
                            selected_names.extend(names)
                        seen = set()
                        unique_names = []
                        for n in selected_names:
                            if n not in seen:
                                seen.add(n)
                                unique_names.append(n)
                        expected = "; ".join(unique_names)
                        update_expected_for_link(chat_id, info["link"], expected)
                        PENDING_EXPECTED.pop(tok, None)
                        msg = "Оставил без продавцов." if not expected else f"Записал: {expected}"
                        send_bot_message(token, chat_id, f"{msg}\nМожешь запускать скрейп.", reply_markup=build_keyboard())
                        return
                    if choice == "skip":
                        PENDING_EXPECTED.pop(tok, None)
                        send_bot_message(token, chat_id, "Оставил без продавцов. Можно запускать скрейп.", reply_markup=build_keyboard())
                        return
                    if choice == "all":
                        selected = set(EXPECTED_OPTIONS.keys())
                    else:
                        if choice in EXPECTED_OPTIONS:
                            if choice in selected:
                                selected.remove(choice)
                            else:
                                selected.add(choice)
                        else:
                            send_bot_message(token, chat_id, "Неизвестная кнопка", reply_markup=build_keyboard())
                            return
                    info["selected"] = selected
                    PENDING_EXPECTED[tok] = info
                    chosen = []
                    for key in selected:
                        label, _ = EXPECTED_OPTIONS[key]
                        chosen.append(label)
                    chosen_text = ", ".join(chosen) if chosen else "ничего не выбрано"
                    send_bot_message(
                        token,
                        chat_id,
                        f"Выбрано: {chosen_text}\nНажми Сохранить, чтобы записать.",
                        reply_markup=build_expected_keyboard(tok, selected),
                    )
            else:
                send_bot_message(token, chat_id, "Неизвестная кнопка", reply_markup=build_keyboard())
        elif data_cb.startswith("citysel:"):
            parts = data_cb.split(":")
            if len(parts) == 3:
                _p, tok, city_choice = parts
                with PENDING_CITY_LOCK:
                    info = PENDING_CITY.pop(tok, None)
                if not info:
                    send_bot_message(token, chat_id, "Не нашёл запрос города, попробуй снова.", reply_markup=build_keyboard())
                    return
                if city_choice == "cancel":
                    send_bot_message(token, chat_id, "Отменено.", reply_markup=build_keyboard())
                    return
                if str(info.get("chat_id")) != str(chat_id):
                    send_bot_message(token, chat_id, "Этот выбор города не для вашего чата.", reply_markup=build_keyboard())
                    return
                mode = info.get("mode")
                if mode == "run_extra":
                    start_scrape(token, chat_id, include_base=False, cities=[city_choice])
                elif mode == "run_alerts":
                    start_scrape(token, chat_id, include_base=False, alert_only=True, cities=[city_choice])
                elif mode == "auto_alerts":
                    interval_sec = info.get("interval_sec", 300)
                    start_auto_alerts(token, chat_id, interval_sec, city_choice)
                else:
                    send_bot_message(token, chat_id, "Неизвестный режим запуска.", reply_markup=build_keyboard())
            else:
                send_bot_message(token, chat_id, "Некорректный выбор города.", reply_markup=build_keyboard())
        else:
            send_bot_message(token, chat_id, "Неизвестная кнопка", reply_markup=build_keyboard())
    except Exception as exc:
        _log_event("callback_error", {"chat_id": chat_id, "data": data_cb, "error": str(exc)[:500]})
        try:
            send_bot_message(token, chat_id, f"❌ Ошибка обработки кнопки: {exc}", reply_markup=build_keyboard())
        except Exception:
            pass


def build_link_actions_keyboard(token_id: str):
    return {
        "inline_keyboard": [
            [
                {"text": "➕ Добавить во временный", "callback_data": f"linkact:add:{token_id}"},
            ],
            [
                {"text": "🔬 Анализ товара", "callback_data": f"linkact:card:{token_id}"},
                {"text": "💸 Анализ цен", "callback_data": f"linkact:price:{token_id}"},
            ],
            [
                {"text": "🔎 Ресёрч", "callback_data": f"linkact:research:{token_id}"},
            ],
            [
                {"text": "📝 Временный список", "callback_data": "list_extra"},
            ],
            [{"text": "Отмена", "callback_data": f"linkact:cancel:{token_id}"}],
        ]
    }


def parse_add_payload(text: str):
    """Парсим /add <link> | <name> | <sellers>"""
    payload = text
    payload = re.sub(r"^/add(@\S+)?", "", payload, flags=re.IGNORECASE).strip()
    parts = [p.strip() for p in payload.split("|")]
    link = parts[0] if parts else ""
    name = parts[1] if len(parts) > 1 else ""
    expected = parts[2] if len(parts) > 2 else ""
    return link, name, expected


def progress_callback(event: str, city: str | None = None, done: int | None = None, total: int | None = None):
    with PROGRESS_LOCK:
        if event == "start":
            PROGRESS.update({"status": "running", "city": "", "done": 0, "note": "Старт"})
        elif event == "city_start":
            PROGRESS.update({"status": "running", "city": city or "", "done": done or 0, "note": f"Город {city}, {total} товаров"})
        elif event == "progress":
            PROGRESS.update({"status": "running", "city": city or PROGRESS.get("city", ""), "done": done or 0, "note": "Идёт сбор"})
        elif event == "city_done":
            PROGRESS.update({"status": "running", "city": city or "", "done": done or PROGRESS.get('done', 0), "note": f"Город {city} завершён"})
        elif event == "no_data":
            PROGRESS.update({"status": "idle", "city": "", "done": 0, "note": "Нет данных"})
        elif event == "stopped":
            PROGRESS.update({"status": "stopped", "city": "", "done": done or 0, "note": "Остановлено"})
        elif event == "finished":
            PROGRESS.update({"status": "idle", "city": "", "done": 0, "note": "Готово"})


def build_keyboard():
    return {
        "inline_keyboard": [
            [
                {"text": "▶️ Запуск (основной)", "callback_data": "run_full"},
            ],
            [
                {"text": "🟢 Запуск только временных", "callback_data": "run_extra"},
            ],
            [
                {"text": "📄 Скрейп prices_natasha", "callback_data": "run_natasha_prices"},
            ],
            [
                {"text": "🔎 Поиск по ключам", "callback_data": "run_keywords"},
            ],
            [
                {"text": "🎯 Миссия", "callback_data": "run_mission"},
            ],
            [
                {"text": "🚨 Запуск (только алерты)", "callback_data": "run_alerts"},
            ],
            [
                {"text": "📂 Последний файл", "callback_data": "last_file"},
            ],
            [
                {"text": "🧹 Очистить временный", "callback_data": "clear_extra"},
            ],
            [
                {"text": "📊 Статус", "callback_data": "status"},
                {"text": "⏹ Остановить", "callback_data": "stop"},
            ],
            [
                {"text": "⏱️ Авто-алерты вкл (5 мин)", "callback_data": "auto_on"},
                {"text": "✋ Авто-алерты выкл", "callback_data": "auto_off"},
            ],
            [
                {"text": "ℹ️ Статус авто-алертов", "callback_data": "auto_status"},
            ],
            [
                {"text": "❓ Спросить GPT", "callback_data": "ask_gpt_info"},
            ],
        ]
    }


def build_add_links_keyboard():
    return {
        "inline_keyboard": [
            [{"text": "Добавить в временный", "callback_data": "confirm_add_links"}],
            [{"text": "Отмена", "callback_data": "cancel_add_links"}],
        ]
    }


EXPECTED_OPTIONS = {
    "zerde": ("Зерде-Фарма", ["Зерде-Фарма"]),
    "a2y": ("Аптека от А до Я", ["Аптека от А до Я"]),
    "msp": ("Аптека MSP (+Шымкент/+Астана)", ["Аптека MSP", "Аптека MSP Шымкент", "Аптека MSP Астана"]),
    "farm": ("ФАРМАКОМ", ["ФАРМАКОМ"]),
}

CITY_OPTIONS = ["Алматы", "Астана", "Шымкент"]


def build_expected_keyboard(token: str, selected: set[str] | None = None):
    selected = selected or set()

    def label(key: str):
        base = EXPECTED_OPTIONS[key][0]
        return f"{'✅' if key in selected else '▫️'} {base}"

    return {
        "inline_keyboard": [
            [
                {"text": label("zerde"), "callback_data": f"setexp:{token}:zerde"},
                {"text": label("a2y"), "callback_data": f"setexp:{token}:a2y"},
            ],
            [
                {"text": label("msp"), "callback_data": f"setexp:{token}:msp"},
                {"text": label("farm"), "callback_data": f"setexp:{token}:farm"},
            ],
            [
                {"text": "Выбрать всё", "callback_data": f"setexp:{token}:all"},
                {"text": "Сохранить", "callback_data": f"setexp:{token}:save"},
            ],
            [
                {"text": "Пропустить", "callback_data": f"setexp:{token}:skip"},
            ],
        ]
    }


def build_city_keyboard(token: str):
    return {
        "inline_keyboard": [
            [{"text": city, "callback_data": f"citysel:{token}:{city}"}] for city in CITY_OPTIONS
        ] + [
            [{"text": "Отмена", "callback_data": f"citysel:{token}:cancel"}]
        ]
    }


def status_text():
    with PROGRESS_LOCK:
        st = PROGRESS.copy()
    return (
        f"Статус: {st.get('status')}\n"
        f"Город: {st.get('city')}\n"
        f"Собрано: {st.get('done')}\n"
        f"Заметка: {st.get('note')}"
    )


def auto_alerts_sku_summary(chat_id: str) -> str:
    csv_path = user_extra_csv(chat_id)
    if not csv_path.exists():
        return "Временный список пуст."
    try:
        df = pd.read_csv(csv_path)
    except Exception as exc:
        return f"Не удалось прочитать SKU: {exc}"
    if df.empty:
        return "Временный список пуст."
    rows: list[str] = []
    for idx, row in df.head(EXTRA_PREVIEW_LIMIT).iterrows():
        label = str(row.get("product_name") or "").strip()
        if not label:
            label = str(row.get("product_link") or "").strip()
        expected = str(row.get("expected_sellers") or "").strip()
        suffix = f" — {expected}" if expected else ""
        rows.append(f"{idx+1}. {label}{suffix}")
    if len(df) > EXTRA_PREVIEW_LIMIT:
        rows.append(f"…и ещё {len(df) - EXTRA_PREVIEW_LIMIT} SKU")
    return "\n".join(rows)


def auto_alert_status_text(chat_id: str) -> str:
    stop_event = AUTO_STOPS.get(chat_id)
    thread = AUTO_THREADS.get(chat_id)
    ctx = AUTO_CONTEXT.get(chat_id) or {}
    active = bool(thread and thread.is_alive() and stop_event and not stop_event.is_set())
    if not active:
        return "Авто-алерты сейчас выключены.\nНажми «⏱️ Авто-алерты вкл» и выбери город."
    city = ctx.get("city") or "Алматы/Астана/Шымкент"
    interval = ctx.get("interval_sec")
    if isinstance(interval, int) and interval >= 60:
        interval_txt = f"{interval // 60} мин"
    elif isinstance(interval, int):
        interval_txt = f"{interval} сек"
    else:
        interval_txt = "не задан"
    sku_text = auto_alerts_sku_summary(chat_id)
    return f"Авто-алерты активны.\nГород: {city}\nИнтервал: {interval_txt}\nТовары:\n{sku_text}"


def extra_list_text(chat_id: str):
    csv_path = user_extra_csv(chat_id)
    if not csv_path.exists():
        return "Временный список пуст."
    try:
        df = pd.read_csv(csv_path)
    except Exception as exc:
        return f"Не удалось прочитать список: {exc}"
    if df.empty:
        return "Временный список пуст."
    rows = []
    for idx, row in df.head(EXTRA_PREVIEW_LIMIT).iterrows():
        name = str(row.get("product_name") or "").strip()
        link = str(row.get("product_link") or "").strip()
        exp = str(row.get("expected_sellers") or "").strip()
        label = name if name else link
        rows.append(f"{idx+1}. {label}\n{link}\n{exp or 'продавцы не заданы'}")
    extra = ""
    if len(df) > EXTRA_PREVIEW_LIMIT:
        extra = f"\n…и ещё {len(df) - EXTRA_PREVIEW_LIMIT} строк(и)"
    return "Временный список:\n" + "\n\n".join(rows) + extra


def latest_result_file() -> Path | None:
    results_dir = BASE_DIR / "RESULTS"
    if not results_dir.exists():
        return None
    marker = BASE_DIR / "state" / "latest_full.txt"
    if marker.exists():
        try:
            path_str = marker.read_text(encoding="utf-8").strip()
            if path_str:
                p = Path(path_str)
                if p.exists():
                    return p
        except Exception:
            pass
    candidates = sorted(results_dir.glob("kaspi_prices_*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0] if candidates else None


def chat_memory_path(chat_id: str) -> Path:
    USER_DATA_DIR.mkdir(parents=True, exist_ok=True)
    return USER_DATA_DIR / f"{chat_id}_memory.json"


def chat_ids_log_path() -> Path:
    USER_DATA_DIR.mkdir(parents=True, exist_ok=True)
    return USER_DATA_DIR / "chat_ids.txt"


def log_chat_id(chat_id: str):
    path = chat_ids_log_path()
    try:
        if path.exists():
            with open(path, "r", encoding="utf-8") as fh:
                ids = {line.strip() for line in fh if line.strip()}
        else:
            ids = set()
        if chat_id not in ids:
            ids.add(chat_id)
            with open(path, "w", encoding="utf-8") as fh:
                fh.write("\n".join(sorted(ids)))
    except Exception:
        pass


def load_chat_memory(chat_id: str) -> list[dict]:
    path = chat_memory_path(chat_id)
    if not path.exists():
        return []
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
            if isinstance(data, list):
                return data
    except Exception:
        return []
    return []


def save_chat_memory(chat_id: str, messages: list[dict]):
    path = chat_memory_path(chat_id)
    try:
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(messages, fh, ensure_ascii=False, indent=2)
    except Exception:
        pass


def price_query_response(text: str) -> str | None:
    low = text.lower()
    match = re.search(r"(цена|стоимость|узнать цену)[^\\n]{0,40}\\s+на\\s+(.+)", low)
    if match:
        product = match.group(2).strip(" .!?")
        if product:
            return (
                f"По запросу «{product}» на Kaspi обычно несколько позиций. "
                "Уточни ссылку или код товара, чтобы смотреть точную цену. "
                "Можешь сразу отправить /add <ссылка> | <название> | <продавцы>."
            )
    return None


def extract_kaspi_links(text: str, entities=None) -> list[str]:
    if not text:
        text = ""
    links: list[str] = []
    # 1) entities от Telegram (url или text_link)
    if entities:
        for ent in entities:
            etype = ent.get("type")
            if etype == "text_link" and ent.get("url"):
                links.append(ent["url"])
            elif etype == "url":
                try:
                    offset = int(ent.get("offset", 0))
                    length = int(ent.get("length", 0))
                    links.append(text[offset:offset + length])
                except Exception:
                    continue
    # 2) Регулярка по тексту/подписям
    candidates = re.findall(r"(https?://[^\s]+|www\.[^\s]+|kaspi\.kz/[^\s]+|l\.kaspi\.kz/[^\s]+)", text, flags=re.IGNORECASE)
    links.extend(candidates)

    seen = set()
    resolved: list[str] = []
    for raw in links:
        raw_clean = (raw or "").strip(" ,;<>\"'()[]")
        if "kaspi.kz" not in raw_clean.lower():
            continue
        fixed = resolve_kaspi_link(raw_clean)
        if fixed and fixed not in seen:
            seen.add(fixed)
            resolved.append(fixed)
    return resolved


def fetch_product_title(url: str, timeout: int = 10) -> str | None:
    try:
        resp = requests.get(
            url,
            timeout=timeout,
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/118.0"},
        )
    except Exception:
        return None
    if resp.status_code != 200:
        return None
    html = resp.text or ""
    match = re.search(r'<meta[^>]+property=["\']og:title["\'][^>]+content=["\']([^"\']+)["\']', html, re.IGNORECASE)
    if match:
        return unescape(match.group(1)).strip()
    title_match = re.search(r"<title>(.*?)</title>", html, re.IGNORECASE | re.DOTALL)
    if title_match:
        return unescape(title_match.group(1)).strip()
    return None


def start_link_preview(token: str, chat_id: str, links: list[str]):
    def _worker():
        try:
            main_link = resolve_kaspi_link(links[0])
        except Exception:
            main_link = links[0]
        product_title = fetch_product_title(main_link)
        token_id = uuid.uuid4().hex
        LINK_ACTIONS[token_id] = {"chat_id": chat_id, "link": main_link, "title": product_title}
        lines = []
        if len(links) > 1:
            lines.append(f"Нашёл {len(links)} ссылок. Работаю с первой:")
        else:
            lines.append("Нашёл ссылку на товар:")
        if product_title:
            lines.append(f"Товар: {product_title}")
        lines.append(main_link)
        lines.append("Выбери действие:")
        try:
            send_bot_message(
                token,
                chat_id,
                "\n".join(lines),
                reply_markup=build_link_actions_keyboard(token_id),
            )
        except Exception as exc:
            print(f"⚠️ Не удалось отправить ссылочный превью: {exc}")

    threading.Thread(target=_worker, daemon=True).start()


def _call_openai_chat(
    messages: list[dict],
    temperature: float = 0.2,
    max_tokens: int = 300,
    timeout: int = 20,
    model: str | None = None,
):
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        return None, "API ключ OpenAI не настроен."
    payload = {
        "model": model or OPENAI_MODEL,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
    }
    headers = {"Authorization": f"Bearer {api_key}"}
    try:
        resp = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload, timeout=timeout)
        data = resp.json()
    except Exception as exc:
        return None, f"GPT запрос не удался: {exc}"
    if resp.status_code == 200 and data.get("choices"):
        return data["choices"][0]["message"]["content"].strip(), None
    return None, f"GPT ошибка: {resp.status_code} {data}"


def gpt_answer(question: str, chat_id: str) -> str:
    with MEMORY_LOCK:
        history = load_chat_memory(chat_id)
        # ограничиваем историю, чтобы не раздуть запрос
        short_history = history[-12:]

    messages = [
        {
            "role": "system",
            "content": (
                "Ты Джарвис из Iron Man, специалист по e-commerce Kaspi. "
                "Отвечай кратко, чётко, на русском. "
                "Помогай пользоваться ботом: /add, /run, /run_extra, /run_alerts, авто-алерты, выбор города, выбор продавцов. "
                "Давай советы по отслеживанию цен/продавцов, но не выдумывай данные."
            ),
        }
    ]

    # Добавляем историю в prompt
    for msg in short_history:
        if msg.get("role") in {"user", "assistant"} and msg.get("content"):
            messages.append({"role": msg["role"], "content": msg["content"]})

    messages.append({"role": "user", "content": question})

    answer, error = _call_openai_chat(messages, temperature=0.2, max_tokens=300, timeout=15)
    if not answer:
        return error or "GPT запрос не удался."
    with MEMORY_LOCK:
        history = load_chat_memory(chat_id)
        history.append({"role": "user", "content": question})
        history.append({"role": "assistant", "content": answer})
        history = history[-20:]
        save_chat_memory(chat_id, history)
    return answer


def _parse_chat_ids(value: str | None) -> set[str]:
    return {
        chat.strip()
        for chat in re.split(r"[,;\s]+", value or "")
        if chat and chat.strip()
    }


def _agent_allowed_chats() -> set[str]:
    load_env_from_file()
    allowed = _parse_chat_ids(os.environ.get("AGENT_ALLOWED_CHAT_IDS", ""))
    if allowed:
        return allowed
    fallback = (os.environ.get("TELEGRAM_CHAT_ID") or "").strip()
    return {fallback} if fallback else set()


def _agent_chat_allowed(chat_id: str) -> bool:
    allowed = _agent_allowed_chats()
    if not allowed:
        return False
    return str(chat_id) in allowed


def _ensure_agent_default_mode(chat_id: str):
    if not AGENT_AUTO_MODE_DEFAULT:
        return
    if not _agent_chat_allowed(chat_id):
        return
    with AGENT_STATE_LOCK:
        AGENT_MODE_CHATS.add(chat_id)


def _agent_mode_enabled(chat_id: str) -> bool:
    with AGENT_STATE_LOCK:
        return chat_id in AGENT_MODE_CHATS


def _set_agent_mode(chat_id: str, enabled: bool):
    with AGENT_STATE_LOCK:
        if enabled:
            AGENT_MODE_CHATS.add(chat_id)
        else:
            AGENT_MODE_CHATS.discard(chat_id)


def _command_payload(text: str, command_name: str) -> str:
    return re.sub(fr"^{re.escape(command_name)}(@\S+)?", "", text, flags=re.IGNORECASE).strip()


def _limit_text(text: str, max_chars: int) -> str:
    if not text:
        return ""
    if len(text) <= max_chars:
        return text
    tail = len(text) - max_chars
    return text[:max_chars] + f"\n... [обрезано {tail} символов]"


def _is_agent_command_safe(command: str) -> tuple[bool, str | None]:
    cmd = (command or "").strip()
    if not cmd:
        return False, "пустая команда"
    if "\n" in cmd:
        return False, "многострочные команды не поддерживаются"
    if len(cmd) > 500:
        return False, "слишком длинная команда"
    for pattern, reason in AGENT_BLOCK_PATTERNS:
        if re.search(pattern, cmd, flags=re.IGNORECASE):
            return False, reason
    return True, None


def _allow_nonzero_for_probe(command: str, code: int, output: str) -> bool:
    if code != 1:
        return False
    cmd = (command or "").lower()
    out = (output or "").strip()
    if out:
        return False
    probe_patterns = [
        r"\bgrep\b",
        r"\brg\b",
        r"\bfind\b",
        r"\btest\s+-[fd]\b",
        r"\[\s*-[fd]\s+",
        r"\bps\b",
    ]
    return any(re.search(pattern, cmd) for pattern in probe_patterns)


def _parse_agent_plan(raw: str) -> dict | None:
    if not raw:
        return None
    cleaned = raw.strip().replace("\ufeff", "")
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?", "", cleaned, flags=re.IGNORECASE).strip()
        cleaned = re.sub(r"```$", "", cleaned).strip()
    start = cleaned.find("{")
    end = cleaned.rfind("}")
    if start == -1 or end == -1 or end < start:
        return None
    try:
        payload = json.loads(cleaned[start : end + 1])
    except Exception:
        return None
    if not isinstance(payload, dict):
        return None
    raw_steps = payload.get("steps")
    if not isinstance(raw_steps, list):
        raw_steps = []
    steps: list[dict[str, str]] = []
    for step in raw_steps:
        if not isinstance(step, dict):
            continue
        title = str(step.get("title") or "").strip()
        command = str(step.get("command") or "").strip()
        if not command:
            continue
        if not title:
            title = f"Шаг {len(steps) + 1}"
        steps.append({"title": title[:140], "command": command})
        if len(steps) >= AGENT_MAX_STEPS:
            break
    summary = str(payload.get("summary") or "").strip()
    return {"summary": summary[:400], "steps": steps}


def _fallback_agent_steps(task_text: str) -> list[dict[str, str]]:
    low = (task_text or "").lower()
    if any(key in low for key in ("процесс", "запущен", "pid", "бот", "service", "сервис")):
        return [
            {
                "title": "Проверить процессы бота",
                "command": 'ps -ef | rg -n "python kaspi_bot\\.py|run_bot\\.sh" -S',
            }
        ]
    if any(key in low for key in ("лог", "ошиб", "error", "tail")):
        return [
            {
                "title": "Проверить последние строки лога",
                "command": "if [ -f kaspi-scraper/logs/kaspi_bot.log ]; then tail -n 80 kaspi-scraper/logs/kaspi_bot.log; elif [ -f logs/kaspi_bot.log ]; then tail -n 80 logs/kaspi_bot.log; else echo 'log not found'; fi",
            }
        ]
    if any(key in low for key in ("папк", "директор", "файл", "список", "ls", "pwd")):
        return [
            {
                "title": "Проверить рабочую папку",
                "command": "pwd && ls -la | sed -n '1,40p'",
            }
        ]
    return [
        {
            "title": "Собрать базовый контекст",
            "command": "pwd && ls -la | sed -n '1,40p'",
        }
    ]


def _preset_agent_plan(task_text: str) -> dict | None:
    low = (task_text or "").lower()
    if "бот" in low and any(word in low for word in ("лог", "ошиб", "error")):
        steps = [
            {
                "title": "Проверить, запущен ли Telegram-бот",
                "command": 'ps -ef | rg -n "python kaspi_bot\\.py|run_bot\\.sh" -S || true',
            },
            {
                "title": "Показать последние 30 строк лога бота",
                "command": "if [ -f /home/vas/kaspi-scraper/logs/kaspi_bot.log ]; then tail -n 30 /home/vas/kaspi-scraper/logs/kaspi_bot.log; elif [ -f /home/vas/kaspi-scraper/logs/kaspi_bot_runner.log ]; then tail -n 30 /home/vas/kaspi-scraper/logs/kaspi_bot_runner.log; else echo 'log not found'; fi",
            },
        ]
        return {
            "summary": "Проверил процессы бота и вывел хвост актуального лога.",
            "steps": steps,
        }
    return None


def _build_agent_plan(task_text: str, chat_id: str) -> tuple[dict | None, str | None]:
    preset = _preset_agent_plan(task_text)
    if preset:
        return preset, None
    planner_prompt = (
        "Ты инженерный агент для Linux shell. "
        "Нужно превратить задачу пользователя в короткий безопасный план выполнения. "
        "Верни только JSON без markdown: "
        '{"summary":"...", "steps":[{"title":"...", "command":"..."}]}. '
        f"Максимум {AGENT_MAX_STEPS} шагов. "
        "Команды должны быть неинтерактивными и запускаться в bash. "
        "Используй проверку контекста перед выводами (например, ls/rg/ps/sed/cat). "
        "Если задача описана общими словами, не отказывайся: сделай безопасные предположения и начни с диагностики контекста. "
        "Не используй sudo, rm -rf, shutdown/reboot, mkfs, dd if=. "
        "Если задачу выполнить нельзя, верни steps пустым списком и укажи причину в summary."
    )
    user_prompt = (
        f"Рабочая папка: {AGENT_WORKDIR}\n"
        f"chat_id: {chat_id}\n"
        f"Задача: {task_text}"
    )
    msgs = [{"role": "system", "content": planner_prompt}, {"role": "user", "content": user_prompt}]
    raw_plan, error = _call_openai_chat(
        msgs,
        temperature=0.1,
        max_tokens=900,
        timeout=35,
        model=AGENT_MODEL,
    )
    if not raw_plan and AGENT_MODEL != OPENAI_MODEL:
        raw_plan, error = _call_openai_chat(
            msgs,
            temperature=0.1,
            max_tokens=900,
            timeout=35,
            model=OPENAI_MODEL,
        )
    if not raw_plan:
        return None, error or "Не удалось получить план от GPT."
    parsed = _parse_agent_plan(raw_plan)
    if not parsed:
        return None, "GPT вернул план в неподдерживаемом формате."
    if not parsed.get("steps"):
        parsed["steps"] = _fallback_agent_steps(task_text)
        if not parsed.get("summary"):
            parsed["summary"] = "Запустил fallback-план: GPT не предложил шаги."
    return parsed, None


def _run_agent_command(command: str) -> tuple[int, str, float, bool]:
    started = time.time()
    try:
        proc = subprocess.run(
            ["bash", "-lc", command],
            cwd=str(AGENT_WORKDIR),
            env=os.environ.copy(),
            capture_output=True,
            text=True,
            timeout=AGENT_CMD_TIMEOUT_SEC,
        )
        elapsed = time.time() - started
        out_parts = []
        if proc.stdout:
            out_parts.append(proc.stdout.strip())
        if proc.stderr:
            out_parts.append(proc.stderr.strip())
        output = _limit_text("\n".join(p for p in out_parts if p), AGENT_OUTPUT_LIMIT)
        return proc.returncode, output, elapsed, False
    except subprocess.TimeoutExpired as exc:
        elapsed = time.time() - started
        out_parts = []
        if exc.stdout:
            out_parts.append(str(exc.stdout).strip())
        if exc.stderr:
            out_parts.append(str(exc.stderr).strip())
        output = _limit_text("\n".join(p for p in out_parts if p), AGENT_OUTPUT_LIMIT)
        return 124, output, elapsed, True


def agent_status_text(chat_id: str) -> str:
    allowed = _agent_chat_allowed(chat_id)
    mode = "ON" if _agent_mode_enabled(chat_id) else "OFF"
    with AGENT_STATE_LOCK:
        thread = AGENT_THREADS.get(chat_id)
        busy = bool(thread and thread.is_alive())
    return (
        f"Agent-доступ: {'разрешён' if allowed else 'запрещён'}\n"
        f"Agent авто-режим: {mode}\n"
        f"Agent занят: {'да' if busy else 'нет'}\n"
        f"Рабочая папка: {AGENT_WORKDIR}"
    )


def extended_bot_help_text() -> str:
    return (
        f"{bot_help_text()}\n\n"
        "Agent-режим (выполнение задач из текста):\n"
        "• /agent <задача> — выполнить задачу по шагам с прогрессом\n"
        "• /agent_on — включить авто-режим (любой обычный текст = задача)\n"
        "• /agent_off — выключить авто-режим\n"
        "• /agent_status — статус agent-режима\n"
        "Пример: /agent проверь, запущен ли bot, и отправь короткий статус."
    )


def start_agent_task(token: str, chat_id: str, task_text: str):
    task = (task_text or "").strip()
    if not task:
        send_bot_message(token, chat_id, "Напиши задачу после /agent", reply_markup=build_keyboard())
        return
    if not _agent_chat_allowed(chat_id):
        send_bot_message(token, chat_id, "Agent-режим не разрешён для этого чата.", reply_markup=build_keyboard())
        return
    with AGENT_STATE_LOCK:
        running = AGENT_THREADS.get(chat_id)
        if running and running.is_alive():
            send_bot_message(token, chat_id, "Agent уже выполняет задачу. Дождись завершения.", reply_markup=build_keyboard())
            return

    def _worker():
        started_at = time.time()
        _log_event("agent_task_start", {"chat_id": chat_id, "task": task[:500]})
        try:
            send_bot_message(
                token,
                chat_id,
                f"🧭 Принял задачу:\n{task}\n\nСоставляю план...",
                reply_markup=build_keyboard(),
            )
            plan, err = _build_agent_plan(task, chat_id)
            if not plan:
                send_bot_message(token, chat_id, f"❌ План не собрался: {err}", reply_markup=build_keyboard())
                return
            steps = plan.get("steps") or []
            if not steps:
                summary = plan.get("summary") or "План пустой: не нашёл безопасных шагов."
                send_bot_message(token, chat_id, f"ℹ️ {summary}", reply_markup=build_keyboard())
                return

            plan_lines = [f"🗺 План ({len(steps)} шагов):"]
            for idx, step in enumerate(steps, start=1):
                plan_lines.append(f"{idx}. {step.get('title')}")
            if plan.get("summary"):
                plan_lines.append(f"Итоговая цель: {plan.get('summary')}")
            send_bot_message(token, chat_id, "\n".join(plan_lines), reply_markup=build_keyboard())

            for idx, step in enumerate(steps, start=1):
                title = str(step.get("title") or f"Шаг {idx}")
                command = str(step.get("command") or "").strip()
                safe, reason = _is_agent_command_safe(command)
                if not safe:
                    send_bot_message(
                        token,
                        chat_id,
                        f"🛑 Шаг {idx} заблокирован: {reason}\nКоманда: {command}",
                        reply_markup=build_keyboard(),
                    )
                    return
                send_bot_message(
                    token,
                    chat_id,
                    f"▶️ Шаг {idx}/{len(steps)}: {title}\nКоманда: {command}",
                    reply_markup=build_keyboard(),
                )
                code, output, elapsed, timed_out = _run_agent_command(command)
                result_lines = [f"{'✅' if code == 0 else '❌'} Шаг {idx}/{len(steps)} завершён за {elapsed:.1f}с (exit {code})"]
                if timed_out:
                    result_lines.append(f"Таймаут: {AGENT_CMD_TIMEOUT_SEC}с")
                result_lines.append(f"Команда: {command}")
                if output:
                    result_lines.append("Вывод:")
                    result_lines.append(output)
                else:
                    result_lines.append("Вывод: пусто")
                send_chunked_message(token, chat_id, "\n".join(result_lines))
                if code != 0 and _allow_nonzero_for_probe(command, code, output):
                    send_bot_message(
                        token,
                        chat_id,
                        f"⚠️ Шаг {idx}/{len(steps)} вернул exit {code} (похоже, просто нет совпадений). Продолжаю.",
                        reply_markup=build_keyboard(),
                    )
                    continue
                if code != 0:
                    send_bot_message(
                        token,
                        chat_id,
                        "Остановил выполнение после ошибки на этом шаге.",
                        reply_markup=build_keyboard(),
                    )
                    return

            total = time.time() - started_at
            summary = str(plan.get("summary") or "").strip()
            done_text = f"✅ Agent завершил задачу за {total:.1f}с."
            if summary:
                done_text += f"\nИтог: {summary}"
            send_bot_message(token, chat_id, done_text, reply_markup=build_keyboard())
        except Exception as exc:
            send_bot_message(token, chat_id, f"❌ Agent ошибка: {exc}", reply_markup=build_keyboard())
        finally:
            with AGENT_STATE_LOCK:
                current = AGENT_THREADS.get(chat_id)
                if current is threading.current_thread():
                    AGENT_THREADS.pop(chat_id, None)
            _log_event("agent_task_end", {"chat_id": chat_id, "task": task[:500]})

    th = threading.Thread(target=_worker, daemon=True)
    with AGENT_STATE_LOCK:
        AGENT_THREADS[chat_id] = th
    th.start()


def request_research(token: str, chat_id: str, query: str, mode: str | None = None):
    query = (query or "").strip()
    if not query:
        send_bot_message(token, chat_id, "Укажи тему: /research <запрос>", reply_markup=build_keyboard())
        return
    if not TAVILY_API_KEY:
        send_bot_message(token, chat_id, "TAVILY_API_KEY не настроен, ресёрч недоступен.", reply_markup=build_keyboard())
        return
    mode_key, preset = research_mode_info(mode)
    max_results = preset.get("max_results", RESEARCH_FAST_RESULTS)
    cache_hit = research_cache_available(query, max_results)
    if not cache_hit:
        allowed, wait = research_quota_status(chat_id)
        if not allowed:
            send_bot_message(
                token,
                chat_id,
                f"Превышен лимит ресёрчей. Подожди {format_wait_time(wait)} и попробуй снова.",
                reply_markup=build_keyboard(),
            )
            return
    start_research_worker(
        token,
        chat_id,
        query,
        followup=False,
        require_quota=not cache_hit,
        mode=mode_key,
        max_results=max_results,
        preset=preset,
    )


def request_research_followup(token: str, chat_id: str, question: str):
    question = (question or "").strip()
    if not question:
        send_bot_message(token, chat_id, "Напиши уточняющий вопрос после /research_more", reply_markup=build_keyboard())
        return
    context = RESEARCH_CONTEXT.get(chat_id) or {}
    start_research_worker(
        token,
        chat_id,
        question,
        followup=True,
        require_quota=False,
        mode=context.get("mode"),
        max_results=context.get("max_results"),
        preset=None,
    )


def format_rating(value) -> str:
    if value is None:
        return "—"
    try:
        num = float(value)
    except Exception:
        return str(value)
    return f"{num:.1f}".rstrip("0").rstrip(".")


def start_research_worker(
    token: str,
    chat_id: str,
    query: str,
    followup: bool,
    require_quota: bool,
    mode: str | None,
    max_results: int | None,
    preset: dict | None,
):
    def _run():
        with RESEARCH_LOCK:
            if chat_id in RESEARCH_IN_PROGRESS:
                send_bot_message(token, chat_id, "Другое исследование ещё выполняется, дождись завершения.", reply_markup=build_keyboard())
                return
            RESEARCH_IN_PROGRESS.add(chat_id)
        try:
            mode_key, mode_info = research_mode_info(mode)
            preset_info = preset or mode_info
            mode_label = preset_info.get("label", "Ресёрч")
            mode_max_results = max_results or preset_info.get("max_results") or RESEARCH_FAST_RESULTS
            if followup:
                context = RESEARCH_CONTEXT.get(chat_id)
                if not context or not context.get("results"):
                    send_bot_message(token, chat_id, "Нет свежего ресёрча. Сначала запусти /research.", reply_markup=build_keyboard())
                    return
                send_bot_message(token, chat_id, "✏️ Думаю над уточняющим вопросом…", reply_markup=build_keyboard())
                base_query = context.get("query") or ""
                results = context.get("results") or []
                if not results:
                    send_bot_message(token, chat_id, "Нет источников для уточнения, запусти ресёрч заново.", reply_markup=build_keyboard())
                    return
                summary = research_summary_via_gpt(query, results, base_query=base_query or query)
                label = f"🔁 Доп. вопрос ({mode_info.get('label', 'Ресёрч')}) по теме «{base_query or 'без названия'}»"
            else:
                send_bot_message(token, chat_id, f"{mode_label}: собираю источники по запросу «{query}»…", reply_markup=build_keyboard())
                results = get_cached_research(query, mode_max_results)
                if not results:
                    if require_quota:
                        allowed, wait = reserve_research_slot(chat_id)
                        if not allowed:
                            send_bot_message(
                                token,
                                chat_id,
                                f"Превышен лимит ресёрчей. Подожди {format_wait_time(wait)} и попробуй снова.",
                                reply_markup=build_keyboard(),
                            )
                            return
                    results = tavily_search(query, max_results=mode_max_results)
                    if not results:
                        send_bot_message(token, chat_id, "Не удалось найти результаты.", reply_markup=build_keyboard())
                        return
                    set_research_cache(query, results, mode_max_results)
                with RESEARCH_LOCK:
                    RESEARCH_CONTEXT[chat_id] = {
                        "query": query,
                        "results": results,
                        "ts": time.time(),
                        "mode": mode_key,
                        "max_results": mode_max_results,
                    }
                summary = research_summary_via_gpt(query, results, base_query=query)
                label = f"🔎 Результаты ({mode_label}) по запросу: {query}"
            sources_text = format_research_sources(results)
            message = f"{label}\n\n{summary}"
            if sources_text:
                message += f"\n\nИсточники:\n{sources_text}"
            if not followup:
                message += "\n\nЧтобы задать уточняющий вопрос, используй /research_more <вопрос>."
            send_chunked_message(token, chat_id, message)
        except Exception as exc:
            send_bot_message(token, chat_id, f"❌ Ресёрч не удался: {exc}", reply_markup=build_keyboard())
        finally:
            with RESEARCH_LOCK:
                RESEARCH_IN_PROGRESS.discard(chat_id)

    threading.Thread(target=_run, daemon=True).start()


def start_scrape(
    token: str,
    chat_id: str,
    include_base: bool,
    alert_only: bool = False,
    cities=None,
    extra_csv_paths=None,
    change_alerts: bool = True,
):
    global SCRAPER_THREAD, SCRAPER_STOP_EVENT
    if SCRAPER_THREAD and SCRAPER_THREAD.is_alive():
        send_bot_message(token, chat_id, "Скрейп уже запущен, дождись завершения текущего.", reply_markup=build_keyboard())
        return

    stop_event = threading.Event()
    SCRAPER_STOP_EVENT = stop_event

    def _run():
        try:
            send_bot_message(token, chat_id, "⏳ Запускаю скрейп...", reply_markup=build_keyboard())
            csv_sources = extra_csv_paths if extra_csv_paths is not None else user_extra_csv(chat_id)
            run_kaspi_scrape(
                include_base=include_base,
                chat_id=chat_id,
                extra_csv_paths=csv_sources,
                stop_event=stop_event,
                progress_callback=progress_callback,
                cities=cities or None,
                alert_only=alert_only,
                change_alerts=change_alerts,
            )
            if stop_event.is_set():
                send_bot_message(token, chat_id, "⏹ Остановлено", reply_markup=build_keyboard())
            else:
                send_bot_message(
                    token,
                    chat_id,
                    "✅ Скрейп завершён (только алерты)" if alert_only else "✅ Скрейп завершён",
                    reply_markup=build_keyboard(),
                )
        except Exception as exc:
            send_bot_message(token, chat_id, f"❌ Ошибка скрейпа: {exc}", reply_markup=build_keyboard())
        finally:
            with PROGRESS_LOCK:
                PROGRESS.update({"status": "idle", "city": "", "done": 0, "note": ""})
            SCRAPER_THREAD = None
            SCRAPER_STOP_EVENT = None

    SCRAPER_THREAD = threading.Thread(target=_run, daemon=True)
    SCRAPER_THREAD.start()


def start_keyword_search(token: str, chat_id: str):
    global SCRAPER_THREAD, SCRAPER_STOP_EVENT
    if SCRAPER_THREAD and SCRAPER_THREAD.is_alive():
        send_bot_message(token, chat_id, "Скрейп уже запущен, дождись завершения текущего.", reply_markup=build_keyboard())
        return

    keyword_path = keyword_csv_path()
    if not keyword_path.exists():
        msg = (
            "Не нашёл файл ключевых слов.\n"
            f"Ожидаю: {keyword_path}\n"
            "Создай search_keywords.csv с колонками keyword,search_url."
        )
        send_bot_message(token, chat_id, msg, reply_markup=build_keyboard())
        return

    stop_event = threading.Event()
    SCRAPER_STOP_EVENT = stop_event

    def _run():
        try:
            send_bot_message(token, chat_id, "⏳ Запускаю поиск по ключевым словам...", reply_markup=build_keyboard())
            csv_sources = user_extra_csv(chat_id)
            report_path = run_keyword_search_report(
                include_base=True,
                chat_id=chat_id,
                extra_csv_paths=csv_sources,
            )
            if report_path:
                send_bot_message(token, chat_id, "✅ Поиск по ключам завершён, файл отправлен.", reply_markup=build_keyboard())
            else:
                send_bot_message(
                    token,
                    chat_id,
                    "Нет данных для отчёта. Проверь search_keywords.csv и список SKU.",
                    reply_markup=build_keyboard(),
                )
        except Exception as exc:
            send_bot_message(token, chat_id, f"❌ Ошибка поиска по ключам: {exc}", reply_markup=build_keyboard())
        finally:
            SCRAPER_THREAD = None
            SCRAPER_STOP_EVENT = None

    SCRAPER_THREAD = threading.Thread(target=_run, daemon=True)
    SCRAPER_THREAD.start()


def start_natasha_price_scrape(token: str, chat_id: str, requester_id: str | None = None):
    global SCRAPER_THREAD, SCRAPER_STOP_EVENT

    if not _natasha_price_admin_allowed(chat_id, requester_id):
        try:
            if _send_latest_natasha_price(token, chat_id):
                return
        except Exception as exc:
            send_bot_message(
                token,
                chat_id,
                f"Не смог отправить последний скрейп prices_natasha: {exc}",
                reply_markup=build_keyboard(),
            )
            return
        send_bot_message(
            token,
            chat_id,
            "Пока нет готового скрейпа по prices_natasha.csv. Попроси админа запустить обновление.",
            reply_markup=build_keyboard(),
        )
        return

    csv_path = natasha_price_csv_path()
    if not csv_path.exists():
        send_bot_message(token, chat_id, f"Не нашёл файл для скрейпа: {csv_path}", reply_markup=build_keyboard())
        return

    if SCRAPER_THREAD and SCRAPER_THREAD.is_alive():
        send_bot_message(
            token,
            chat_id,
            "Скрейп уже запущен, дождись завершения текущего. Последний готовый файл могу отправить по кнопке.",
            reply_markup=build_keyboard(),
        )
        return

    stop_event = threading.Event()
    SCRAPER_STOP_EVENT = stop_event
    target_chat_id = str(requester_id or chat_id)

    def _run():
        global SCRAPER_THREAD, SCRAPER_STOP_EVENT
        stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        results_dir = BASE_DIR / "RESULTS"
        results_dir.mkdir(parents=True, exist_ok=True)
        output_path = results_dir / f"{NATASHA_PRICE_OUTPUT_PREFIX}_{stamp}.xlsx"
        latest_copy = BASE_DIR / f"{NATASHA_PRICE_OUTPUT_PREFIX}.xlsx"
        log_path = BASE_DIR / "logs" / f"{NATASHA_PRICE_OUTPUT_PREFIX}_{stamp}.log"
        log_path.parent.mkdir(parents=True, exist_ok=True)
        script_path = BASE_DIR / "scrape_prices_natasha_csv.py"
        cmd = [
            sys.executable,
            str(script_path),
            "--input",
            str(csv_path),
            "--output",
            str(output_path),
            "--workers",
            str(NATASHA_PRICE_WORKERS),
            "--iteka-max-pages",
            str(NATASHA_PRICE_ITEKA_MAX_PAGES),
        ]
        last_lines: list[str] = []
        try:
            with PROGRESS_LOCK:
                PROGRESS.update({"status": "running", "city": "", "done": 0, "note": "prices_natasha"})
            send_bot_message(
                token,
                chat_id,
                f"⏳ Запускаю скрейп prices_natasha.csv\nФайл: {csv_path.name}\nРезультат пришлю в Telegram после завершения.",
                reply_markup=build_keyboard(),
            )
            with open(log_path, "w", encoding="utf-8") as log_fh:
                process = subprocess.Popen(
                    cmd,
                    cwd=str(BASE_DIR),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1,
                )
                assert process.stdout is not None
                while True:
                    line = process.stdout.readline()
                    if line:
                        log_fh.write(line)
                        log_fh.flush()
                        last_lines.append(line.rstrip())
                        last_lines = last_lines[-20:]
                    if stop_event.is_set() and process.poll() is None:
                        process.terminate()
                        last_lines.append("Stopped by user request")
                    if not line and process.poll() is not None:
                        break
                    if not line:
                        time.sleep(1)
                return_code = process.wait()

            if stop_event.is_set():
                send_bot_message(token, chat_id, "⏹ Скрейп prices_natasha остановлен.", reply_markup=build_keyboard())
                return

            if return_code != 0 or not output_path.exists():
                reason = "\n".join(last_lines[-8:]) or f"process return code {return_code}"
                send_bot_message(
                    token,
                    chat_id,
                    f"❌ Скрейп prices_natasha не завершился.\nПричина/лог:\n{reason[:3000]}\n\nЛог: {log_path}",
                    reply_markup=build_keyboard(),
                )
                return

            try:
                shutil.copy2(output_path, latest_copy)
            except Exception:
                pass
            caption = f"✅ Скрейп prices_natasha готов — {_natasha_price_time_label(output_path)}"
            try:
                send_bot_file(token, target_chat_id, output_path, caption=caption)
            except Exception as exc:
                if target_chat_id != chat_id:
                    send_bot_file(token, chat_id, output_path, caption=caption)
                else:
                    raise exc
            if target_chat_id != chat_id:
                send_bot_message(token, chat_id, "✅ Скрейп готов, файл отправил тебе в личку.", reply_markup=build_keyboard())
        except Exception as exc:
            send_bot_message(token, chat_id, f"❌ Ошибка скрейпа prices_natasha: {exc}", reply_markup=build_keyboard())
        finally:
            with PROGRESS_LOCK:
                PROGRESS.update({"status": "idle", "city": "", "done": 0, "note": ""})
            SCRAPER_THREAD = None
            SCRAPER_STOP_EVENT = None

    SCRAPER_THREAD = threading.Thread(target=_run, daemon=True)
    SCRAPER_THREAD.start()


def start_competitor_scrape(token: str, chat_id: str, requester_id: str | None = None):
    start_natasha_price_scrape(token, chat_id, requester_id=requester_id)


def start_legacy_competitor_scrape(token: str, chat_id: str):
    csv_paths, missing = competitor_csv_paths()
    if not csv_paths:
        if missing:
            missing_lines = "\n".join(f"• {path}" for path in missing)
            msg = (
                "Не нашёл файл(ы) со списком конкурентов:\n"
                f"{missing_lines}\n"
                "Укажи путь в переменной COMPETITOR_PRODUCTS_CSV или создай competitor_products.csv рядом со скриптом."
            )
        else:
            msg = (
                "Не найден список конкурентов. "
                "Укажи COMPETITOR_PRODUCTS_CSV или создай competitor_products.csv рядом со скриптом."
            )
        send_bot_message(token, chat_id, msg, reply_markup=build_keyboard())
        return
    start_scrape(token, chat_id, include_base=False, extra_csv_paths=csv_paths)


def start_mission_december_scrape(token: str, chat_id: str):
    try:
        send_bot_message(
            token,
            chat_id,
            "⏳ Запускаю миссию…",
            reply_markup=build_keyboard(),
        )
    except Exception as exc:
        print(f"⚠️ Could not send mission start message: {exc}")
    # Если пользователь не в списке админов миссии — отдаём последний готовый файл, без запуска скрейпа.
    if chat_id not in MISSION_ADMIN_CHATS:
        if _send_latest_mission(token, chat_id):
            MISSION_SUBSCRIBERS.add(chat_id)
            ensure_mission_push_loop(token, interval_sec=1800)
            return
        else:
            send_bot_message(
                token,
                chat_id,
                "Нет свежего отчёта миссии. Попроси админа обновить или подождать.",
                reply_markup=build_keyboard(),
            )
            return

    # Админ: сразу отправляем свежий отчёт (если есть) и запускаем автоскрейп раз в час.
    sent = _send_latest_mission(token, chat_id)
    if not sent:
        try:
            send_bot_message(
                token,
                chat_id,
                "Свежего отчёта нет. Запускаю новый сбор, отправлю после завершения.",
                reply_markup=build_keyboard(),
            )
        except Exception as exc:
            print(f"⚠️ Could not send mission missing-report message: {exc}")
    if MISSION_BOT_AUTO_ENABLED:
        mission_interval = int(
            os.environ.get("MISSION_APRIL_INTERVAL_SEC")
            or os.environ.get("MISSION_FEBRUARY_INTERVAL_SEC")
            or "1800"
        )
        start_mission_auto(token, chat_id, interval_sec=mission_interval)


def start_main_scrape(token: str, chat_id: str):
    # если не админ — отдаём свежий ежедневный файл, не запускаем скрейп
    if chat_id not in MAIN_ADMIN_CHATS:
        latest = latest_report_file(MAIN_FILE_PREFIX)
        if latest and latest.exists():
            ts = datetime.fromtimestamp(latest.stat().st_mtime).strftime("%d.%m %H:%M")
            caption = f"Свежий отчёт (ежедневный) — {ts}"
            try:
                send_bot_file(token, chat_id, latest, caption=caption)
                return
            except Exception as exc:
                send_bot_message(
                    token,
                    chat_id,
                    f"Не смог отправить файл: {exc}. Попробуй позже или попроси админа обновить.",
                    reply_markup=build_keyboard(),
                )
                return
        send_bot_message(
            token,
            chat_id,
            "Нет свежего ежедневного отчёта. Попроси админа запустить /run.",
            reply_markup=build_keyboard(),
        )
        return

    # админам — обычный запуск полного скрейпа
    start_scrape(token, chat_id, include_base=True)


def request_stop(token: str, chat_id: str):
    global SCRAPER_STOP_EVENT
    if SCRAPER_THREAD and SCRAPER_THREAD.is_alive() and SCRAPER_STOP_EVENT:
        SCRAPER_STOP_EVENT.set()
        send_bot_message(token, chat_id, "⏹ Запросили остановку. Ждите завершения текущих задач.", reply_markup=build_keyboard())
    else:
        send_bot_message(token, chat_id, "Сейчас ничего не запущено.", reply_markup=build_keyboard())


def start_auto_alerts(token: str, chat_id: str, interval_sec: int, city: str | None):
    stop_event = AUTO_STOPS.get(chat_id)
    if stop_event:
        stop_event.set()
    stop_event = threading.Event()
    AUTO_STOPS[chat_id] = stop_event
    ctx_payload = {"city": city, "interval_sec": interval_sec}
    AUTO_CONTEXT[chat_id] = ctx_payload

    def _loop(stop_event=stop_event, ctx_payload=ctx_payload):
        send_bot_message(token, chat_id, f"⏱️ Авто-алерты включены, каждые {interval_sec} сек, город {city or 'Алматы/Астана/Шымкент'}", reply_markup=build_keyboard())
        while not stop_event.is_set():
            try:
                run_kaspi_scrape(
                    include_base=False,
                    chat_id=chat_id,
                    extra_csv_paths=user_extra_csv(chat_id),
                    cities=[city] if city else None,
                    alert_only=True,
                    priority="secondary",
                )
            except Exception as exc:
                send_bot_message(token, chat_id, f"❌ Ошибка авто-алертов: {exc}", reply_markup=build_keyboard())
            # ожидание с возможностью прерывания
            for _ in range(interval_sec):
                if stop_event.is_set():
                    break
                time.sleep(1)
        send_bot_message(token, chat_id, "✋ Авто-алерты остановлены.", reply_markup=build_keyboard())
        if AUTO_CONTEXT.get(chat_id) is ctx_payload:
            AUTO_CONTEXT.pop(chat_id, None)
        if AUTO_THREADS.get(chat_id) is threading.current_thread():
            AUTO_THREADS.pop(chat_id, None)
        if AUTO_STOPS.get(chat_id) is stop_event:
            AUTO_STOPS.pop(chat_id, None)

    th = threading.Thread(target=_loop, daemon=True)
    AUTO_THREADS[chat_id] = th
    th.start()


def stop_auto_alerts(token: str, chat_id: str):
    stop_event = AUTO_STOPS.get(chat_id)
    if stop_event:
        stop_event.set()
        AUTO_CONTEXT.pop(chat_id, None)
    else:
        send_bot_message(token, chat_id, "Авто-алерты не запущены.", reply_markup=build_keyboard())


def request_analysis_link(token: str, chat_id: str):
    ANALYSIS_REQUESTS[chat_id] = {"status": "await_link"}
    send_bot_message(
        token,
        chat_id,
        "Пришли ссылку на товар Kaspi, который нужно проанализировать.\n"
        "Соберу рейтинг, отзывы, цены и дам мнение специалиста.",
        reply_markup=build_keyboard(),
    )


def start_analysis_worker(token: str, chat_id: str, product_url: str):
    def _run():
        with ANALYSIS_LOCK:
            if chat_id in ANALYSIS_IN_PROGRESS:
                send_bot_message(token, chat_id, "Анализ уже выполняется, дождись завершения.", reply_markup=build_keyboard())
                return
            ANALYSIS_IN_PROGRESS.add(chat_id)
        try:
            send_bot_message(token, chat_id, "🔍 Собираю данные по карточке товара...", reply_markup=build_keyboard())
            resolved = resolve_kaspi_link(product_url)
            meta = analyze_product_card(resolved, city="Алматы")
            if not meta or not meta.get("product_name"):
                send_bot_message(token, chat_id, "Не удалось открыть карточку товара.", reply_markup=build_keyboard())
                return
            summary = format_product_analysis(meta)
            ai_text = gpt_product_opinion(meta)
            if ai_text:
                report = f"{summary}\n\n🤖 Отчет специалиста:\n{ai_text}"
            else:
                report = f"{summary}\n\n🤖 Отчет специалиста недоступен."
            send_chunked_message(token, chat_id, report)
        except Exception as exc:
            send_bot_message(token, chat_id, f"❌ Не удалось выполнить анализ: {exc}", reply_markup=build_keyboard())
        finally:
            with ANALYSIS_LOCK:
                ANALYSIS_IN_PROGRESS.discard(chat_id)

    th = threading.Thread(target=_run, daemon=True)
    th.start()


def _format_price_value(val):
    if val is None:
        return "—"
    try:
        return f"{int(val):,} ₸".replace(",", " ")
    except Exception:
        return str(val)


def format_price_report(meta: dict) -> str:
    lines = [
        f"💰 Анализ цен: {meta.get('product_name') or 'без названия'}",
        f"🔗 {meta.get('resolved_url')}",
    ]
    stats = meta.get("price_stats") or {}
    if stats.get("min_price") is not None:
        min_price = stats.get("min_price")
        max_price = stats.get("max_price")
        spread = stats.get("spread")
        step = stats.get("min_step")
        lines.append(f"📉 Мин: {_format_price_value(min_price)}")
        lines.append(f"📈 Макс: {_format_price_value(max_price)}")
        lines.append(f"↔️ Разброс: {_format_price_value(spread)}")
        lines.append(f"🪜 Минимальный шаг: {_format_price_value(step)}")
        lines.append(f"👥 Продавцов: {stats.get('seller_count') or 0}")
        if stats.get("min_sellers"):
            lines.append("🔽 Минимум у: " + ", ".join(stats["min_sellers"][:3]))
        if stats.get("max_sellers"):
            lines.append("🔼 Максимум у: " + ", ".join(stats["max_sellers"][:3]))
    sellers = meta.get("seller_prices") or []
    if sellers:
        lines.append("")
        lines.append("🏷️ Продавцы:")
        for entry in sellers[:10]:
            price_text = entry.get("price_text")
            if price_text:
                price_str = price_text
            else:
                price_str = _format_price_value(entry.get("price_kzt"))
            lines.append(f"• {entry.get('seller')}: {price_str}")
    return "\n".join(line for line in lines if line)


def start_price_analysis_worker(token: str, chat_id: str, product_url: str):
    def _run():
        with ANALYSIS_LOCK:
            if chat_id in PRICE_ANALYSIS_IN_PROGRESS:
                send_bot_message(token, chat_id, "Другой анализ цен ещё выполняется, дождись завершения.", reply_markup=build_keyboard())
                return
            PRICE_ANALYSIS_IN_PROGRESS.add(chat_id)
        try:
            send_bot_message(token, chat_id, "💸 Собираю цены по карточке...", reply_markup=build_keyboard())
            resolved = resolve_kaspi_link(product_url)
            meta = analyze_product_card(resolved, city="Алматы")
            if not meta or not meta.get("product_name"):
                send_bot_message(token, chat_id, "Не удалось открыть карточку товара.", reply_markup=build_keyboard())
                return
            report = format_price_report(meta)
            send_chunked_message(token, chat_id, report)
        except Exception as exc:
            send_bot_message(token, chat_id, f"❌ Не удалось сделать анализ цен: {exc}", reply_markup=build_keyboard())
        finally:
            with ANALYSIS_LOCK:
                PRICE_ANALYSIS_IN_PROGRESS.discard(chat_id)

    threading.Thread(target=_run, daemon=True).start()



def main():
    if not os.environ.get("TELEGRAM_BOT_TOKEN"):
        load_env_from_file()
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not token:
        raise RuntimeError("TELEGRAM_BOT_TOKEN is not set")

    base_url = f"https://api.telegram.org/bot{token}"
    offset = None

    print("🤖 Kaspi bot started. Commands: /add, /run, /run_extra, /stop, /status, /help, /agent")

    while True:
        try:
            resp = requests.get(
                f"{base_url}/getUpdates",
                params={"timeout": 25, "offset": offset},
                timeout=35,
            )
            data = resp.json()
            if not data.get("ok"):
                _log_event("telegram_poll_error", {"status": resp.status_code, "body": str(data)[:500]})
                time.sleep(2)
                continue

            for update in data.get("result", []):
                offset = update["update_id"] + 1
                # Callback query (кнопки)
                if update.get("callback_query"):
                    cb = update["callback_query"]
                    chat_id = str(cb["message"]["chat"]["id"])
                    data_cb = cb.get("data") or ""
                    _log_event(
                        "callback",
                        {
                            "chat_id": chat_id,
                            "data": data_cb,
                            "user": cb.get("from", {}).get("username") or cb.get("from", {}).get("id"),
                        },
                    )
                    # подтверждаем callback
                    try:
                        requests.post(f"{base_url}/answerCallbackQuery", data={"callback_query_id": cb["id"]})
                    except Exception:
                        pass
                    requester_id = str(cb.get("from", {}).get("id") or "")
                    handle_callback(token, chat_id, data_cb, requester_id=requester_id)
                    continue

                message = update.get("message") or update.get("edited_message")
                if not message:
                    continue
                chat_id = str(message["chat"]["id"])
                sender_id = str(message.get("from", {}).get("id") or "")
                text = message.get("text") or message.get("caption") or ""
                entities = message.get("entities") or message.get("caption_entities") or []

                # разрешаем всем

                log_chat_id(chat_id)
                _ensure_agent_default_mode(chat_id)

                low = text.lower().strip()

                pending_analysis = ANALYSIS_REQUESTS.get(chat_id)
                if pending_analysis and pending_analysis.get("status") == "await_link":
                    links = extract_kaspi_links(text, entities=entities)
                    if not links:
                        send_bot_message(token, chat_id, "Нужна ссылка на товар Kaspi. Отправь её сообщением.", reply_markup=build_keyboard())
                        continue
                    ANALYSIS_REQUESTS.pop(chat_id, None)
                    start_analysis_worker(token, chat_id, links[0])
                    continue

                pending_research = PENDING_RESEARCH.get(chat_id)
                if pending_research:
                    if low.startswith("/cancel"):
                        PENDING_RESEARCH.pop(chat_id, None)
                        send_bot_message(token, chat_id, "Отменил ожидание темы для ресёрча.", reply_markup=build_keyboard())
                        continue
                    if not low.startswith("/"):
                        query_text = text.strip()
                        if not query_text:
                            send_bot_message(token, chat_id, "Нужна тема для поиска. Напиши текст без команд или нажми Отмена.", reply_markup=build_keyboard())
                            continue
                        PENDING_RESEARCH.pop(chat_id, None)
                        request_research(token, chat_id, query_text, mode=pending_research.get("mode"))
                        continue
                if low.startswith("/start") or low.startswith("/help"):
                    send_bot_message(token, chat_id, extended_bot_help_text(), reply_markup=build_keyboard())
                    continue

                if low.startswith("/agent_status"):
                    send_bot_message(token, chat_id, agent_status_text(chat_id), reply_markup=build_keyboard())
                    continue

                if low.startswith("/agent_on"):
                    if not _agent_chat_allowed(chat_id):
                        send_bot_message(token, chat_id, "Agent-режим недоступен для этого чата.", reply_markup=build_keyboard())
                        continue
                    _set_agent_mode(chat_id, True)
                    send_bot_message(
                        token,
                        chat_id,
                        "✅ Agent авто-режим включён. Теперь обычный текст будет выполняться как задача.\n"
                        "Выключить: /agent_off",
                        reply_markup=build_keyboard(),
                    )
                    continue

                if low.startswith("/agent_off"):
                    _set_agent_mode(chat_id, False)
                    send_bot_message(token, chat_id, "✅ Agent авто-режим выключен.", reply_markup=build_keyboard())
                    continue

                if low.startswith("/agent"):
                    task = _command_payload(text, "/agent")
                    start_agent_task(token, chat_id, task)
                    continue

                if low.startswith("/status"):
                    send_bot_message(token, chat_id, status_text(), reply_markup=build_keyboard())
                    continue

                if low.startswith("/analyze"):
                    links = extract_kaspi_links(text, entities=entities)
                    if links:
                        start_analysis_worker(token, chat_id, links[0])
                    else:
                        request_analysis_link(token, chat_id)
                    continue

                if low.startswith("/research_more"):
                    question = text.replace("/research_more", "", 1).strip()
                    request_research_followup(token, chat_id, question)
                    continue

                if low.startswith("/research_deep"):
                    query = text.replace("/research_deep", "", 1).strip()
                    request_research(token, chat_id, query, mode="deep")
                    continue

                if low.startswith("/research"):
                    query = text.replace("/research", "", 1).strip()
                    request_research(token, chat_id, query, mode="fast")
                    continue

                if low.startswith("/auto_status"):
                    send_bot_message(token, chat_id, auto_alert_status_text(chat_id), reply_markup=build_keyboard())
                    continue

                if low.startswith("/stop"):
                    request_stop(token, chat_id)
                    continue

                if low.startswith("/ask"):
                    question = text.replace("/ask", "", 1).strip()
                    if not question:
                        send_bot_message(token, chat_id, "Напиши вопрос после /ask", reply_markup=build_keyboard())
                        continue
                    answer = gpt_answer(question, chat_id)
                    send_bot_message(token, chat_id, answer, reply_markup=build_keyboard())
                    continue

                if low.startswith("/add"):
                    link, name, expected = parse_add_payload(text)
                    if not link:
                        send_bot_message(token, chat_id, "Нужна ссылка. Пример: /add <link> | <имя> | <продавцы>", reply_markup=build_keyboard())
                        continue
                    append_product_row(chat_id, link, name, expected)
                    if expected:
                        send_bot_message(token, chat_id, "✅ Добавил в временный список", reply_markup=build_keyboard())
                    else:
                        # нет продавцов — предложим быстрый выбор
                        with PENDING_LOCK:
                            tok = uuid.uuid4().hex
                            PENDING_EXPECTED[tok] = {"link": link, "name": name, "selected": set(), "chat_id": chat_id}
                        send_bot_message(
                            token,
                            chat_id,
                            "✅ Добавил без продавцов. Выбери, кого отслеживать (можно несколько):",
                            reply_markup=build_expected_keyboard(tok, set()),
                        )
                    continue

                if low.startswith("/run_natasha_prices") or low.startswith("/run_competitors"):
                    start_natasha_price_scrape(token, chat_id, requester_id=sender_id)
                    continue

                if low.startswith("/run_keywords"):
                    start_keyword_search(token, chat_id)
                    continue

                if low.startswith("/run_extra"):
                    tok = uuid.uuid4().hex
                    with PENDING_CITY_LOCK:
                        PENDING_CITY[tok] = {"chat_id": chat_id, "mode": "run_extra"}
                    send_bot_message(token, chat_id, "Выбери город:", reply_markup=build_city_keyboard(tok))
                    continue

                if low.startswith("/run"):
                    start_main_scrape(token, chat_id)
                    continue

                if low.startswith("/broadcast"):
                    if chat_id not in BROADCAST_IDS:
                        send_bot_message(token, chat_id, "Нет прав на рассылку.", reply_markup=build_keyboard())
                        continue
                    msg_text = text.replace("/broadcast", "", 1).strip()
                    if not msg_text:
                        send_bot_message(token, chat_id, "Укажи текст: /broadcast <сообщение>", reply_markup=build_keyboard())
                        continue
                    # если нет списка, шлём только себе
                    targets = BROADCAST_IDS or {chat_id}
                    ok = 0
                    fail = 0
                    for tgt in targets:
                        try:
                            send_bot_message(token, tgt, msg_text, reply_markup=build_keyboard())
                            ok += 1
                        except Exception:
                            fail += 1
                    send_bot_message(token, chat_id, f"Рассылка завершена. Успешно: {ok}, ошибки: {fail}", reply_markup=build_keyboard())
                    continue

                if low.startswith("/auto_alerts_off"):
                    stop_auto_alerts(token, chat_id)
                    continue

                if low.startswith("/auto_alerts_on"):
                    parts = low.split()
                    minutes = 5
                    if len(parts) > 1 and parts[1].isdigit():
                        minutes = max(1, int(parts[1]))
                    interval_sec = minutes * 60
                    tok = uuid.uuid4().hex
                    with PENDING_CITY_LOCK:
                        PENDING_CITY[tok] = {"chat_id": chat_id, "mode": "auto_alerts", "interval_sec": interval_sec}
                    send_bot_message(token, chat_id, f"Выбери город для авто-алертов (каждые {minutes} мин):", reply_markup=build_city_keyboard(tok))
                    continue

                links = extract_kaspi_links(text, entities=entities)
                if links:
                    start_link_preview(token, chat_id, links)
                    continue

                # Если это вопрос про цену "цена на ..." — уточняем SKU/ссылку
                price_resp = price_query_response(text)
                if price_resp:
                    send_bot_message(token, chat_id, price_resp, reply_markup=build_keyboard())
                    continue

                if _agent_mode_enabled(chat_id):
                    start_agent_task(token, chat_id, text)
                    continue

                # Любое другое сообщение — отвечаем GPT (Jarvis)
                answer = gpt_answer(text, chat_id)
                send_bot_message(token, chat_id, answer, reply_markup=build_keyboard())

        except KeyboardInterrupt:
            print("👋 Bot stopped by user.")
            break
        except Exception as exc:
            print(f"⚠️ Bot loop error: {exc}")
            time.sleep(3)


if __name__ == "__main__":
    main()
