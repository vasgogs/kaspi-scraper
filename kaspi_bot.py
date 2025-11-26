import os
import time
import re
import json
import threading
import uuid
from pathlib import Path
import pandas as pd
import requests

from Scraper_Kaspi import run_kaspi_scrape, bot_help_text, resolve_kaspi_link


BASE_DIR = Path(__file__).resolve().parent
USER_DATA_DIR = BASE_DIR / "user_data"

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
PENDING_ADD: dict[str, dict] = {}
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-3.5-turbo")
MEMORY_LOCK = threading.Lock()
BROADCAST_IDS = {
    chat.strip()
    for chat in re.split(r"[;,]", os.environ.get("BROADCAST_CHAT_IDS", ""))
    if chat.strip()
}
ALLOWED_CHATS: set[str] | None = None


def user_extra_csv(chat_id: str) -> Path:
    USER_DATA_DIR.mkdir(parents=True, exist_ok=True)
    return USER_DATA_DIR / f"{chat_id}_extra.csv"


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


def send_bot_message(token: str, chat_id: str, text: str, reply_markup=None):
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    data = {"chat_id": chat_id, "text": text}
    if reply_markup:
        data["reply_markup"] = json.dumps(reply_markup)
    resp = requests.post(url, data=data)
    if resp.status_code != 200:
        raise RuntimeError(f"Telegram send failed: {resp.status_code} {resp.text}")


def send_bot_file(token: str, chat_id: str, file_path: Path, caption: str | None = None):
    url = f"https://api.telegram.org/bot{token}/sendDocument"
    with open(file_path, "rb") as fh:
        files = {"document": fh}
        data = {"chat_id": chat_id}
        if caption:
            data["caption"] = caption
        resp = requests.post(url, data=data, files=files, timeout=60)
    if resp.status_code != 200:
        raise RuntimeError(f"Telegram file send failed: {resp.status_code} {resp.text}")


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
                {"text": "🚨 Запуск (только алерты)", "callback_data": "run_alerts"},
            ],
            [
                {"text": "📂 Последний файл", "callback_data": "last_file"},
            ],
            [
                {"text": "📝 Временный список", "callback_data": "list_extra"},
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


def gpt_answer(question: str, chat_id: str) -> str:
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        return "API ключ OpenAI не настроен."

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

    payload = {
        "model": OPENAI_MODEL,
        "messages": messages,
        "temperature": 0.2,
        "max_tokens": 300,
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
    }
    try:
        resp = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload, timeout=15)
        data = resp.json()
        if resp.status_code == 200 and data.get("choices"):
            answer = data["choices"][0]["message"]["content"].strip()
            with MEMORY_LOCK:
                history = load_chat_memory(chat_id)
                history.append({"role": "user", "content": question})
                history.append({"role": "assistant", "content": answer})
                history = history[-20:]  # лимитируем историю
                save_chat_memory(chat_id, history)
            return answer
        err = f"GPT ошибка: {resp.status_code} {data}"
        return err
    except Exception as exc:
        return f"GPT запрос не удался: {exc}"


def start_scrape(token: str, chat_id: str, include_base: bool, alert_only: bool = False, cities=None):
    global SCRAPER_THREAD, SCRAPER_STOP_EVENT
    if SCRAPER_THREAD and SCRAPER_THREAD.is_alive():
        send_bot_message(token, chat_id, "Скрейп уже запущен", reply_markup=build_keyboard())
        return

    stop_event = threading.Event()
    SCRAPER_STOP_EVENT = stop_event

    def _run():
        try:
            send_bot_message(token, chat_id, "⏳ Запускаю скрейп...", reply_markup=build_keyboard())
            run_kaspi_scrape(
                include_base=include_base,
                chat_id=chat_id,
                extra_csv_paths=user_extra_csv(chat_id),
                stop_event=stop_event,
                progress_callback=progress_callback,
                cities=cities or None,
                alert_only=alert_only,
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

    def _loop():
        send_bot_message(token, chat_id, f"⏱️ Авто-алерты включены, каждые {interval_sec} сек, город {city or 'Алматы/Астана/Шымкент'}", reply_markup=build_keyboard())
        while not stop_event.is_set():
            try:
                run_kaspi_scrape(
                    include_base=False,
                    chat_id=chat_id,
                    extra_csv_paths=user_extra_csv(chat_id),
                    cities=[city] if city else None,
                    alert_only=True,
                )
            except Exception as exc:
                send_bot_message(token, chat_id, f"❌ Ошибка авто-алертов: {exc}", reply_markup=build_keyboard())
            # ожидание с возможностью прерывания
            for _ in range(interval_sec):
                if stop_event.is_set():
                    break
                time.sleep(1)
        send_bot_message(token, chat_id, "✋ Авто-алерты остановлены.", reply_markup=build_keyboard())

    th = threading.Thread(target=_loop, daemon=True)
    AUTO_THREADS[chat_id] = th
    th.start()


def stop_auto_alerts(token: str, chat_id: str):
    stop_event = AUTO_STOPS.get(chat_id)
    if stop_event:
        stop_event.set()
    else:
        send_bot_message(token, chat_id, "Авто-алерты не запущены.", reply_markup=build_keyboard())


def main():
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not token:
        raise RuntimeError("TELEGRAM_BOT_TOKEN is not set")

    base_url = f"https://api.telegram.org/bot{token}"
    offset = None

    print("🤖 Kaspi bot started. Commands: /add, /run, /run_extra, /stop, /status, /help")

    while True:
        try:
            resp = requests.get(
                f"{base_url}/getUpdates",
                params={"timeout": 25, "offset": offset},
                timeout=35,
            )
            data = resp.json()
            if not data.get("ok"):
                time.sleep(2)
                continue

            for update in data.get("result", []):
                offset = update["update_id"] + 1
                # Callback query (кнопки)
                if update.get("callback_query"):
                    cb = update["callback_query"]
                    chat_id = str(cb["message"]["chat"]["id"])
                    data_cb = cb.get("data") or ""
                    # подтверждаем callback
                    requests.post(f"{base_url}/answerCallbackQuery", data={"callback_query_id": cb["id"]})
                    # разрешаем всем
                    if data_cb == "run_full":
                        start_scrape(token, chat_id, include_base=True)
                    elif data_cb == "run_extra":
                        tok = uuid.uuid4().hex
                        with PENDING_CITY_LOCK:
                            PENDING_CITY[tok] = {"chat_id": chat_id, "mode": "run_extra"}
                        send_bot_message(token, chat_id, "Выбери город:", reply_markup=build_city_keyboard(tok))
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
                        send_bot_message(token, chat_id, "Используй /ask <вопрос>, пример: /ask как добавить товар и выбрать продавцов?", reply_markup=build_keyboard())
                    elif data_cb.startswith("setexp:"):
                        parts = data_cb.split(":")
                        if len(parts) == 3:
                            _prefix, tok, choice = parts
                            with PENDING_LOCK:
                                info = PENDING_EXPECTED.get(tok)
                                if info is None:
                                    send_bot_message(token, chat_id, "Не нашёл, попробуй ещё раз /add", reply_markup=build_keyboard())
                                    continue
                                if str(info.get("chat_id")) != str(chat_id):
                                    send_bot_message(token, chat_id, "Эта выборка продавцов принадлежит другому чату.", reply_markup=build_keyboard())
                                    continue
                                selected: set[str] = info.get("selected", set())
                                if choice == "save":
                                    # сохраняем выбор
                                    selected_names: list[str] = []
                                    for key in selected:
                                        _, names = EXPECTED_OPTIONS[key]
                                        selected_names.extend(names)
                                    # убираем дубли, сохраняем порядок
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
                                    continue
                                if choice == "skip":
                                    PENDING_EXPECTED.pop(tok, None)
                                    send_bot_message(token, chat_id, "Оставил без продавцов. Можно запускать скрейп.", reply_markup=build_keyboard())
                                    continue
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
                                        continue
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
                                continue
                            if city_choice == "cancel":
                                send_bot_message(token, chat_id, "Отменено.", reply_markup=build_keyboard())
                                continue
                            if str(info.get("chat_id")) != str(chat_id):
                                send_bot_message(token, chat_id, "Этот выбор города не для вашего чата.", reply_markup=build_keyboard())
                                continue
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
                    continue

                message = update.get("message") or update.get("edited_message")
                if not message:
                    continue
                chat_id = str(message["chat"]["id"])
                text = message.get("text") or message.get("caption") or ""
                entities = message.get("entities") or message.get("caption_entities") or []

                # разрешаем всем

                log_chat_id(chat_id)

                low = text.lower().strip()
                if low.startswith("/start") or low.startswith("/help"):
                    send_bot_message(token, chat_id, bot_help_text(), reply_markup=build_keyboard())
                    continue

                if low.startswith("/status"):
                    send_bot_message(token, chat_id, status_text(), reply_markup=build_keyboard())
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

                if low.startswith("/run_extra"):
                    tok = uuid.uuid4().hex
                    with PENDING_CITY_LOCK:
                        PENDING_CITY[tok] = {"chat_id": chat_id, "mode": "run_extra"}
                    send_bot_message(token, chat_id, "Выбери город:", reply_markup=build_city_keyboard(tok))
                    continue

                if low.startswith("/run"):
                    start_scrape(token, chat_id, include_base=True)
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
                    PENDING_ADD[chat_id] = {"links": links}
                    send_bot_message(
                        token,
                        chat_id,
                        f"Нашёл {len(links)} ссылок Kaspi. Добавить их в твой временный список?",
                        reply_markup=build_add_links_keyboard(),
                    )
                    continue

                # Если это вопрос про цену "цена на ..." — уточняем SKU/ссылку
                price_resp = price_query_response(text)
                if price_resp:
                    send_bot_message(token, chat_id, price_resp, reply_markup=build_keyboard())
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
