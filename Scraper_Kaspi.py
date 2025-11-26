from playwright.sync_api import sync_playwright, TimeoutError
import os
import pandas as pd
import re
import tempfile
import requests
from collections.abc import Sequence
from datetime import datetime
import threading
import json
from pathlib import Path
from urllib.parse import urlparse
from concurrent.futures import ProcessPoolExecutor, as_completed

DEFAULT_SHAREPOINT_SITE_URL = "https://stadaarz-my.sharepoint.com/personal/vasily_gogolev_stada_kz"
DEFAULT_SHAREPOINT_FILE_URL = "/personal/vasily_gogolev_stada_kz/Documents/Book 1.xlsx"
DEFAULT_SHAREPOINT_USERNAME = "vasily.gogolev@stada.kz"
DEFAULT_SHAREPOINT_PASSWORD = os.environ.get("SHAREPOINT_PASSWORD", "")
DEFAULT_TENANT_ID = os.environ.get("AZURE_TENANT_ID", "organizations")
DEFAULT_GRAPH_UPLOAD_PATH = "Documents/Book 1.xlsx"
TELEGRAM_SAFE_CHUNK = 3000  # запас до лимита 4096 и учёт смайлов/UTF-8

# Быстрые таймауты/паузы, можно повысить при блокировках.
FAST_NAV_TIMEOUT = 12000  # первая попытка загрузки, мс
SLOW_NAV_TIMEOUT = 25000  # вторая попытка загрузки, мс
POST_LOAD_DELAY_FAST = 800  # задержка после goto, мс
POST_LOAD_DELAY_SLOW = 2000  # задержка после goto (повтор), мс
POST_PAGINATION_DELAY = 500  # задержка после клика Следующая, мс
SELLER_WAIT_TIMEOUT = 12000  # ожидание блока продавцов, мс
STATE_DIR = Path(__file__).resolve().parent / "state"
SNAPSHOT_TTL_HOURS = 12


try:
    from office365.runtime.auth.user_credential import UserCredential
    from office365.runtime.auth.client_credential import ClientCredential
    from office365.sharepoint.client_context import ClientContext
    from office365.sharepoint.files.file import File
    import msal
except ImportError:
    ClientContext = None
    UserCredential = None
    ClientCredential = None
    File = None
    msal = None


def _split_telegram_text(text: str, max_len: int = TELEGRAM_SAFE_CHUNK) -> list[str]:
    """Разбиваем длинный текст на части ниже лимита Telegram."""
    if not text:
        return []
    max_len = max(1, min(max_len, 4096))
    chunks: list[str] = []
    current = ""
    for line in text.splitlines():
        line = line.rstrip()
        if not current:
            if len(line) <= max_len:
                current = line
            else:
                for i in range(0, len(line), max_len):
                    chunks.append(line[i:i + max_len])
                current = ""
            continue
        projected = len(current) + 1 + len(line)
        if projected <= max_len:
            current = f"{current}\n{line}"
        else:
            chunks.append(current)
            if len(line) <= max_len:
                current = line
            else:
                for i in range(0, len(line), max_len):
                    chunks.append(line[i:i + max_len])
                current = ""
    if current:
        chunks.append(current)
    return [chunk for chunk in chunks if chunk]


def extract_price(text: str):
    digits = re.sub(r"[^\d]", "", text)
    return int(digits) if digits else None


def extract_product_code(url: str) -> str:
    """Получаем код товара из URL каспи (/p/<name>-123456/)."""
    match = re.search(r"/p/[^/]*?-(\d+)(?:/|$)", url)
    return match.group(1) if match else ""


def normalize_seller_name(name: str) -> str:
    """Приведение имени продавца к единому виду для сравнения."""
    return re.sub(r"\s+", " ", (name or "").strip()).lower()


def resolve_kaspi_link(url: str, timeout: int = 8) -> str:
    """Разворачиваем короткие l.kaspi.kz ссылки в полноценный product URL."""
    if not url:
        return url
    try:
        parsed = urlparse(url if "://" in url else f"https://{url}")
        if parsed.netloc.lower() == "l.kaspi.kz" or parsed.path.startswith("/shp/"):
            target = url if "://" in url else f"https://{url}"
            resp = requests.get(target, allow_redirects=True, timeout=timeout)
            if resp.url:
                return resp.url
        elif not parsed.scheme:
            # добавляем https для прямых kaspi.kz ссылок без схемы
            return f"https://{url}"
    except Exception:
        pass
    return url


def parse_expected_sellers(raw_value) -> list[str]:
    """Парсим колонку expected_sellers из CSV (разделители ; или ,)."""
    if raw_value is None:
        return []
    try:
        if pd.isna(raw_value):
            return []
    except Exception:
        pass
    parts = re.split(r"[;,]", str(raw_value))
    return [part.strip() for part in parts if part and part.strip()]


def find_missing_expected_sellers(records: list[dict], job_meta_list: list[dict]) -> list[dict]:
    """Сравниваем найденных продавцов с ожидаемыми и собираем алерты."""
    alerts = []
    for meta in job_meta_list:
        expected = meta.get("expected_sellers") or []
        if not expected:
            continue
        expected_norm = {normalize_seller_name(name) for name in expected if name}
        if not expected_norm:
            continue
        matched_records = [
            rec
            for rec in records
            if rec.get("region") == meta.get("city")
            and (
                (meta.get("product_code") and rec.get("product_code") == meta.get("product_code"))
                or rec.get("product_url") == meta.get("product_url")
            )
        ]
        found_norm = {
            normalize_seller_name(rec.get("seller", ""))
            for rec in matched_records
            if rec.get("seller")
        }
        missing = [name for name in expected if normalize_seller_name(name) not in found_norm]
        if missing:
            found_raw = sorted({rec.get("seller", "") for rec in matched_records if rec.get("seller")})
            product_label = ""
            if matched_records:
                product_label = (
                    matched_records[0].get("product")
                    or matched_records[0].get("input_product")
                    or ""
                )
            if not product_label:
                product_label = meta.get("product_name") or meta.get("product_url")
            alerts.append({
                "city": meta.get("city"),
                "product": product_label,
                "product_code": meta.get("product_code") or "",
                "product_url": meta.get("product_url"),
                "missing": missing,
                "found": found_raw,
            })
    return alerts


def _combine_excel(existing_df: pd.DataFrame, new_df: pd.DataFrame) -> pd.DataFrame:
    has_existing_structure = len(existing_df.columns) > 0
    frames = [existing_df, new_df] if has_existing_structure else [new_df]
    combined_df = pd.concat(frames, ignore_index=True, sort=False)

    column_order = existing_df.columns.tolist() if has_existing_structure else new_df.columns.tolist()
    for col in new_df.columns:
        if col not in column_order:
            column_order.append(col)
    return combined_df.reindex(columns=column_order)


def _snapshot_key(record: dict) -> str:
    product_code = record.get("product_code") or extract_product_code(record.get("product_url", ""))
    return f"{record.get('region','')}|{product_code or record.get('product_url','')}|{normalize_seller_name(record.get('seller',''))}"


def load_snapshot(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        mtime = path.stat().st_mtime
        age_hours = (datetime.now().timestamp() - mtime) / 3600
        if age_hours > SNAPSHOT_TTL_HOURS:
            return {}
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def save_snapshot(path: Path, snapshot: dict):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(snapshot, fh, ensure_ascii=False, indent=2)


def _graph_upload(df: pd.DataFrame) -> bool:
    """Upload via Microsoft Graph if credentials are provided."""
    if not msal:
        return False

    client_id = os.environ.get("GRAPH_CLIENT_ID") or os.environ.get("SHAREPOINT_CLIENT_ID")
    client_secret = os.environ.get("GRAPH_CLIENT_SECRET") or os.environ.get("SHAREPOINT_CLIENT_SECRET")
    tenant_id = os.environ.get("GRAPH_TENANT_ID") or os.environ.get("SHAREPOINT_TENANT_ID") or DEFAULT_TENANT_ID
    upload_path = os.environ.get("GRAPH_UPLOAD_PATH") or DEFAULT_GRAPH_UPLOAD_PATH
    target_user = (
        os.environ.get("GRAPH_USER_ID")
        or os.environ.get("GRAPH_USER_UPN")
        or os.environ.get("SHAREPOINT_USERNAME")
        or DEFAULT_SHAREPOINT_USERNAME
    )

    # Если нет client credentials, пробуем resource owner password (устарело, может блокироваться при MFA)
    username = os.environ.get("GRAPH_USERNAME") or os.environ.get("SHAREPOINT_USERNAME")
    password = os.environ.get("GRAPH_PASSWORD") or os.environ.get("SHAREPOINT_PASSWORD")

    scopes = ["https://graph.microsoft.com/.default"]
    app = None
    token = None

    if client_id and client_secret:
        try:
            authority = f"https://login.microsoftonline.com/{tenant_id}"
            app = msal.ConfidentialClientApplication(
                client_id=client_id,
                authority=authority,
                client_credential=client_secret,
            )
            token = app.acquire_token_silent(scopes, account=None)
            if not token:
                token = app.acquire_token_for_client(scopes=scopes)
            if "access_token" not in token:
                raise RuntimeError(token.get("error_description") or "No access token")
            print("🔐 Using Graph client_credentials flow")
        except Exception as exc:
            print(f"⚠️ Graph client_credentials failed: {exc}")

    if token is None and client_id and username and password:
        try:
            authority = f"https://login.microsoftonline.com/{tenant_id}"
            app = msal.PublicClientApplication(client_id=client_id, authority=authority)
            token = app.acquire_token_by_username_password(
                username=username,
                password=password,
                scopes=["Files.ReadWrite.All", "offline_access"],
            )
            if "access_token" not in token:
                raise RuntimeError(token.get("error_description") or "No access token")
            print("🔐 Using Graph username/password flow")
        except Exception as exc:
            print(f"⚠️ Graph username/password failed: {exc}")

    if not token or "access_token" not in token:
        return False

    headers = {"Authorization": f"Bearer {token['access_token']}"}
    download_url = f"https://graph.microsoft.com/v1.0/users/{target_user}/drive/root:/{upload_path}:/content"

    existing_df = pd.DataFrame()
    try:
        dl_resp = requests.get(download_url, headers=headers)
        if dl_resp.status_code == 200:
            with tempfile.NamedTemporaryFile(suffix=".xlsx") as tmp_dl:
                tmp_dl.write(dl_resp.content)
                tmp_dl.flush()
                existing_df = pd.read_excel(tmp_dl.name)
        elif dl_resp.status_code != 404:
            print(f"⚠️ Graph download failed: HTTP {dl_resp.status_code} {dl_resp.text}")
    except Exception as exc:
        print(f"⚠️ Graph download error: {exc}")

    combined_df = _combine_excel(existing_df, df)

    try:
        with tempfile.NamedTemporaryFile(suffix=".xlsx") as tmp_out:
            combined_df.to_excel(tmp_out.name, index=False)
            tmp_out.flush()
            with open(tmp_out.name, "rb") as fh:
                put_resp = requests.put(download_url, headers=headers, data=fh.read())
        if 200 <= put_resp.status_code < 300:
            print("✅ Uploaded rows to SharePoint via Graph")
            return True
        print(f"⚠️ Graph upload failed: HTTP {put_resp.status_code} {put_resp.text}")
    except Exception as exc:
        print(f"⚠️ Graph upload error: {exc}")

    return False


def send_telegram_message(text: str, chat_id: str | None = None):
    """Send a plain-text message to Telegram. Fails hard if not configured."""
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = chat_id or os.environ.get("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        raise RuntimeError("TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID must be set for Telegram alerts")
    url = f"https://api.telegram.org/bot{token}/sendMessage"
    chunks = _split_telegram_text(text, TELEGRAM_SAFE_CHUNK)
    if not chunks:
        return

    def _send_chunk(chunk: str, allow_resplit: bool = True):
        resp = requests.post(url, data={"chat_id": chat_id, "text": chunk})
        if resp.status_code == 200:
            return
        if resp.status_code == 413 and allow_resplit and len(chunk) > 1200:
            smaller_chunks = _split_telegram_text(chunk, max_len=1200)
            for sub in smaller_chunks:
                _send_chunk(sub, allow_resplit=False)
            return
        raise RuntimeError(f"Telegram send failed: {resp.status_code} {resp.text}")

    try:
        for idx, chunk in enumerate(chunks, start=1):
            _send_chunk(chunk, allow_resplit=True)
    except Exception as exc:
        raise RuntimeError(f"Telegram send error: {exc}") from exc


def send_telegram_file(file_path: Path, caption: str | None = None, chat_id: str | None = None):
    """Send a file to Telegram chat. Fails hard if not configured."""
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = chat_id or os.environ.get("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        raise RuntimeError("TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID must be set for Telegram alerts")
    url = f"https://api.telegram.org/bot{token}/sendDocument"
    try:
        with open(file_path, "rb") as fh:
            files = {"document": fh}
            data = {"chat_id": chat_id}
            if caption:
                data["caption"] = caption
            resp = requests.post(url, data=data, files=files, timeout=60)
        if resp.status_code != 200:
            raise RuntimeError(f"Telegram file send failed: {resp.status_code} {resp.text}")
    except Exception as exc:
        raise RuntimeError(f"Telegram file send error: {exc}") from exc


def append_results_to_sharepoint(df: pd.DataFrame):
    """Append scraped rows to the configured SharePoint workbook."""
    site_url = os.environ.get("SHAREPOINT_SITE_URL") or DEFAULT_SHAREPOINT_SITE_URL
    file_url = os.environ.get("SHAREPOINT_FILE_RELATIVE_URL") or DEFAULT_SHAREPOINT_FILE_URL
    username = os.environ.get("SHAREPOINT_USERNAME") or DEFAULT_SHAREPOINT_USERNAME
    password = os.environ.get("SHAREPOINT_PASSWORD") or DEFAULT_SHAREPOINT_PASSWORD
    client_id = os.environ.get("SHAREPOINT_CLIENT_ID")
    client_secret = os.environ.get("SHAREPOINT_CLIENT_SECRET")
    tenant_id = os.environ.get("SHAREPOINT_TENANT_ID") or DEFAULT_TENANT_ID

    # 0) Graph API (client credentials или username/password) — самая надёжная автоматизация при наличии app registration
    if _graph_upload(df):
        return

    if not ClientContext or not File:
        print("ℹ️ Install 'office365-rest-python-client' to enable SharePoint uploads")
        return

    ctx = None

    # 1) client credentials (если заданы id/secret)
    if client_id and client_secret and ClientCredential:
        try:
            creds = ClientCredential(client_id, client_secret)
            ctx = ClientContext(site_url).with_credentials(creds)
            print("🔐 Using SharePoint client credentials flow")
        except Exception as exc:
            print(f"⚠️ Client credentials auth failed: {exc}")

    # 2) user/password (fallback)
    if ctx is None and username and password and UserCredential:
        try:
            creds = UserCredential(username, password)
            ctx = ClientContext(site_url).with_credentials(creds)
            print("🔐 Using SharePoint user/password auth")
        except Exception as exc:
            print(f"⚠️ User/password auth failed: {exc}")

    if ctx is None:
        print("ℹ️ No valid SharePoint auth configured; skipping cloud upload")
        return

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir) / "kaspi_upload.xlsx"
        existing_df = pd.DataFrame()
        try:
            response = File.open_binary(ctx, file_url)
            with open(tmp_path, "wb") as download_file:
                download_file.write(response.content)
            existing_df = pd.read_excel(tmp_path)
        except Exception as exc:
            print(f"ℹ️ Could not download existing workbook ({exc}); creating a new file")

        has_existing_structure = len(existing_df.columns) > 0
        frames = [existing_df, df] if has_existing_structure else [df]
        combined_df = pd.concat(frames, ignore_index=True, sort=False)

        column_order = existing_df.columns.tolist() if has_existing_structure else df.columns.tolist()
        for col in df.columns:
            if col not in column_order:
                column_order.append(col)
        combined_df = combined_df.reindex(columns=column_order)

        combined_df.to_excel(tmp_path, index=False)
        try:
            with open(tmp_path, "rb") as upload_file:
                File.save_binary(ctx, file_url, upload_file.read())
            print("✅ Uploaded rows to SharePoint workbook")
        except Exception as exc:
            print(f"⚠️ Failed to upload to SharePoint: {exc}")
            return


def scrape_single_product(page, url: str, input_name: str | None = None, city: str = "Алматы", scraped_date: str | None = None):
    product_name = (input_name or "").strip()
    # первая попытка: быстрый заход
    try:
        page.set_default_timeout(FAST_NAV_TIMEOUT)  # короче таймауты для скорости
        page.goto(url, wait_until="domcontentloaded", timeout=FAST_NAV_TIMEOUT)
        page.wait_for_timeout(POST_LOAD_DELAY_FAST)
    except TimeoutError:
        print(f"⏱️ First attempt timed out for {input_name or url}, retrying with relaxed settings...")
        try:
            page.set_default_timeout(SLOW_NAV_TIMEOUT)
            page.goto(url, timeout=SLOW_NAV_TIMEOUT)  # без wait_until
            page.wait_for_timeout(POST_LOAD_DELAY_SLOW)
        except TimeoutError:
            print(f"❌ Final timeout for {input_name or url}: Page.goto failed twice, skipping.")
            return None

    page_product_name_selectors = [
        "h1[data-product-name]",
        "h1[itemprop='name']",
        "[data-test='productName']",
        "h1",
    ]
    for selector in page_product_name_selectors:
        locator = page.locator(selector).first
        if locator.count():
            try:
                product_name = product_name or locator.inner_text().strip()
            except Exception:
                pass
            if product_name:
                break
    if not product_name:
        try:
            product_name = page.title().strip()
        except Exception:
            product_name = input_name or ""

    product_code = extract_product_code(url)
    if not product_code:
        meta_sku = page.locator("meta[itemprop='sku'], meta[name='sku']").first
        if meta_sku.count():
            product_code = meta_sku.get_attribute("content") or ""

    try:
        page.wait_for_selector("text=Выберите ваш город", timeout=5000)
        page.click(f"a:has-text('{city}')")
        page.wait_for_timeout(1000)
    except Exception:
        pass  # city modal did not appear

    sellers_tab = None
    tab_selectors = [
        "li[data-tab='offers']",
        "[data-test='tabOffers']",
        "text=Продавцы",
    ]
    for tab_selector in tab_selectors:
        try:
            page.wait_for_selector(tab_selector, timeout=SELLER_WAIT_TIMEOUT)
            locator = page.locator(tab_selector).first
            if locator.count():
                sellers_tab = locator
                break
        except TimeoutError:
            continue

    if not sellers_tab or not sellers_tab.count():
        print(f"⚠️ Sellers tab not found for {input_name or url}")
        return None

    try:
        sellers_tab.scroll_into_view_if_needed(timeout=3000)
        sellers_tab.click(force=True, timeout=5000)  # кликаем даже если элемент вне вьюпорта
    except Exception:
        try:
            sellers_tab.click(timeout=5000)
        except Exception as exc:
            print(f"⚠️ Could not click sellers tab for {input_name or url}: {exc}")
            return None
    seller_table_rows = "table.sellers-table__self tbody tr"
    seller_card = "div[data-test='sellerItem']"
    seller_wait_selector = f"{seller_table_rows}, {seller_card}"
    try:
        page.wait_for_selector(seller_wait_selector, timeout=SELLER_WAIT_TIMEOUT)
    except TimeoutError:
        print(f"⚠️ No sellers block found for {input_name or url}")
        return None

    results = []
    while True:
        if page.locator(seller_table_rows).count():
            sellers = page.locator(seller_table_rows)
            count = sellers.count()
            for i in range(count):
                row = sellers.nth(i)
                seller_link = row.locator("td").first.locator("a")
                if not seller_link.count():
                    continue
                name = seller_link.first.inner_text().strip()
                price_cell = row.locator("div.sellers-table__price-cell-text").first
                if not price_cell.count():
                    continue
                price_text = price_cell.inner_text().replace("\xa0", " ").strip()

                installment_cell = row.locator("div.sellers-table__price-cell-text._installments-price").first
                installment_text = (
                    installment_cell.inner_text().replace("\xa0", " ").strip()
                    if installment_cell.count()
                    else ""
                )

                delivery_options = row.locator("div.sellers-table__delivery-cell-option")
                delivery_types = []
                delivery_details = []
                option_count = delivery_options.count()
                for j in range(option_count):
                    option_text = delivery_options.nth(j).inner_text().replace("\xa0", " ").strip()
                    if not option_text:
                        continue
                    parts = [part.strip() for part in option_text.split(",")]
                    if parts:
                        delivery_types.append(parts[0])
                        if len(parts) > 1:
                            detail = ", ".join(parts[1:]).strip()
                            if detail:
                                delivery_details.append(f"{parts[0]}: {detail}")

                numeric_price = extract_price(price_text)
                numeric_installment = extract_price(installment_text) if installment_text else None

                results.append({
                    "input_product": input_name or "",
                    "product": product_name,
                    "product_url": url,
                    "seller": name,
                    "product_code": product_code,
                    "price_text": price_text,
                    "price_kzt": numeric_price,
                    "installment_text": installment_text,
                    "installment_kzt": numeric_installment,
                    "delivery_types": "; ".join(dict.fromkeys(delivery_types)),
                    "delivery_details": " | ".join(delivery_details),
                    "scraped_date": scraped_date or datetime.now().strftime("%d.%m.%Y"),
                    "region": city,
                })
        elif page.locator(seller_card).count():
            sellers = page.locator(seller_card)
            count = sellers.count()
            for i in range(count):
                seller = sellers.nth(i)
                name = ""
                price_text = ""
                installment_text = ""
                try:
                    name = seller.locator("a").first.inner_text().strip()
                except Exception:
                    pass
                price_locators = [
                    "[data-test='item-price']",
                    "span:has-text('₸')",
                    "div:has-text('₸')",
                ]
                for selector in price_locators:
                    loc = seller.locator(selector).first
                    if loc.count():
                        price_text = loc.inner_text().replace("\xa0", " ").strip()
                        break
                installment_loc = seller.locator("[data-test='item-installment-price']").first
                if installment_loc.count():
                    installment_text = installment_loc.inner_text().replace("\xa0", " ").strip()

                delivery_blocks = seller.locator("[data-test='deliveryOption'], .delivery-option")
                delivery_types = []
                delivery_details = []
                option_count = delivery_blocks.count()
                for j in range(option_count):
                    option_text = delivery_blocks.nth(j).inner_text().replace("\xa0", " ").strip()
                    if not option_text:
                        continue
                    parts = [part.strip() for part in option_text.split(",")]
                    if parts:
                        delivery_types.append(parts[0])
                        if len(parts) > 1:
                            detail = ", ".join(parts[1:]).strip()
                            if detail:
                                delivery_details.append(f"{parts[0]}: {detail}")

                numeric_price = extract_price(price_text)
                numeric_installment = extract_price(installment_text) if installment_text else None

                results.append({
                    "input_product": input_name or "",
                    "product": product_name,
                    "product_url": url,
                    "seller": name,
                    "product_code": product_code,
                    "price_text": price_text,
                    "price_kzt": numeric_price,
                    "installment_text": installment_text,
                    "installment_kzt": numeric_installment,
                    "delivery_types": "; ".join(dict.fromkeys(delivery_types)),
                    "delivery_details": " | ".join(delivery_details),
                    "scraped_date": scraped_date or datetime.now().strftime("%d.%m.%Y"),
                    "region": city,
                })
        else:
            print(f"⚠️ Seller list disappeared for {input_name or url}")
            break

        next_btn = page.locator(".pagination__el", has_text="Следующая")
        if not next_btn.count():
            next_btn = page.locator("button:has-text('Следующая'), a:has-text('Следующая')")
        if not next_btn.count() or "_disabled" in (next_btn.first.get_attribute("class") or ""):
            break
        next_btn.first.click()
        page.wait_for_timeout(POST_PAGINATION_DELAY)
        try:
            page.wait_for_selector(seller_wait_selector, timeout=SELLER_WAIT_TIMEOUT)
        except TimeoutError:
            break

    return results


def _scrape_product_job(args):
    """Запуск одного товара в отдельном процессе/окне для параллельной работы."""
    url, provided_name, city, scrape_date, _expected_sellers = args
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(viewport={"width": 1600, "height": 900})
            page = context.new_page()
            res = scrape_single_product(page, url, provided_name, city=city, scraped_date=scrape_date)
            context.close()
            browser.close()
            return res or []
    except Exception as exc:
        print(f"❌ Failed in worker for {provided_name or url} ({city}): {exc}")
        return []


def scrape_products_from_csv(
    csv_paths: Path | Sequence[Path],
    cities: Sequence[str] | None = None,
    stop_event: threading.Event | None = None,
    progress_callback=None,
    chat_id: str | None = None,
    alert_only: bool = False,
    include_base: bool = True,
):
    if isinstance(csv_paths, (str, Path)):
        csv_paths = [Path(csv_paths)]
    csv_paths = [Path(p) for p in csv_paths]

    frames = []
    for path in csv_paths:
        if not path.exists():
            print(f"ℹ️ CSV not found, skipping: {path}")
            continue
        df = pd.read_csv(path)
        if "product_link" not in df.columns:
            print(f"ℹ️ CSV {path} skipped: no 'product_link' column")
            continue
        frames.append(df)

    if not frames:
        raise FileNotFoundError(f"No valid CSV files found in: {', '.join(str(p) for p in csv_paths)}")

    products_df = pd.concat(frames, ignore_index=True, sort=False)
    products_df = products_df.drop_duplicates(subset=["product_link"]).reset_index(drop=True)

    all_records = []
    job_meta_map = {}
    job_meta_list = []
    cities = list(cities) if cities else ["Алматы", "Астана", "Шымкент"]

    scrape_date = datetime.now().strftime("%d.%m.%Y")
    max_workers = int(os.environ.get("SCRAPER_WORKERS", "8"))

    for city in cities:
        if stop_event and stop_event.is_set():
            break
        jobs = []
        for _, row in products_df.iterrows():
            product_url = resolve_kaspi_link(str(row["product_link"]).strip())
            if not product_url:
                continue
            provided_name = str(row.get("product_name", "")).strip()
            expected_sellers = tuple(parse_expected_sellers(row.get("expected_sellers", "")))
            product_code = extract_product_code(product_url)
            job = (product_url, provided_name, city, scrape_date, expected_sellers)
            jobs.append(job)
            meta = {
                "product_url": product_url,
                "product_name": provided_name,
                "city": city,
                "expected_sellers": list(expected_sellers),
                "product_code": product_code,
            }
            job_meta_map[job] = meta
            job_meta_list.append(meta)
        if not jobs:
            continue

        if progress_callback:
            progress_callback("city_start", city=city, total=len(jobs))

        print(f"🚀 Starting parallel scrape for {len(jobs)} items in {city} with {max_workers} workers")
        with ProcessPoolExecutor(max_workers=max_workers) as executor:
            future_map = {executor.submit(_scrape_product_job, job): job for job in jobs}
            for future in as_completed(future_map):
                if stop_event and stop_event.is_set():
                    executor.shutdown(cancel_futures=True)
                    break
                job = future_map[future]
                product_url, provided_name, city_label, _, _ = job
                meta = job_meta_map.get(job, {})
                label = provided_name or product_url
                try:
                    product_records = future.result()
                except Exception as exc:
                    print(f"❌ Worker crashed for {label} ({city_label}): {exc}")
                    continue
                if not product_records:
                    print(f"⚠️ No seller data for {label} ({city_label})")
                else:
                    for rec in product_records:
                        if not rec.get("input_product"):
                            rec["input_product"] = meta.get("product_name") or ""
                        if not rec.get("product_code"):
                            rec["product_code"] = meta.get("product_code") or ""
                        if not rec.get("product_url"):
                            rec["product_url"] = meta.get("product_url") or ""
                        rec["expected_sellers"] = "; ".join(meta.get("expected_sellers", []))
                    all_records.extend(product_records)
                if progress_callback:
                    progress_callback("progress", city=city_label, done=len(all_records))

        if progress_callback:
            progress_callback("city_done", city=city, total=len(jobs))

        if stop_event and stop_event.is_set():
            break

    if stop_event and stop_event.is_set():
        if progress_callback:
            progress_callback("stopped", city=None, done=len(all_records))
        return
    if not all_records:
        print("⚠️ No data collected; skipping Excel export")
        if progress_callback:
            progress_callback("no_data", city=None, done=0)
        return

    missing_alerts = find_missing_expected_sellers(all_records, job_meta_list)

    df = pd.DataFrame(all_records, columns=[
        "input_product",
        "product",
        "product_url",
        "seller",
        "expected_sellers",
        "price_text",
        "price_kzt",
        "installment_text",
        "installment_kzt",
        "delivery_types",
        "delivery_details",
        "scraped_date",
        "region",
        "product_code",
    ])

    # Preserve integers while allowing missing values
    df["price_kzt"] = df["price_kzt"].astype("Int64")
    df["installment_kzt"] = df["installment_kzt"].astype("Int64")

    # Сравнение со снимком
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    snapshot_path = STATE_DIR / "last_snapshot.json"
    prev_snapshot = load_snapshot(snapshot_path)
    first_snapshot = len(prev_snapshot) == 0
    curr_snapshot = {}
    changes = []

    product_keys = set()
    for rec in all_records:
        key = _snapshot_key(rec)
        prod_key = rec.get("product_code") or extract_product_code(rec.get("product_url", "")) or rec.get("product_url")
        if prod_key:
            product_keys.add(f"{rec.get('region','')}|{prod_key}")
        curr_snapshot[key] = {
            "price_kzt": rec.get("price_kzt"),
            "installment_kzt": rec.get("installment_kzt"),
            "seller": rec.get("seller"),
            "product": rec.get("product"),
            "product_code": rec.get("product_code"),
            "region": rec.get("region"),
            "product_url": rec.get("product_url"),
        }

    if not first_snapshot:
        for key, cur in curr_snapshot.items():
            prev = prev_snapshot.get(key)
            if prev:
                if prev.get("price_kzt") != cur.get("price_kzt") or prev.get("installment_kzt") != cur.get("installment_kzt"):
                    changes.append(f"💸 {cur.get('product')} — {cur.get('region')} — {cur.get('seller')}: цена {prev.get('price_kzt')} → {cur.get('price_kzt')}")
            else:
                changes.append(f"🆕 {cur.get('product')} — {cur.get('region')} — новый продавец {cur.get('seller')} цена {cur.get('price_kzt')}")

        for key, prev in prev_snapshot.items():
            if key not in curr_snapshot:
                prev_prod_key = f"{prev.get('region','')}|{prev.get('product_code') or extract_product_code(prev.get('product_url','')) or prev.get('product_url')}"
                # если этого продукта нет в текущем списке скрейпа, пропускаем
                if prev_prod_key not in product_keys:
                    continue
                changes.append(f"❌ {prev.get('product')} — {prev.get('region')} — продавец пропал: {prev.get('seller')}")
    else:
        print("ℹ️ Нет предыдущего снимка, сохраняю базу и алерты по изменениям будут со следующего запуска.")

    save_snapshot(snapshot_path, curr_snapshot)

    results_dir = Path(__file__).resolve().parent / "RESULTS"
    results_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    output_path = results_dir / f"kaspi_prices_{timestamp}.xlsx"

    alerts_sent = False
    alert_lines: list[str] = []
    if changes:
        alert_lines.append("⚠️ Изменения:")
        alert_lines.extend(changes)
        alert_lines.append("")
    if missing_alerts:
        alert_lines.append("⚠️ Отсутствуют ожидаемые продавцы:")
        for alert in missing_alerts:
            sku_info = f" | SKU: {alert['product_code']}" if alert.get("product_code") else ""
            alert_lines.append(f"{alert['product']} — {alert['city']}{sku_info}")
            alert_lines.append(f"Ожидал: {', '.join(alert['missing'])}")
            current_line = ", ".join(alert.get("found") or ["нет продавцов на карточке"])
            alert_lines.append(f"Сейчас: {current_line}")
            alert_lines.append("")
    if first_snapshot and not alert_lines and alert_only:
        alert_lines.append("📌 Сохранил базовый снимок. Изменения по ценам и продавцам пришлю со следующего запуска.")
    if alert_lines:
        try:
            send_telegram_message("\n".join(alert_lines).strip(), chat_id=chat_id)
            alerts_sent = True
        except Exception as exc:
            print(f"⚠️ Could not send Telegram alerts: {exc}")

    if not alert_only:
        df.to_excel(output_path, index=False)
        print(f"✅ Done! Saved file {output_path.name} in {results_dir}")
        summary = (
            f"Kaspi scrape finished\n"
            f"Rows: {len(df)}\n"
            f"Cities: {', '.join(cities)}\n"
            f"File: {output_path.name}"
        )
        try:
            send_telegram_message(summary, chat_id=chat_id)
        except Exception as exc:
            print(f"⚠️ Could not send Telegram summary: {exc}")
        try:
            send_telegram_file(output_path, caption=summary, chat_id=chat_id)
        except Exception as exc:
            print(f"⚠️ Could not send Telegram file: {exc}")
        # запомним последний полный файл, если скрейп включал базовый список
        if include_base:
            STATE_DIR.mkdir(parents=True, exist_ok=True)
            marker = STATE_DIR / "latest_full.txt"
            try:
                marker.write_text(str(output_path), encoding="utf-8")
            except Exception:
                pass
    else:
        if not alerts_sent:
            print("ℹ️ alert_only enabled and no changes; nothing sent.")

def _resolve_csv_paths(
    base_dir: Path,
    extra_csv_paths: Path | Sequence[Path] | str | None = None,
    include_base: bool = True,
) -> list[Path]:
    base_csv = base_dir / "my_products.csv"
    paths: list[Path] = [base_csv] if include_base else []

    def add_path(p: Path):
        if not p.is_absolute():
            p = base_dir / p
        if p.exists():
            paths.append(p)
        else:
            print(f"ℹ️ Extra CSV not found, skipping: {p}")

    if extra_csv_paths:
        if isinstance(extra_csv_paths, (str, Path)):
            extra_csv_paths = [extra_csv_paths]
        for part in extra_csv_paths:
            if isinstance(part, str):
                for split_part in re.split(r"[;,]", part):
                    split_part = split_part.strip()
                    if split_part:
                        add_path(Path(split_part))
            else:
                add_path(Path(part))
    else:
        extra_env = os.environ.get("EXTRA_PRODUCTS_CSV")
        if extra_env:
            for part in re.split(r"[;,]", extra_env):
                part = part.strip()
                if part:
                    add_path(Path(part))
        else:
            default_extra = base_dir / "my_products_extra.csv"
            if default_extra.exists():
                add_path(default_extra)

    return paths


def run_kaspi_scrape(
    cities: Sequence[str] | None = None,
    extra_csv_paths: Path | Sequence[Path] | str | None = None,
    include_base: bool = True,
    chat_id: str | None = None,
    stop_event: threading.Event | None = None,
    progress_callback=None,
    alert_only: bool = False,
):
    """Удобный вызов из чат-бота: старт скрейпа и алертов."""
    base_dir = Path(__file__).resolve().parent
    csvs = _resolve_csv_paths(base_dir, extra_csv_paths, include_base=include_base)
    if progress_callback:
        progress_callback("start", city=None, done=0)
    scrape_products_from_csv(
        csvs,
        cities=cities or ["Алматы", "Астана", "Шымкент"],
        stop_event=stop_event,
        progress_callback=progress_callback,
        chat_id=chat_id,
        alert_only=alert_only,
        include_base=include_base,
    )
    if stop_event and stop_event.is_set():
        return
    if progress_callback:
        progress_callback("finished", city=None, done=None)


def bot_help_text() -> str:
    """Подсказка для /help в боте."""
    base_dir = Path(__file__).resolve().parent
    default_extra = base_dir / "my_products_extra.csv"
    return (
        "Что могу:\n"
        "• /add <ссылка> | <имя товара> | <продавцы через ;> — добавить товар в временный список\n"
        "• /run — запустить сбор по основному списку + временным\n"
        "• /run_extra — запустить сбор только по временным\n"
        "• /help — показать эту подсказку\n"
        f"По умолчанию города: Алматы, Астана, Шымкент. Доп. список: {default_extra.name} или переменная EXTRA_PRODUCTS_CSV."
    )


if __name__ == "__main__":
    run_kaspi_scrape(cities=["Алматы", "Астана", "Шымкент"])
