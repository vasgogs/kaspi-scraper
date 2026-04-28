#!/usr/bin/env python3
"""Search a brand across Glovo pharmacies and maintain a growing product catalog."""

from __future__ import annotations

import argparse
import csv
import os
import sys
import time
from collections import Counter
from datetime import datetime
from pathlib import Path
from urllib.parse import quote_plus
from uuid import uuid4

import requests

from wolt_brand_search_monitor import (
    DEFAULT_FONT_PATH,
    build_best_offers,
    build_brand_reference_rows,
    build_vitrum_reference_rows,
    build_session,
    canonicalize_brand_name,
    export_active_ingredient_reference,
    item_matches_brand_name,
    load_catalog,
    load_env_from_file,
    normalize_text,
    save_catalog,
    send_telegram_document,
    send_telegram_message,
    send_telegram_photo,
    slugify_token,
    write_csv,
)


BASE_DIR = Path(__file__).resolve().parent
GLOVO_PROJECT_DIR = BASE_DIR / "glovo_project"
DEFAULT_RESULTS_DIR = GLOVO_PROJECT_DIR / "RESULTS"
DEFAULT_ITEMS_CATALOG = GLOVO_PROJECT_DIR / "state" / "glovo_item_ids_catalog.csv"
DEFAULT_ACTIVE_REFERENCE = GLOVO_PROJECT_DIR / "state" / "glovo_active_ingredient_reference.csv"
DEFAULT_CITY_SLUG = "almaty"
DEFAULT_LANGUAGE = "ru"
DEFAULT_COUNTRY = "kz"
DEFAULT_FONT = DEFAULT_FONT_PATH


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run brand search through Glovo pharmacies and update product catalog."
    )
    parser.add_argument("--pharmacies-csv", required=True, help="CSV with Glovo pharmacy list")
    parser.add_argument("--brand", required=True, help="Brand label for catalog")
    parser.add_argument(
        "--query",
        action="append",
        default=[],
        help="Search query. Can be passed multiple times; defaults to --brand value.",
    )
    parser.add_argument("--city-slug", default=DEFAULT_CITY_SLUG, help="City slug for links")
    parser.add_argument("--language", default=DEFAULT_LANGUAGE, help="Language code")
    parser.add_argument("--country", default=DEFAULT_COUNTRY, help="Country code")
    parser.add_argument("--timeout", type=float, default=25.0, help="HTTP timeout in seconds")
    parser.add_argument("--sleep-ms", type=int, default=250, help="Pause between pharmacy requests, ms")
    parser.add_argument("--retry-attempts", type=int, default=4, help="Extra retries after rate limits")
    parser.add_argument("--retry-backoff", type=float, default=2.0, help="Base sleep seconds for 429 retry")
    parser.add_argument("--results-dir", default=str(DEFAULT_RESULTS_DIR), help="Folder for output reports")
    parser.add_argument("--items-catalog", default=str(DEFAULT_ITEMS_CATALOG), help="Global item catalog CSV path")
    parser.add_argument(
        "--active-reference-csv",
        default=str(DEFAULT_ACTIVE_REFERENCE),
        help="Active ingredient reference output CSV",
    )
    parser.add_argument(
        "--send-telegram",
        action="store_true",
        help="Send summary + files to Telegram (uses TELEGRAM_BOT_TOKEN/TELEGRAM_CHAT_ID by default)",
    )
    parser.add_argument("--telegram-bot-token", default="", help="Telegram bot token override")
    parser.add_argument("--telegram-chat-id", default="", help="Telegram chat_id override")
    parser.add_argument(
        "--telegram-top-n",
        type=int,
        default=12,
        help="How many best offers to include in Telegram message (default: %(default)s)",
    )
    return parser.parse_args()


def price_minor_from_amount(amount: object) -> int | None:
    if isinstance(amount, int):
        return amount * 100
    if isinstance(amount, float):
        return int(round(amount * 100))
    if isinstance(amount, str):
        raw = amount.strip().replace("\xa0", " ")
        if not raw:
            return None
        raw = raw.replace("₸", "").replace(" ", "").replace(",", ".")
        try:
            return int(round(float(raw) * 100))
        except Exception:
            return None
    return None


def price_major_from_minor(value: int | None) -> str:
    if value is None:
        return ""
    return f"{value / 100.0:.2f}"


def read_pharmacies(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        raise FileNotFoundError(f"Pharmacies CSV not found: {path}")
    rows: list[dict[str, str]] = []
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            row = {k: (v or "").strip() for k, v in row.items()}
            is_pharmacy = normalize_text(row.get("is_pharmacy", "1"))
            if is_pharmacy in {"0", "false", "no", "off", "нет"}:
                continue
            if not row.get("slug") or not row.get("store_id") or not row.get("address_id"):
                continue
            rows.append(row)
    if not rows:
        raise ValueError(f"No valid pharmacy rows found in {path}")
    return rows


def fetch_search_results(
    session: requests.Session,
    *,
    store_id: str,
    address_id: str,
    query: str,
    timeout: float,
    retry_attempts: int,
    retry_backoff: float,
) -> list[dict]:
    url = f"https://api.glovoapp.com/v3/stores/{store_id}/addresses/{address_id}/search"
    last_exc: Exception | None = None

    for attempt in range(1, max(1, retry_attempts) + 1):
        try:
            response = session.get(
                url,
                params={"query": query, "searchId": str(uuid4())},
                timeout=timeout,
            )
            if response.status_code == 429 and attempt < retry_attempts:
                retry_after = response.headers.get("Retry-After", "").strip()
                try:
                    delay = float(retry_after)
                except ValueError:
                    delay = retry_backoff * attempt
                time.sleep(max(delay, retry_backoff))
                continue

            response.raise_for_status()
            payload = response.json()
            results = payload.get("results") if isinstance(payload, dict) else []
            products: list[dict] = []
            if not isinstance(results, list):
                return products
            for entry in results:
                if not isinstance(entry, dict):
                    continue
                for product in entry.get("products") or []:
                    if isinstance(product, dict):
                        products.append(product)
            return products
        except requests.RequestException as exc:
            last_exc = exc
            if "429" in str(exc) and attempt < retry_attempts:
                time.sleep(retry_backoff * attempt)
                continue
            raise

    if last_exc is not None:
        raise last_exc
    return []


def export_brand_reference(
    catalog: dict[tuple[str, str], dict[str, object]],
    brand: str,
    reference_path: Path,
    canonical_path: Path,
    unmapped_path: Path,
) -> dict[str, object]:
    detailed_rows, canonical_rows, unmapped_rows = build_brand_reference_rows(catalog=catalog, brand=brand)
    write_csv(
        reference_path,
        detailed_rows,
        [
            "brand",
            "item_id",
            "item_name",
            "first_seen",
            "last_seen",
            "seen_runs",
            "last_seen_venue_slug",
            "last_seen_pharmacy",
            "canonical_sku",
            "canonical_name",
            "product_line",
            "dosage_or_volume",
            "pack_size",
            "form_factor",
            "flavor",
            "active_ingredient",
            "confidence",
            "rule",
            "normalized_name",
        ],
    )
    write_csv(
        canonical_path,
        canonical_rows,
        [
            "canonical_sku",
            "canonical_name",
            "product_line",
            "dosage_or_volume",
            "pack_size",
            "form_factor",
            "flavor",
            "active_ingredient",
            "item_ids_count",
            "aliases_count",
            "seen_runs_total",
            "confidence_min",
            "rules_used",
            "item_ids",
            "aliases_examples",
        ],
    )
    write_csv(
        unmapped_path,
        unmapped_rows,
        [
            "brand",
            "item_id",
            "item_name",
            "canonical_sku",
            "canonical_name",
            "active_ingredient",
            "confidence",
            "rule",
            "normalized_name",
            "first_seen",
            "last_seen",
            "last_seen_pharmacy",
        ],
    )
    return {
        "reference_path": reference_path,
        "canonical_path": canonical_path,
        "unmapped_path": unmapped_path,
        "detailed_count": len(detailed_rows),
        "canonical_count": len(canonical_rows),
        "unmapped_count": len(unmapped_rows),
    }


def export_vitrum_reference(
    catalog: dict[tuple[str, str], dict[str, object]],
    reference_path: Path,
    canonical_path: Path,
    unmapped_path: Path,
) -> dict[str, object]:
    detailed_rows, canonical_rows, unmapped_rows = build_vitrum_reference_rows(catalog)
    write_csv(
        reference_path,
        detailed_rows,
        [
            "brand",
            "item_id",
            "item_name",
            "first_seen",
            "last_seen",
            "seen_runs",
            "last_seen_venue_slug",
            "last_seen_pharmacy",
            "canonical_sku",
            "canonical_name",
            "product_line",
            "dosage_or_volume",
            "pack_size",
            "form_factor",
            "flavor",
            "active_ingredient",
            "confidence",
            "rule",
            "is_vitrum",
            "normalized_name",
        ],
    )
    write_csv(
        canonical_path,
        canonical_rows,
        [
            "canonical_sku",
            "canonical_name",
            "product_line",
            "dosage_or_volume",
            "pack_size",
            "form_factor",
            "flavor",
            "active_ingredient",
            "item_ids_count",
            "aliases_count",
            "seen_runs_total",
            "confidence_min",
            "rules_used",
            "item_ids",
            "aliases_examples",
        ],
    )
    write_csv(
        unmapped_path,
        unmapped_rows,
        [
            "brand",
            "item_id",
            "item_name",
            "canonical_sku",
            "canonical_name",
            "active_ingredient",
            "confidence",
            "rule",
            "normalized_name",
            "first_seen",
            "last_seen",
            "last_seen_pharmacy",
        ],
    )
    return {
        "reference_path": reference_path,
        "canonical_path": canonical_path,
        "unmapped_path": unmapped_path,
        "detailed_count": len(detailed_rows),
        "canonical_count": len(canonical_rows),
        "unmapped_count": len(unmapped_rows),
    }


def create_visual_summary(
    out_path: Path,
    brand: str,
    query: str,
    checked_at: str,
    pharmacies_count: int,
    item_rows_count: int,
    in_stock_count: int,
    not_found_pharmacies: int,
    new_ids_count: int,
    top_offers: list[dict],
) -> bool:
    try:
        from PIL import Image, ImageDraw, ImageFont
    except Exception:
        return False

    width = 1500
    row_h = 42
    top_n = min(len(top_offers), 12)
    height = 270 + top_n * row_h + 50
    img = Image.new("RGB", (width, height), "#F5F7FB")
    draw = ImageDraw.Draw(img)

    font_regular = ImageFont.load_default()
    font_bold = ImageFont.load_default()
    if DEFAULT_FONT.exists():
        try:
            font_regular = ImageFont.truetype(str(DEFAULT_FONT), 22)
            font_bold = ImageFont.truetype(str(DEFAULT_FONT), 30)
        except Exception:
            pass

    draw.rectangle((0, 0, width, 95), fill="#2563EB")
    draw.text((28, 26), f"Glovo Price Monitor | {brand}", fill="white", font=font_bold)
    draw.text((28, 62), f"query: {query} | {checked_at}", fill="white", font=font_regular)

    cards = [
        f"Pharmacies: {pharmacies_count}",
        f"Items found: {item_rows_count}",
        f"In stock: {in_stock_count}",
        f"No matches: {not_found_pharmacies}",
        f"New items: {new_ids_count}",
    ]
    x = 26
    y = 118
    for card in cards:
        draw.rounded_rectangle((x, y, x + 250, y + 58), radius=12, fill="white", outline="#D8DEE9")
        draw.text((x + 14, y + 17), card, fill="#1F2937", font=font_regular)
        x += 264

    draw.text((28, 196), "Top best prices", fill="#111827", font=font_bold)
    y0 = 230
    draw.rectangle((24, y0, width - 24, y0 + row_h), fill="#E2E8F0")
    draw.text((36, y0 + 10), "#", fill="#111827", font=font_regular)
    draw.text((74, y0 + 10), "Product", fill="#111827", font=font_regular)
    draw.text((970, y0 + 10), "Price", fill="#111827", font=font_regular)
    draw.text((1120, y0 + 10), "Pharmacy", fill="#111827", font=font_regular)

    for idx, row in enumerate(top_offers[:top_n], start=1):
        y = y0 + idx * row_h
        fill = "#FFFFFF" if idx % 2 else "#F8FAFC"
        draw.rectangle((24, y, width - 24, y + row_h), fill=fill)

        name = str(row.get("item_name") or "")
        if len(name) > 62:
            name = name[:59] + "..."
        pharmacy = str(row.get("pharmacy") or "")
        if len(pharmacy) > 40:
            pharmacy = pharmacy[:37] + "..."
        price_text = str(row.get("price") or "")
        if price_text:
            price_text = f"{price_text} KZT"

        draw.text((36, y + 10), str(idx), fill="#111827", font=font_regular)
        draw.text((74, y + 10), name, fill="#111827", font=font_regular)
        draw.text((970, y + 10), price_text, fill="#111827", font=font_regular)
        draw.text((1120, y + 10), pharmacy, fill="#111827", font=font_regular)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(out_path, format="PNG")
    return True


def build_telegram_text(
    *,
    brand: str,
    query: str,
    checked_at: str,
    pharmacies_count: int,
    item_rows: list[dict],
    summary_rows: list[dict],
    new_ids_count: int,
    top_offers: list[dict],
    top_n: int,
) -> str:
    status_counts = Counter(str(row.get("status", "")) for row in item_rows)
    not_found_pharmacies = sum(1 for row in summary_rows if int(row.get("matched_items") or 0) == 0 and not row.get("error"))
    lines = [
        "Glovo prices report",
        f"Brand: {brand}",
        f"Query: {query}",
        f"Run: {checked_at}",
        "",
        f"Pharmacies: {pharmacies_count}",
        f"Items found: {len(item_rows)}",
        f"In stock: {status_counts.get('in_stock', 0)}",
        f"Unavailable: {status_counts.get('unavailable', 0)}",
        f"No matches pharmacies: {not_found_pharmacies}",
        f"New item_id this run: {new_ids_count}",
        "",
        f"Top {min(top_n, len(top_offers))} best offers:",
    ]
    for idx, row in enumerate(top_offers[:top_n], start=1):
        lines.append(f"{idx}. {row.get('item_name')} | {row.get('price')} KZT | {row.get('pharmacy')}")
        lines.append(f"   Product: {row.get('product_link')}")
        lines.append(f"   Venue: {row.get('venue_url')}")
    return "\n".join(lines)


def run() -> int:
    args = parse_args()
    load_env_from_file()
    brand = args.brand.strip()
    if not brand:
        raise ValueError("Brand must not be empty")
    queries = [value.strip() for value in args.query if (value or "").strip()]
    if not queries:
        queries = [brand]
    query_label = " | ".join(queries)

    brand_norm = normalize_text(brand)
    is_vitrum_brand = brand_norm == "vitrum" or "витрум" in brand_norm
    strict_brand_filter = True

    pharmacies = read_pharmacies(Path(args.pharmacies_csv).expanduser().resolve())
    session = build_session()
    checked_at = datetime.now().isoformat(timespec="seconds")

    item_rows: list[dict] = []
    summary_rows: list[dict] = []
    errors = 0

    for idx, pharmacy in enumerate(pharmacies, start=1):
        slug = pharmacy["slug"]
        name = pharmacy["name"]
        store_url = pharmacy["store_url"]
        pharmacy_address = pharmacy.get("address", "")
        try:
            counts = {"in_stock": 0, "unavailable": 0}
            raw_items_count = 0
            filtered_out_count = 0
            matched_items_count = 0
            seen_ids: set[str] = set()
            for query in queries:
                products = fetch_search_results(
                    session=session,
                    store_id=pharmacy["store_id"],
                    address_id=pharmacy["address_id"],
                    query=query,
                    timeout=args.timeout,
                    retry_attempts=args.retry_attempts,
                    retry_backoff=args.retry_backoff,
                )

                for product in products:
                    if not isinstance(product, dict):
                        continue
                    raw_items_count += 1
                    item_id = str(product.get("id") or "").strip()
                    item_name = str(product.get("name") or "").strip()
                    if not item_id or not item_name:
                        continue
                    if item_id in seen_ids:
                        continue
                    seen_ids.add(item_id)
                    if strict_brand_filter and not item_matches_brand_name(brand, item_name):
                        filtered_out_count += 1
                        continue

                    price_minor = price_minor_from_amount(product.get("price"))
                    status = "in_stock" if price_minor is not None else "unavailable"
                    counts[status] += 1
                    matched_items_count += 1
                    canonical = canonicalize_brand_name(brand, item_name)
                    search_link = f"{store_url}?search={quote_plus(item_name)}"

                    item_rows.append(
                        {
                            "checked_at": checked_at,
                            "brand": brand,
                            "query": query,
                            "pharmacy": name,
                            "pharmacy_address": pharmacy_address,
                            "venue_slug": slug,
                            "venue_url": store_url,
                            "item_id": item_id,
                            "external_id": str(product.get("externalId") or "").strip(),
                            "item_name": item_name,
                            "status": status,
                            "disable_text": "",
                            "price_minor": price_minor if price_minor is not None else "",
                            "price": price_major_from_minor(price_minor),
                            "price_display": str(((product.get("priceInfo") or {}).get("displayText")) or "").strip(),
                            "product_link": search_link,
                            "search_link": search_link,
                            "canonical_sku": canonical.get("canonical_sku", ""),
                            "canonical_name": canonical.get("canonical_name", ""),
                            "canonical_line": canonical.get("product_line", ""),
                            "canonical_pack_size": canonical.get("pack_size", ""),
                            "canonical_form_factor": canonical.get("form_factor", ""),
                            "canonical_dosage_or_volume": canonical.get("dosage_or_volume", ""),
                            "canonical_flavor": canonical.get("flavor", ""),
                            "canonical_active_ingredient": canonical.get("active_ingredient", ""),
                            "canonical_confidence": canonical.get("confidence", ""),
                            "canonical_rule": canonical.get("rule", ""),
                        }
                    )

            summary_rows.append(
                {
                    "checked_at": checked_at,
                    "brand": brand,
                    "query": query_label,
                    "pharmacy": name,
                    "pharmacy_address": pharmacy_address,
                    "venue_slug": slug,
                    "venue_url": store_url,
                    "raw_items": raw_items_count,
                    "filtered_out_items": filtered_out_count,
                    "matched_items": matched_items_count,
                    "in_stock_items": counts["in_stock"],
                    "out_of_stock_items": 0,
                    "unavailable_items": counts["unavailable"],
                    "error": "",
                }
            )
        except Exception as exc:
            errors += 1
            summary_rows.append(
                {
                    "checked_at": checked_at,
                    "brand": brand,
                    "query": query_label,
                    "pharmacy": name,
                    "pharmacy_address": pharmacy_address,
                    "venue_slug": slug,
                    "venue_url": store_url,
                    "raw_items": 0,
                    "filtered_out_items": 0,
                    "matched_items": 0,
                    "in_stock_items": 0,
                    "out_of_stock_items": 0,
                    "unavailable_items": 0,
                    "error": str(exc),
                }
            )

        if args.sleep_ms > 0:
            time.sleep(args.sleep_ms / 1000.0)
        if idx % 10 == 0 or idx == len(pharmacies):
            print(f"[{idx}/{len(pharmacies)}] processed, item rows={len(item_rows)}, errors={errors}")

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    results_dir = Path(args.results_dir).expanduser().resolve()
    results_dir.mkdir(parents=True, exist_ok=True)
    safe_brand = slugify_token(brand) or "brand"
    base = f"glovo_brand_{safe_brand}_{timestamp}"

    items_report_path = results_dir / f"{base}_items.csv"
    summary_report_path = results_dir / f"{base}_summary.csv"

    write_csv(
        items_report_path,
        item_rows,
        [
            "checked_at",
            "brand",
            "query",
            "pharmacy",
            "pharmacy_address",
            "venue_slug",
            "venue_url",
            "item_id",
            "external_id",
            "item_name",
            "status",
            "disable_text",
            "price_minor",
            "price",
            "price_display",
            "product_link",
            "search_link",
            "canonical_sku",
            "canonical_name",
            "canonical_line",
            "canonical_pack_size",
            "canonical_form_factor",
            "canonical_dosage_or_volume",
            "canonical_flavor",
            "canonical_active_ingredient",
            "canonical_confidence",
            "canonical_rule",
        ],
    )

    write_csv(
        summary_report_path,
        summary_rows,
        [
            "checked_at",
            "brand",
            "query",
            "pharmacy",
            "pharmacy_address",
            "venue_slug",
            "venue_url",
            "raw_items",
            "filtered_out_items",
            "matched_items",
            "in_stock_items",
            "out_of_stock_items",
            "unavailable_items",
            "error",
        ],
    )

    catalog_path = Path(args.items_catalog).expanduser().resolve()
    catalog_path.parent.mkdir(parents=True, exist_ok=True)
    catalog = load_catalog(catalog_path)
    observed_now: dict[tuple[str, str], dict[str, str]] = {}
    new_rows: list[dict] = []

    for row in item_rows:
        key = (normalize_text(str(row["brand"])), str(row["item_id"]))
        observed_now[key] = {
            "brand": str(row["brand"]),
            "item_id": str(row["item_id"]),
            "item_name": str(row["item_name"]),
            "venue_slug": str(row["venue_slug"]),
            "pharmacy": str(row["pharmacy"]),
            "canonical_sku": str(row.get("canonical_sku") or ""),
            "canonical_name": str(row.get("canonical_name") or ""),
            "canonical_line": str(row.get("canonical_line") or ""),
            "canonical_pack_size": str(row.get("canonical_pack_size") or ""),
            "canonical_form_factor": str(row.get("canonical_form_factor") or ""),
            "canonical_dosage_or_volume": str(row.get("canonical_dosage_or_volume") or ""),
            "canonical_flavor": str(row.get("canonical_flavor") or ""),
            "canonical_active_ingredient": str(row.get("canonical_active_ingredient") or ""),
            "canonical_confidence": str(row.get("canonical_confidence") or ""),
            "canonical_rule": str(row.get("canonical_rule") or ""),
        }

    for key, obs in observed_now.items():
        if key not in catalog:
            catalog[key] = {
                "brand": obs["brand"],
                "item_id": obs["item_id"],
                "item_name": obs["item_name"],
                "first_seen": checked_at,
                "last_seen": checked_at,
                "seen_runs": 1,
                "last_seen_venue_slug": obs["venue_slug"],
                "last_seen_pharmacy": obs["pharmacy"],
                "canonical_sku": obs["canonical_sku"],
                "canonical_name": obs["canonical_name"],
                "canonical_line": obs["canonical_line"],
                "canonical_pack_size": obs["canonical_pack_size"],
                "canonical_form_factor": obs["canonical_form_factor"],
                "canonical_dosage_or_volume": obs["canonical_dosage_or_volume"],
                "canonical_flavor": obs["canonical_flavor"],
                "canonical_active_ingredient": obs["canonical_active_ingredient"],
                "canonical_confidence": obs["canonical_confidence"],
                "canonical_rule": obs["canonical_rule"],
            }
            new_rows.append(
                {
                    "brand": obs["brand"],
                    "item_id": obs["item_id"],
                    "item_name": obs["item_name"],
                    "canonical_sku": obs["canonical_sku"],
                    "canonical_name": obs["canonical_name"],
                    "canonical_active_ingredient": obs["canonical_active_ingredient"],
                    "canonical_confidence": obs["canonical_confidence"],
                    "first_seen": checked_at,
                }
            )
        else:
            rec = catalog[key]
            rec["item_name"] = obs["item_name"] or rec.get("item_name", "")
            rec["last_seen"] = checked_at
            rec["seen_runs"] = int(rec.get("seen_runs", 0)) + 1
            rec["last_seen_venue_slug"] = obs["venue_slug"]
            rec["last_seen_pharmacy"] = obs["pharmacy"]
            rec["canonical_sku"] = obs["canonical_sku"] or rec.get("canonical_sku", "")
            rec["canonical_name"] = obs["canonical_name"] or rec.get("canonical_name", "")
            rec["canonical_line"] = obs["canonical_line"] or rec.get("canonical_line", "")
            rec["canonical_pack_size"] = obs["canonical_pack_size"] or rec.get("canonical_pack_size", "")
            rec["canonical_form_factor"] = obs["canonical_form_factor"] or rec.get("canonical_form_factor", "")
            rec["canonical_dosage_or_volume"] = obs["canonical_dosage_or_volume"] or rec.get("canonical_dosage_or_volume", "")
            rec["canonical_flavor"] = obs["canonical_flavor"] or rec.get("canonical_flavor", "")
            rec["canonical_active_ingredient"] = obs["canonical_active_ingredient"] or rec.get("canonical_active_ingredient", "")
            rec["canonical_confidence"] = obs["canonical_confidence"] or rec.get("canonical_confidence", "")
            rec["canonical_rule"] = obs["canonical_rule"] or rec.get("canonical_rule", "")

    for rec in catalog.values():
        rec_brand_raw = str(rec.get("brand") or "").strip()
        if normalize_text(rec_brand_raw) != brand_norm:
            continue
        rec_name = str(rec.get("item_name") or "")
        canonical = canonicalize_brand_name(rec_brand_raw or brand, rec_name)
        rec["canonical_sku"] = canonical.get("canonical_sku", "")
        rec["canonical_name"] = canonical.get("canonical_name", "")
        rec["canonical_line"] = canonical.get("product_line", "")
        rec["canonical_pack_size"] = canonical.get("pack_size", "")
        rec["canonical_form_factor"] = canonical.get("form_factor", "")
        rec["canonical_dosage_or_volume"] = canonical.get("dosage_or_volume", "")
        rec["canonical_flavor"] = canonical.get("flavor", "")
        rec["canonical_active_ingredient"] = canonical.get("active_ingredient", "")
        rec["canonical_confidence"] = canonical.get("confidence", "")
        rec["canonical_rule"] = canonical.get("rule", "")

    removed_noise = 0
    if strict_brand_filter:
        to_remove: list[tuple[str, str]] = []
        for key, rec in catalog.items():
            rec_brand_raw = str(rec.get("brand") or "").strip()
            if normalize_text(rec_brand_raw) != brand_norm:
                continue
            if not item_matches_brand_name(rec_brand_raw or brand, str(rec.get("item_name") or "")):
                to_remove.append(key)
        for key in to_remove:
            catalog.pop(key, None)
        removed_noise = len(to_remove)

    save_catalog(catalog_path, catalog)
    active_reference_path = Path(args.active_reference_csv).expanduser().resolve()
    active_reference_path.parent.mkdir(parents=True, exist_ok=True)
    export_active_ingredient_reference(active_reference_path)

    brand_export: dict[str, object] | None = None
    vitrum_export: dict[str, object] | None = None
    if is_vitrum_brand:
        vitrum_export = export_vitrum_reference(
            catalog=catalog,
            reference_path=catalog_path.parent / "glovo_vitrum_item_reference.csv",
            canonical_path=catalog_path.parent / "glovo_vitrum_canonical_catalog.csv",
            unmapped_path=catalog_path.parent / "glovo_vitrum_unmapped.csv",
        )
    else:
        brand_export = export_brand_reference(
            catalog=catalog,
            brand=brand,
            reference_path=catalog_path.parent / f"glovo_{safe_brand}_item_reference.csv",
            canonical_path=catalog_path.parent / f"glovo_{safe_brand}_canonical_catalog.csv",
            unmapped_path=catalog_path.parent / f"glovo_{safe_brand}_unmapped.csv",
        )

    new_items_path = results_dir / f"{base}_new_item_ids.csv"
    write_csv(
        new_items_path,
        new_rows,
        [
            "brand",
            "item_id",
            "item_name",
            "canonical_sku",
            "canonical_name",
            "canonical_active_ingredient",
            "canonical_confidence",
            "first_seen",
        ],
    )

    top_offers = build_best_offers(item_rows)
    best_prices_path = results_dir / f"{base}_best_prices.csv"
    write_csv(
        best_prices_path,
        top_offers,
        [
            "checked_at",
            "brand",
            "query",
            "pharmacy",
            "pharmacy_address",
            "venue_slug",
            "venue_url",
            "item_id",
            "external_id",
            "item_name",
            "status",
            "price_minor",
            "price",
            "price_display",
            "product_link",
            "search_link",
            "canonical_sku",
            "canonical_name",
            "canonical_line",
            "canonical_pack_size",
            "canonical_form_factor",
            "canonical_dosage_or_volume",
            "canonical_flavor",
            "canonical_active_ingredient",
            "canonical_confidence",
            "canonical_rule",
        ],
    )

    status_counts = Counter(str(row.get("status", "")) for row in item_rows)
    not_found_pharmacies = sum(1 for row in summary_rows if int(row.get("matched_items") or 0) == 0 and not row.get("error"))
    dashboard_path = results_dir / f"{base}_dashboard.png"
    dashboard_ok = create_visual_summary(
        out_path=dashboard_path,
        brand=brand,
        query=query_label,
        checked_at=checked_at,
        pharmacies_count=len(pharmacies),
        item_rows_count=len(item_rows),
        in_stock_count=status_counts.get("in_stock", 0),
        not_found_pharmacies=not_found_pharmacies,
        new_ids_count=len(new_rows),
        top_offers=top_offers,
    )

    telegram_text = build_telegram_text(
        brand=brand,
        query=query_label,
        checked_at=checked_at,
        pharmacies_count=len(pharmacies),
        item_rows=item_rows,
        summary_rows=summary_rows,
        new_ids_count=len(new_rows),
        top_offers=top_offers,
        top_n=max(1, args.telegram_top_n),
    )

    if args.send_telegram:
        token = args.telegram_bot_token.strip() or os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
        chat_id = args.telegram_chat_id.strip() or os.environ.get("TELEGRAM_CHAT_ID", "").strip()
        if not token or not chat_id:
            raise RuntimeError("Telegram send enabled, but TELEGRAM_BOT_TOKEN/TELEGRAM_CHAT_ID is missing")
        if dashboard_ok and dashboard_path.exists():
            send_telegram_photo(token, chat_id, dashboard_path, caption=f"Glovo {brand}: dashboard")
        send_telegram_message(token, chat_id, telegram_text)
        send_telegram_document(token, chat_id, best_prices_path, caption=f"Glovo {brand}: best prices + links")
        send_telegram_document(token, chat_id, items_report_path, caption=f"Glovo {brand}: full items report")
        send_telegram_document(token, chat_id, new_items_path, caption=f"Glovo {brand}: new item_ids ({len(new_rows)})")
        if vitrum_export:
            send_telegram_document(
                token,
                chat_id,
                Path(vitrum_export["canonical_path"]),
                caption=f"Glovo {brand}: canonical SKU catalog",
            )
        elif brand_export:
            send_telegram_document(
                token,
                chat_id,
                Path(brand_export["canonical_path"]),
                caption=f"Glovo {brand}: canonical SKU catalog",
            )

    print(f"Checked pharmacies: {len(pharmacies)}")
    print(f"Item rows: {len(item_rows)}")
    print(f"Pharmacy errors: {errors}")
    print(f"Items report: {items_report_path}")
    print(f"Summary report: {summary_report_path}")
    print(f"Best prices report: {best_prices_path}")
    print(f"Updated item catalog: {catalog_path}")
    print(f"Active ingredients reference: {active_reference_path}")
    if strict_brand_filter:
        print(f"Removed noisy non-brand IDs from catalog: {removed_noise}")
    print(f"New item IDs this run: {len(new_rows)}")
    print(f"New IDs file: {new_items_path}")
    if vitrum_export:
        print(
            "Vitrum reference: "
            f"{vitrum_export['reference_path']} "
            f"(rows={vitrum_export['detailed_count']})"
        )
        print(
            "Vitrum canonical catalog: "
            f"{vitrum_export['canonical_path']} "
            f"(sku={vitrum_export['canonical_count']}, unmapped={vitrum_export['unmapped_count']})"
        )
        print(f"Vitrum unmapped: {vitrum_export['unmapped_path']}")
    elif brand_export:
        print(
            "Brand reference: "
            f"{brand_export['reference_path']} "
            f"(rows={brand_export['detailed_count']})"
        )
        print(
            "Brand canonical catalog: "
            f"{brand_export['canonical_path']} "
            f"(sku={brand_export['canonical_count']}, unmapped={brand_export['unmapped_count']})"
        )
        print(f"Brand unmapped: {brand_export['unmapped_path']}")
    if dashboard_ok:
        print(f"Dashboard image: {dashboard_path}")
    if args.send_telegram:
        print("Telegram: sent")
    return 0 if errors == 0 else 1


def main() -> None:
    try:
        code = run()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(2)
    sys.exit(code)


if __name__ == "__main__":
    main()
