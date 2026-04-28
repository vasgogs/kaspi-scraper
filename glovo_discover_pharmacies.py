#!/usr/bin/env python3
"""Discover Glovo pharmacies from a category landing page."""

from __future__ import annotations

import argparse
import csv
import re
import sys
from datetime import datetime
from pathlib import Path

from wolt_brand_search_monitor import build_session, normalize_text, write_csv


BASE_DIR = Path(__file__).resolve().parent
GLOVO_PROJECT_DIR = BASE_DIR / "glovo_project"
DEFAULT_RESULTS_DIR = GLOVO_PROJECT_DIR / "RESULTS"
DEFAULT_STATE_DIR = GLOVO_PROJECT_DIR / "state"
DEFAULT_CATEGORY_URL = "https://glovoapp.com/ru/kz/almaty/categories/apteki-i-kosmetika_3"
DEFAULT_SOURCE_URLS = [
    DEFAULT_CATEGORY_URL,
    "https://glovoapp.com/en/kz/almaty/categories/apteki-i-kosmetika_3",
    "https://glovoapp.com/en/kz/almaty/categories/food_1?type=pharmacy_35365",
    "https://glovoapp.com/ru/kz/almaty/categories/food_1?type=pharmacy_35365",
    "https://glovoapp.com/kz/en/delivery_glovo/pharmacy/",
]
DEFAULT_CITY_SLUG = "almaty"
DEFAULT_LANGUAGE = "ru"
DEFAULT_COUNTRY = "kz"
STORE_URL_PATTERNS = (
    re.compile(r"https://glovoapp\.com/(?:ru|en)/kz/almaty/stores/([^\"?#/]+)", re.IGNORECASE),
    re.compile(r"https://glovoapp\.com/kz/(?:ru|en)/almaty/stores/([^\"?#/]+)", re.IGNORECASE),
    re.compile(r"/(?:ru|en)/kz/almaty/stores/([^\"?#/]+)", re.IGNORECASE),
    re.compile(r"/kz/(?:ru|en)/almaty/stores/([^\"?#/]+)", re.IGNORECASE),
)
STORE_TITLE_PATTERN = re.compile(r"<h1[^>]*>([^<]+)</h1>", re.IGNORECASE)
STORE_ID_PATTERN = re.compile(r"store_id=(\d+)")
ADDRESS_ID_PATTERN = re.compile(r"storeAddressId\\?\":\\?\"?(\d+)")
PHARMACY_HINTS = (
    "аптека",
    "apteka",
    "pharma",
    "фарма",
    "europharma",
    "sadykhan",
    "добрая",
    "со склада",
    "sklada",
)
NON_PHARMACY_HINTS = (
    "yves rocher",
    "ив роше",
    "cosmetic",
    "космет",
    "beauty",
    "parfum",
    "парфюм",
    " lab ",
    "lab ",
    " lab",
    "pcr",
    "анализ",
    "analysis",
)
ADDRESS_TITLES = {"address", "адрес"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Discover pharmacies from a Glovo category page")
    parser.add_argument(
        "--category-url",
        default=DEFAULT_CATEGORY_URL,
        help="Primary Glovo category URL (kept for backward compatibility)",
    )
    parser.add_argument(
        "--source-url",
        action="append",
        default=[],
        help="Extra discovery source URL. Can be passed multiple times.",
    )
    parser.add_argument("--results-dir", default=str(DEFAULT_RESULTS_DIR), help="Folder for discovery outputs")
    parser.add_argument(
        "--catalog-path",
        default=str(DEFAULT_STATE_DIR / "glovo_almaty_pharmacies_catalog.csv"),
        help="Cumulative pharmacy catalog CSV path",
    )
    parser.add_argument("--city-slug", default=DEFAULT_CITY_SLUG, help="City slug used in Glovo URLs")
    parser.add_argument("--language", default=DEFAULT_LANGUAGE, help="Path language code")
    parser.add_argument("--country", default=DEFAULT_COUNTRY, help="Path country code")
    parser.add_argument("--timeout", type=float, default=30.0, help="HTTP timeout in seconds")
    return parser.parse_args()


def extract_store_urls(category_html: str) -> list[str]:
    urls: list[str] = []
    seen: set[str] = set()
    for pattern in STORE_URL_PATTERNS:
        for slug in pattern.findall(category_html):
            slug = (slug or "").strip().strip("/")
            if not slug or slug in seen:
                continue
            seen.add(slug)
            urls.append(f"https://glovoapp.com/ru/kz/almaty/stores/{slug}")
    return urls


def parse_store_page(html: str, store_url: str) -> dict[str, str]:
    title_match = STORE_TITLE_PATTERN.search(html)
    store_id_match = STORE_ID_PATTERN.search(html)
    address_id_match = ADDRESS_ID_PATTERN.search(html)
    slug = store_url.rstrip("/").rsplit("/", 1)[-1]
    title = title_match.group(1).strip() if title_match else slug
    title = re.sub(r"\s+", " ", title).strip()
    return {
        "name": title,
        "slug": slug,
        "store_url": store_url,
        "store_id": store_id_match.group(1) if store_id_match else "",
        "address_id": address_id_match.group(1) if address_id_match else "",
    }


def _extract_label_text(element: dict) -> str:
    if not isinstance(element, dict):
        return ""
    data = element.get("data") or {}
    if not isinstance(data, dict):
        return ""
    text = data.get("text")
    if isinstance(text, str):
        return re.sub(r"\s+", " ", text.replace("\n", ", ")).strip(" ,")
    label = data.get("label") or {}
    if isinstance(label, dict):
        return _extract_label_text(label)
    return ""


def fetch_store_info(session, *, store_id: str, address_id: str, timeout: float) -> dict[str, str]:
    if not store_id or not address_id:
        return {"address": ""}
    url = f"https://api.glovoapp.com/v3/stores/{store_id}/addresses/{address_id}/store_info_screen"
    response = session.get(url, params={"translation": "undefined"}, timeout=timeout)
    response.raise_for_status()
    payload = response.json()
    address = ""
    for section in payload.get("sections") or []:
        if not isinstance(section, dict):
            continue
        data = section.get("data") or {}
        title = normalize_text(data.get("title") or "")
        if title not in ADDRESS_TITLES:
            continue
        for element in data.get("elements") or []:
            address = _extract_label_text(element)
            if address:
                break
        if address:
            break
    return {"address": address}


def classify_store(name: str, slug: str) -> tuple[str, str, str]:
    haystack = normalize_text(f"{name} {slug}".replace("-", " "))
    if any(keyword in haystack for keyword in PHARMACY_HINTS):
        return "1", "pharmacy", ""
    if any(keyword in haystack for keyword in NON_PHARMACY_HINTS):
        return "0", "non_pharmacy", "matched_non_pharmacy_keyword"
    return "1", "unknown", ""


def resolve_source_urls(args: argparse.Namespace) -> list[str]:
    urls: list[str] = []
    seen: set[str] = set()
    for raw in [args.category_url, *args.source_url, *DEFAULT_SOURCE_URLS]:
        url = (raw or "").strip()
        if not url or url in seen:
            continue
        seen.add(url)
        urls.append(url)
    return urls


def load_catalog(path: Path) -> dict[str, dict[str, str]]:
    if not path.exists():
        return {}
    rows: dict[str, dict[str, str]] = {}
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            slug = (row.get("slug") or "").strip()
            if not slug:
                continue
            rows[slug] = {k: (v or "").strip() for k, v in row.items()}
    return rows


def update_catalog(
    *,
    catalog_path: Path,
    rows: list[dict[str, str]],
    checked_at: str,
) -> None:
    catalog = load_catalog(catalog_path)
    for row in rows:
        slug = row.get("slug", "").strip()
        if not slug:
            continue
        source_urls = [value for value in str(row.get("discovery_sources") or "").split(" || ") if value]
        rec = catalog.get(slug)
        if not rec:
            catalog[slug] = {
                "slug": slug,
                "name": row.get("name", ""),
                "store_url": row.get("store_url", ""),
                "store_id": row.get("store_id", ""),
                "address_id": row.get("address_id", ""),
                "address": row.get("address", ""),
                "is_pharmacy": row.get("is_pharmacy", ""),
                "store_type": row.get("store_type", ""),
                "skip_reason": row.get("skip_reason", ""),
                "first_seen": checked_at,
                "last_seen": checked_at,
                "discovery_hits": "1",
                "source_urls": " || ".join(source_urls),
            }
            continue

        rec["name"] = row.get("name", "") or rec.get("name", "")
        rec["store_url"] = row.get("store_url", "") or rec.get("store_url", "")
        rec["store_id"] = row.get("store_id", "") or rec.get("store_id", "")
        rec["address_id"] = row.get("address_id", "") or rec.get("address_id", "")
        rec["address"] = row.get("address", "") or rec.get("address", "")
        rec["is_pharmacy"] = row.get("is_pharmacy", "") or rec.get("is_pharmacy", "")
        rec["store_type"] = row.get("store_type", "") or rec.get("store_type", "")
        rec["skip_reason"] = row.get("skip_reason", "") or rec.get("skip_reason", "")
        rec["last_seen"] = checked_at
        rec["discovery_hits"] = str(int(rec.get("discovery_hits") or "0") + 1)
        known_sources = {value for value in str(rec.get("source_urls") or "").split(" || ") if value}
        known_sources.update(source_urls)
        rec["source_urls"] = " || ".join(sorted(known_sources))

    catalog_rows = sorted(
        catalog.values(),
        key=lambda item: (
            normalize_text(item.get("is_pharmacy", "")) not in {"1", "true", "yes"},
            normalize_text(item.get("name", "")),
            item.get("slug", ""),
        ),
    )
    write_csv(
        catalog_path,
        catalog_rows,
        [
            "slug",
            "name",
            "store_url",
            "store_id",
            "address_id",
            "address",
            "is_pharmacy",
            "store_type",
            "skip_reason",
            "first_seen",
            "last_seen",
            "discovery_hits",
            "source_urls",
        ],
    )


def run() -> int:
    args = parse_args()
    results_dir = Path(args.results_dir).expanduser().resolve()
    results_dir.mkdir(parents=True, exist_ok=True)
    catalog_path = Path(args.catalog_path).expanduser().resolve()
    catalog_path.parent.mkdir(parents=True, exist_ok=True)

    session = build_session()
    source_urls = resolve_source_urls(args)
    store_to_sources: dict[str, set[str]] = {}
    for source_url in source_urls:
        resp = session.get(source_url, timeout=args.timeout)
        resp.raise_for_status()
        extracted = extract_store_urls(resp.text)
        print(f"[source] {source_url} -> stores={len(extracted)}")
        for store_url in extracted:
            store_to_sources.setdefault(store_url, set()).add(source_url)

    store_urls = sorted(store_to_sources)
    if not store_urls:
        raise RuntimeError(f"No stores found across discovery sources: {source_urls}")

    rows: list[dict[str, str]] = []
    for idx, store_url in enumerate(store_urls, start=1):
        resp = session.get(store_url, timeout=args.timeout)
        resp.raise_for_status()
        row = parse_store_page(resp.text, store_url)
        store_info = fetch_store_info(
            session,
            store_id=row.get("store_id", ""),
            address_id=row.get("address_id", ""),
            timeout=args.timeout,
        )
        row.update(store_info)
        row["is_pharmacy"], row["store_type"], row["skip_reason"] = classify_store(row["name"], row["slug"])
        row["category_url"] = args.category_url
        row["discovery_sources"] = " || ".join(sorted(store_to_sources.get(store_url) or []))
        row["city_slug"] = args.city_slug
        row["language"] = args.language
        row["country"] = args.country
        rows.append(row)
        print(
            f"[{idx}/{len(store_urls)}] {row['name']} | slug={row['slug']} | "
            f"store_id={row['store_id']} | address_id={row['address_id']} | "
            f"is_pharmacy={row['is_pharmacy']} | sources={len(store_to_sources.get(store_url) or [])}"
        )

    now = datetime.now()
    checked_at = now.isoformat(timespec="seconds")
    timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")
    output_path = results_dir / f"glovo_almaty_pharmacies_{timestamp}.csv"
    write_csv(
        output_path,
        rows,
        [
            "name",
            "slug",
            "store_url",
            "store_id",
            "address_id",
            "address",
            "is_pharmacy",
            "store_type",
            "skip_reason",
            "category_url",
            "discovery_sources",
            "city_slug",
            "language",
            "country",
        ],
    )
    update_catalog(catalog_path=catalog_path, rows=rows, checked_at=checked_at)
    print(f"Discovered pharmacies: {len(rows)}")
    print(f"Output: {output_path}")
    print(f"Cumulative catalog: {catalog_path}")
    return 0


def main() -> None:
    try:
        code = run()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(2)
    sys.exit(code)


if __name__ == "__main__":
    main()
