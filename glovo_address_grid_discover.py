#!/usr/bin/env python3
"""Discover Glovo pharmacies across Almaty using address-driven delivery context."""

from __future__ import annotations

import argparse
import csv
import html
import json
import math
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any

import requests

from glovo_discover_pharmacies import (
    DEFAULT_CATEGORY_URL,
    DEFAULT_CITY_SLUG,
    DEFAULT_COUNTRY,
    DEFAULT_LANGUAGE,
    DEFAULT_SOURCE_URLS,
    classify_store,
    extract_store_urls,
    fetch_store_info,
    load_catalog,
    parse_store_page,
    update_catalog,
)
from wolt_brand_search_monitor import build_session, normalize_text, write_csv


BASE_DIR = Path(__file__).resolve().parent
GLOVO_PROJECT_DIR = BASE_DIR / "glovo_project"
DEFAULT_RESULTS_DIR = GLOVO_PROJECT_DIR / "RESULTS"
DEFAULT_STATE_DIR = GLOVO_PROJECT_DIR / "state"
DEFAULT_CATALOG_PATH = DEFAULT_STATE_DIR / "glovo_almaty_pharmacies_catalog.csv"
DEFAULT_WOLT_CATALOG_PATH = BASE_DIR / "wolt_project" / "state" / "wolt_pharmacies_catalog.csv"
DEFAULT_GRID_ROWS = 5
DEFAULT_GRID_COLS = 6
DEFAULT_GRID_PADDING = 0.006
DEFAULT_COMPARE_DISTANCE_M = 250.0
DEFAULT_SLEEP_MS = 180
DEFAULT_TIMEOUT = 30.0
DEFAULT_MANUAL_ADDRESSES = [
    "Достык 107, Алматы",
    "Тимирязева 34, Алматы",
    "Жандосова 108, Алматы",
    "Алтынсарина 26, Алматы",
    "Сулейменова 24, Алматы",
    "Сейфуллина 9А, Алматы",
    "Майлина 2, Алматы",
]
CHAIN_ALIASES = {
    "europharma": ("europharma",),
    "dobraya pharmacy": ("добрая аптека", "dobraya pharmacy", "dobraya apteka"),
    "sadykhan": ("садыхан", "sadykhan", "sadyhan"),
    "pharmacom": ("pharmacom",),
    "apteka so sklada": ("аптека со склада", "so sklada", "sklada"),
    "alma pharmacy": ("alma pharmacy", "alma pharm", "алма"),
    "keruen pharma": ("keruen pharma",),
    "melissa pharmacy": ("melissa pharmacy",),
}


def is_bad_store_name(value: str, slug: str = "") -> bool:
    text = str(value or "").strip()
    slug_norm = normalize_text(str(slug or "").replace("-", " "))
    norm = normalize_text(text)
    if not text or text in {"--", "-", "—"}:
        return True
    if norm in {"", slug_norm}:
        return True
    if re.fullmatch(r"\d+\s*%?", text):
        return True
    if len(norm) <= 2 and not any(ch.isalpha() for ch in text):
        return True
    return False


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Discover Glovo pharmacies using address grid + delivery context cookies."
    )
    parser.add_argument("--category-url", default=DEFAULT_CATEGORY_URL, help="Primary Glovo category URL")
    parser.add_argument(
        "--source-url",
        action="append",
        default=[],
        help="Extra SSR discovery source URL. Can be passed multiple times.",
    )
    parser.add_argument(
        "--manual-address",
        action="append",
        default=[],
        help="Manual address probe. Can be passed multiple times.",
    )
    parser.add_argument("--results-dir", default=str(DEFAULT_RESULTS_DIR), help="Folder for output discovery files")
    parser.add_argument("--catalog-path", default=str(DEFAULT_CATALOG_PATH), help="Cumulative pharmacy catalog CSV")
    parser.add_argument(
        "--wolt-catalog-path",
        default=str(DEFAULT_WOLT_CATALOG_PATH),
        help="Wolt pharmacy catalog for grid bounds and overlap comparison",
    )
    parser.add_argument("--grid-rows", type=int, default=DEFAULT_GRID_ROWS, help="Number of latitude grid rows")
    parser.add_argument("--grid-cols", type=int, default=DEFAULT_GRID_COLS, help="Number of longitude grid cols")
    parser.add_argument(
        "--grid-padding",
        type=float,
        default=DEFAULT_GRID_PADDING,
        help="Extra padding added to Wolt bbox in degrees",
    )
    parser.add_argument(
        "--sleep-ms",
        type=int,
        default=DEFAULT_SLEEP_MS,
        help="Pause between address probes / store fetches",
    )
    parser.add_argument("--timeout", type=float, default=DEFAULT_TIMEOUT, help="HTTP timeout in seconds")
    parser.add_argument(
        "--compare-distance-m",
        type=float,
        default=DEFAULT_COMPARE_DISTANCE_M,
        help="Max distance for Glovo/Wolt overlap match",
    )
    parser.add_argument("--city-slug", default=DEFAULT_CITY_SLUG, help="City slug used in URLs")
    parser.add_argument("--language", default=DEFAULT_LANGUAGE, help="Language code used in URLs")
    parser.add_argument("--country", default=DEFAULT_COUNTRY, help="Country code used in URLs")
    return parser.parse_args()


def resolve_source_urls(category_url: str, extra_urls: list[str]) -> list[str]:
    raw_urls = [category_url, *extra_urls, *DEFAULT_SOURCE_URLS]
    seen: set[str] = set()
    urls: list[str] = []
    for raw in raw_urls:
        url = (raw or "").strip()
        if not url or url in seen:
            continue
        seen.add(url)
        urls.append(url)
    return urls


def normalize_address_text(value: str) -> str:
    text = normalize_text(value)
    replacements = {
        "проспект": "",
        "prospect": "",
        "avenue": "",
        "ave": "",
        "street": "",
        "st": "",
        "улица": "",
        "ulica": "",
        "ул ": "",
        " микр": " ",
        "microdistrict": "",
        "district": "",
        "район": "",
        "kazakhstan": "",
        "алматы": "",
        "almaty": "",
        "казахстан": "",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = text.replace("ё", "е")
    text = " ".join(text.split())
    return text


def slug_to_store_url(slug: str, *, language: str, country: str, city_slug: str) -> str:
    return f"https://glovoapp.com/{language}/{country}/{city_slug}/stores/{slug}"


def extract_store_card_title(html_text: str, slug: str) -> str:
    if not html_text or not slug:
        return ""
    for pattern in (f'/stores/{slug}"', f"/stores/{slug}"):
        index = html_text.find(pattern)
        if index < 0:
            continue
        snippet = html_text[max(0, index - 500): index + 2500]
        for title_pattern in (
            r'alt="([^"]+)"',
            r"<p[^>]*StoreCardStoreWall_title[^>]*>([^<]+)</p>",
        ):
            match = re.search(title_pattern, snippet, re.IGNORECASE)
            if match:
                return html.unescape(match.group(1)).strip()
    return ""


def category_request_headers(session: requests.Session) -> dict[str, str]:
    ua = session.headers.get("User-Agent") or "Mozilla/5.0"
    return {
        "User-Agent": ua,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
    }


def api_request_headers(session: requests.Session) -> dict[str, str]:
    ua = session.headers.get("User-Agent") or "Mozilla/5.0"
    return {"User-Agent": ua, "Accept": "application/json"}


def _pick_best_address_match(payload: dict[str, Any]) -> dict[str, Any] | None:
    addresses = payload.get("addresses") if isinstance(payload, dict) else None
    if not isinstance(addresses, list):
        return None
    for item in addresses:
        if isinstance(item, dict) and item.get("partialMatch") is False:
            return item
    for item in addresses:
        if isinstance(item, dict):
            return item
    return None


def lookup_address(session: requests.Session, query: str, timeout: float) -> dict[str, Any] | None:
    response = session.get(
        "https://api.glovoapp.com/v3/addresslookup/pub/address",
        params={"address": query},
        headers=api_request_headers(session),
        timeout=timeout,
    )
    response.raise_for_status()
    return _pick_best_address_match(response.json())


def reverse_lookup_coordinates(
    session: requests.Session,
    *,
    latitude: float,
    longitude: float,
    timeout: float,
) -> dict[str, Any] | None:
    response = session.get(
        "https://api.glovoapp.com/v3/addresslookup/pub/coordinates",
        params={"latitude": latitude, "longitude": longitude},
        headers=api_request_headers(session),
        timeout=timeout,
    )
    response.raise_for_status()
    payload = response.json()
    return payload if isinstance(payload, dict) and payload.get("placeId") else None


def build_delivery_address_payload(place: dict[str, Any]) -> dict[str, Any]:
    components = place.get("addressComponents") or {}
    locality = components.get("locality") or components.get("administrative_area_level_2") or "Almaty"
    postal_code = components.get("postal_code") or ""
    return {
        "geo": {"lat": place.get("latitude"), "lng": place.get("longitude")},
        "city": {
            "code": place.get("cityCode") or "ALA",
            "name": locality,
            "countryCode": place.get("countryCode") or "KZ",
        },
        "placeId": place.get("placeId") or "",
        "text": place.get("title") or "",
        "details": place.get("subtitle") or "",
        "postalCode": postal_code,
        "isVerified": True,
    }


def build_delivery_cookies(place: dict[str, Any]) -> dict[str, str]:
    payload = build_delivery_address_payload(place)
    timestamp = str(int(time.time() * 1000))
    return {
        "Glovo-Location-Country-Code": str(place.get("countryCode") or "KZ"),
        "Glovo-Location-City-Code": str(place.get("cityCode") or "ALA"),
        "Glovo-Delivery-Location-Longitude": str(place.get("longitude") or ""),
        "Glovo-Delivery-Location-Latitude": str(place.get("latitude") or ""),
        "Glovo-Delivery-Location-Timestamp": timestamp,
        "Glovo-Delivery-Location-Accuracy": "0",
        "glovo_delivery_address": requests.utils.quote(json.dumps(payload, ensure_ascii=False)),
    }


def contextual_session_for_place(session: requests.Session, place: dict[str, Any]) -> requests.Session:
    probe = build_session()
    probe.headers.update({"Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7"})
    for key, value in build_delivery_cookies(place).items():
        probe.cookies.set(key, value, domain="glovoapp.com", path="/")
    return probe


def fetch_category_html_for_place(
    session: requests.Session,
    *,
    category_url: str,
    place: dict[str, Any],
    timeout: float,
    attempts: int = 4,
) -> str:
    cookies = build_delivery_cookies(place)
    headers = category_request_headers(session)
    last_text = ""
    last_exc: Exception | None = None
    for attempt in range(1, attempts + 1):
        probe = build_session()
        probe.headers.update({"Accept": headers["Accept"]})
        response = None
        try:
            for key, value in cookies.items():
                probe.cookies.set(key, value, domain="glovoapp.com", path="/")
            response = probe.get(category_url, headers=headers, timeout=timeout)
            response.raise_for_status()
            text = response.text
            last_text = text
            if "Oh, no!" in text or "problem, but we're working on it" in text.lower():
                raise RuntimeError("Glovo returned fallback error page")
            if extract_store_urls(text):
                return text
            if attempt == attempts:
                return text
        except Exception as exc:  # noqa: BLE001
            last_exc = exc
            if attempt == attempts:
                break
            time.sleep(0.8 * attempt)
            continue
        finally:
            if response is not None:
                response.close()
            probe.close()
    if last_text:
        return last_text
    if last_exc is not None:
        raise last_exc
    raise RuntimeError("Unable to fetch Glovo category page")


def fetch_store_page_html(
    session: requests.Session,
    *,
    store_url: str,
    timeout: float,
    context_address: str = "",
) -> str:
    headers = category_request_headers(session)
    context_query = (context_address or "").strip()
    if not context_query:
        response = session.get(store_url, headers=headers, timeout=timeout)
        response.raise_for_status()
        return response.text

    place = lookup_address(session, context_query, timeout)
    if not place:
        response = session.get(store_url, headers=headers, timeout=timeout)
        response.raise_for_status()
        return response.text
    coords_place = reverse_lookup_coordinates(
        session,
        latitude=float(place["latitude"]),
        longitude=float(place["longitude"]),
        timeout=timeout,
    ) or place
    probe = contextual_session_for_place(session, coords_place)
    try:
        response = probe.get(store_url, headers=headers, timeout=timeout)
        response.raise_for_status()
        return response.text
    finally:
        probe.close()


def fetch_store_page_html_for_place(
    session: requests.Session,
    *,
    store_url: str,
    place: dict[str, Any],
    timeout: float,
) -> str:
    headers = category_request_headers(session)
    probe = contextual_session_for_place(session, place)
    try:
        response = probe.get(store_url, headers=headers, timeout=timeout)
        response.raise_for_status()
        return response.text
    finally:
        probe.close()


def load_wolt_locations(path: Path) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if not path.exists():
        return rows
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            try:
                lat = float(str(row.get("lat") or "").strip())
                lon = float(str(row.get("lon") or "").strip())
            except Exception:
                continue
            rows.append(
                {
                    "name": str(row.get("name") or "").strip(),
                    "slug": str(row.get("slug") or "").strip(),
                    "address": str(row.get("address") or "").strip(),
                    "lat": lat,
                    "lon": lon,
                }
            )
    return rows


def build_grid_points(
    wolt_rows: list[dict[str, Any]],
    *,
    rows: int,
    cols: int,
    padding: float,
) -> list[dict[str, Any]]:
    if not wolt_rows or rows <= 0 or cols <= 0:
        return []
    min_lat = min(item["lat"] for item in wolt_rows) - padding
    max_lat = max(item["lat"] for item in wolt_rows) + padding
    min_lon = min(item["lon"] for item in wolt_rows) - padding
    max_lon = max(item["lon"] for item in wolt_rows) + padding

    lat_step = (max_lat - min_lat) / rows
    lon_step = (max_lon - min_lon) / cols
    points: list[dict[str, Any]] = []
    for row_idx in range(rows):
        for col_idx in range(cols):
            lat = min_lat + lat_step * (row_idx + 0.5)
            lon = min_lon + lon_step * (col_idx + 0.5)
            points.append(
                {
                    "probe_type": "grid",
                    "probe_label": f"grid-r{row_idx + 1:02d}-c{col_idx + 1:02d}",
                    "input_address": "",
                    "latitude": lat,
                    "longitude": lon,
                }
            )
    return points


def build_manual_points(addresses: list[str]) -> list[dict[str, Any]]:
    points: list[dict[str, Any]] = []
    for idx, address in enumerate(addresses, start=1):
        query = (address or "").strip()
        if not query:
            continue
        points.append(
            {
                "probe_type": "manual",
                "probe_label": f"manual-{idx:02d}",
                "input_address": query,
                "latitude": None,
                "longitude": None,
            }
        )
    return points


def resolve_probe_place(session: requests.Session, point: dict[str, Any], timeout: float) -> dict[str, Any] | None:
    if point.get("input_address"):
        address_hit = lookup_address(session, str(point["input_address"]), timeout)
        if not address_hit:
            return None
        coords_hit = reverse_lookup_coordinates(
            session,
            latitude=float(address_hit["latitude"]),
            longitude=float(address_hit["longitude"]),
            timeout=timeout,
        )
        return coords_hit or address_hit

    lat = point.get("latitude")
    lon = point.get("longitude")
    if lat is None or lon is None:
        return None
    return reverse_lookup_coordinates(session, latitude=float(lat), longitude=float(lon), timeout=timeout)


def discover_stores_from_sources(
    session: requests.Session,
    *,
    source_urls: list[str],
    timeout: float,
) -> dict[str, dict[str, Any]]:
    store_map: dict[str, dict[str, Any]] = {}
    for source_url in source_urls:
        try:
            response = session.get(source_url, timeout=timeout)
            response.raise_for_status()
        except Exception as exc:  # noqa: BLE001
            print(f"[source] {source_url} -> ERROR: {exc}")
            continue
        store_urls = extract_store_urls(response.text)
        print(f"[source] {source_url} -> stores={len(store_urls)}")
        for store_url in store_urls:
            slug = store_url.rstrip("/").rsplit("/", 1)[-1]
            if not slug:
                continue
            entry = store_map.setdefault(
                slug,
                {
                    "store_url": store_url,
                    "display_names": set(),
                    "source_urls": set(),
                    "probe_hits": 0,
                    "probe_labels": set(),
                    "sample_contexts": set(),
                    "sample_probe_addresses": set(),
                    "methods": set(),
                },
            )
            title = extract_store_card_title(response.text, slug)
            if title:
                entry["display_names"].add(title)
            entry["source_urls"].add(source_url)
            entry["methods"].add("source")
    return store_map


def discover_stores_from_probes(
    session: requests.Session,
    *,
    category_url: str,
    points: list[dict[str, Any]],
    timeout: float,
    sleep_ms: int,
    language: str,
    country: str,
    city_slug: str,
) -> tuple[dict[str, dict[str, Any]], list[dict[str, Any]]]:
    store_map: dict[str, dict[str, Any]] = {}
    probe_rows: list[dict[str, Any]] = []
    for index, point in enumerate(points, start=1):
        try:
            place = resolve_probe_place(session, point, timeout)
        except Exception as exc:  # noqa: BLE001
            probe_rows.append(
                {
                    "probe_label": point.get("probe_label") or "",
                    "probe_type": point.get("probe_type") or "",
                    "input_address": point.get("input_address") or "",
                    "resolved_title": "",
                    "resolved_subtitle": "",
                    "resolved_full_address": "",
                    "place_id": "",
                    "latitude": point.get("latitude") or "",
                    "longitude": point.get("longitude") or "",
                    "store_count": 0,
                    "store_slugs": "",
                    "status": "resolve_failed",
                    "error": str(exc),
                }
            )
            continue
        if not place:
            probe_rows.append(
                {
                    "probe_label": point.get("probe_label") or "",
                    "probe_type": point.get("probe_type") or "",
                    "input_address": point.get("input_address") or "",
                    "resolved_title": "",
                    "resolved_subtitle": "",
                    "resolved_full_address": "",
                    "place_id": "",
                    "latitude": point.get("latitude") or "",
                    "longitude": point.get("longitude") or "",
                    "store_count": 0,
                    "store_slugs": "",
                    "status": "resolve_failed",
                    "error": "",
                }
            )
            continue
        try:
            html = fetch_category_html_for_place(session, category_url=category_url, place=place, timeout=timeout)
        except Exception as exc:  # noqa: BLE001
            probe_rows.append(
                {
                    "probe_label": point.get("probe_label") or "",
                    "probe_type": point.get("probe_type") or "",
                    "input_address": point.get("input_address") or "",
                    "resolved_title": place.get("title") or "",
                    "resolved_subtitle": place.get("subtitle") or "",
                    "resolved_full_address": place.get("fullAddress") or "",
                    "place_id": place.get("placeId") or "",
                    "latitude": place.get("latitude") or point.get("latitude") or "",
                    "longitude": place.get("longitude") or point.get("longitude") or "",
                    "store_count": 0,
                    "store_slugs": "",
                    "status": "fetch_failed",
                    "error": str(exc),
                }
            )
            print(
                f"[probe {index}/{len(points)}] {point.get('probe_label')} | "
                f"{place.get('title') or point.get('input_address') or ''} -> ERROR: {exc}"
            )
            if sleep_ms > 0:
                time.sleep(sleep_ms / 1000.0)
            continue
        slugs: list[str] = []
        seen: set[str] = set()
        for store_url in extract_store_urls(html):
            slug = store_url.rstrip("/").rsplit("/", 1)[-1]
            if not slug or slug in seen:
                continue
            seen.add(slug)
            slugs.append(slug)
            entry = store_map.setdefault(
                slug,
                {
                    "store_url": slug_to_store_url(slug, language=language, country=country, city_slug=city_slug),
                    "display_names": set(),
                    "source_urls": set(),
                    "probe_hits": 0,
                    "probe_labels": set(),
                    "sample_contexts": set(),
                    "sample_probe_addresses": set(),
                    "methods": set(),
                },
            )
            title = extract_store_card_title(html, slug)
            if title:
                entry["display_names"].add(title)
            entry["probe_hits"] += 1
            entry["probe_labels"].add(str(point.get("probe_label") or ""))
            entry["sample_contexts"].add(str(place.get("title") or ""))
            entry["sample_probe_addresses"].add(str(place.get("fullAddress") or place.get("title") or ""))
            entry["methods"].add("grid" if point.get("probe_type") == "grid" else "manual")

        probe_rows.append(
            {
                "probe_label": point.get("probe_label") or "",
                "probe_type": point.get("probe_type") or "",
                "input_address": point.get("input_address") or "",
                "resolved_title": place.get("title") or "",
                "resolved_subtitle": place.get("subtitle") or "",
                "resolved_full_address": place.get("fullAddress") or "",
                "place_id": place.get("placeId") or "",
                "latitude": place.get("latitude") or point.get("latitude") or "",
                "longitude": place.get("longitude") or point.get("longitude") or "",
                "store_count": len(slugs),
                "store_slugs": " || ".join(slugs),
                "status": "ok" if slugs else "empty",
                "error": "",
            }
        )
        print(
            f"[probe {index}/{len(points)}] {point.get('probe_label')} | "
            f"{place.get('title') or point.get('input_address') or ''} -> stores={len(slugs)}"
        )
        if sleep_ms > 0:
            time.sleep(sleep_ms / 1000.0)
    return store_map, probe_rows


def geocode_store_address(session: requests.Session, address: str, timeout: float) -> dict[str, Any]:
    raw = (address or "").strip()
    if not raw:
        return {"store_geo_lat": "", "store_geo_lon": "", "store_geo_place_id": "", "store_geo_address": ""}
    candidates = [raw]
    cleaned = (
        raw.replace("\u200b", " ")
        .replace("​", " ")
        .replace("Алматинская область", "Алматы")
        .replace("Kazakhstan", "")
        .replace("Казахстан", "")
        .replace("050040", "")
        .replace("050000", "")
    )
    cleaned = cleaned.replace(" ,", ",").replace(" , ", ", ")
    cleaned = " ".join(cleaned.split())
    if cleaned and cleaned not in candidates:
        candidates.append(cleaned)
    first_part = cleaned.split(",", 1)[0].strip()
    if first_part:
        short_query = f"{first_part}, Алматы"
        if short_query not in candidates:
            candidates.append(short_query)
    if "," in cleaned:
        first_two = ", ".join(part.strip() for part in cleaned.split(",")[:2] if part.strip())
        if first_two and first_two not in candidates:
            candidates.append(first_two)
    hit = None
    for candidate in candidates:
        hit = lookup_address(session, candidate, timeout)
        if hit:
            break
    if not hit:
        return {"store_geo_lat": "", "store_geo_lon": "", "store_geo_place_id": "", "store_geo_address": ""}
    return {
        "store_geo_lat": hit.get("latitude") or "",
        "store_geo_lon": hit.get("longitude") or "",
        "store_geo_place_id": hit.get("placeId") or "",
        "store_geo_address": hit.get("fullAddress") or hit.get("title") or raw,
    }


def canonical_chain_name(value: str) -> str:
    text = normalize_text(value).replace("-", " ")
    for chain, aliases in CHAIN_ALIASES.items():
        if any(alias in text for alias in aliases):
            return chain
    return text.split(" №", 1)[0].split(" no ", 1)[0].strip()


def haversine_m(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    radius = 6371000.0
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    d_phi = math.radians(lat2 - lat1)
    d_lambda = math.radians(lon2 - lon1)
    a = math.sin(d_phi / 2.0) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(d_lambda / 2.0) ** 2
    return 2.0 * radius * math.atan2(math.sqrt(a), math.sqrt(1.0 - a))


def compare_with_wolt(
    glovo_rows: list[dict[str, Any]],
    wolt_rows: list[dict[str, Any]],
    *,
    max_distance_m: float,
) -> list[dict[str, Any]]:
    comparison_rows: list[dict[str, Any]] = []
    for row in glovo_rows:
        if normalize_text(str(row.get("is_pharmacy") or "")) in {"0", "false", "no"}:
            continue
        lat_raw = row.get("store_geo_lat")
        lon_raw = row.get("store_geo_lon")
        try:
            glovo_lat = float(str(lat_raw))
            glovo_lon = float(str(lon_raw))
        except Exception:
            glovo_lat = None
            glovo_lon = None

        chain = canonical_chain_name(str(row.get("name") or ""))
        best_same_chain: tuple[float, dict[str, Any]] | None = None
        best_any: tuple[float, dict[str, Any]] | None = None
        for wolt in wolt_rows:
            distance = None
            if glovo_lat is not None and glovo_lon is not None:
                distance = haversine_m(glovo_lat, glovo_lon, float(wolt["lat"]), float(wolt["lon"]))
                if best_any is None or distance < best_any[0]:
                    best_any = (distance, wolt)
            if canonical_chain_name(str(wolt.get("name") or "")) != chain:
                continue
            if distance is None:
                continue
            if best_same_chain is None or distance < best_same_chain[0]:
                best_same_chain = (distance, wolt)

        match_status = "glovo_only"
        matched = None
        match_distance = None
        if best_same_chain and best_same_chain[0] <= max_distance_m:
            match_status = "same_chain_overlap"
            match_distance, matched = best_same_chain
        elif best_same_chain:
            match_status = "same_chain_far"
            match_distance, matched = best_same_chain
        elif best_any and best_any[0] <= max_distance_m:
            match_status = "nearby_other_chain"
            match_distance, matched = best_any
        elif best_any:
            match_status = "nearest_other_chain_far"
            match_distance, matched = best_any

        comparison_rows.append(
            {
                "glovo_name": row.get("name") or "",
                "glovo_slug": row.get("slug") or "",
                "glovo_address": row.get("address") or "",
                "glovo_chain": chain,
                "glovo_store_url": row.get("store_url") or "",
                "glovo_lat": row.get("store_geo_lat") or "",
                "glovo_lon": row.get("store_geo_lon") or "",
                "match_status": match_status,
                "distance_m": f"{match_distance:.1f}" if isinstance(match_distance, (int, float)) else "",
                "wolt_name": matched.get("name") if matched else "",
                "wolt_slug": matched.get("slug") if matched else "",
                "wolt_address": matched.get("address") if matched else "",
                "wolt_lat": matched.get("lat") if matched else "",
                "wolt_lon": matched.get("lon") if matched else "",
            }
        )
    comparison_rows.sort(key=lambda item: (item["match_status"], item["glovo_chain"], item["glovo_name"]))
    return comparison_rows


def summarize_store_map(store_map: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for slug, meta in sorted(store_map.items()):
        rows.append(
            {
                "slug": slug,
                "store_url": meta.get("store_url") or "",
                "display_names": " || ".join(sorted(meta.get("display_names") or [])),
                "discovery_methods": " || ".join(sorted(meta.get("methods") or [])),
                "grid_probe_hits": int(meta.get("probe_hits") or 0),
                "source_urls": " || ".join(sorted(meta.get("source_urls") or [])),
                "probe_labels": " || ".join(sorted(meta.get("probe_labels") or [])),
                "sample_contexts": " || ".join(sorted(meta.get("sample_contexts") or [])),
                "sample_probe_addresses": " || ".join(sorted(meta.get("sample_probe_addresses") or [])),
            }
        )
    return rows


def _split_probe_slugs(value: str) -> list[str]:
    raw = str(value or "").strip()
    if not raw:
        return []
    return [slug.strip() for slug in raw.split(" || ") if slug.strip()]


def _branch_key(*, store_id: str, address_id: str, address: str, slug: str) -> str:
    if store_id and address_id:
        return f"{store_id}:{address_id}"
    address_norm = normalize_address_text(address)
    if address_norm:
        return f"addr:{address_norm}"
    return f"slug:{slug}"


def resolve_branch_contexts(
    session: requests.Session,
    *,
    store_map: dict[str, dict[str, Any]],
    probe_rows: list[dict[str, Any]],
    pharmacy_rows: list[dict[str, Any]],
    timeout: float,
    sleep_ms: int,
) -> list[dict[str, Any]]:
    pharmacy_by_slug = {
        str(row.get("slug") or "").strip().lower(): row
        for row in pharmacy_rows
        if normalize_text(str(row.get("is_pharmacy") or "")) in {"1", "true", "yes"}
    }
    successful_pairs: list[tuple[dict[str, Any], str]] = []
    for probe in probe_rows:
        if str(probe.get("status") or "").strip().lower() != "ok":
            continue
        for slug in _split_probe_slugs(str(probe.get("store_slugs") or "")):
            if slug.lower() not in pharmacy_by_slug:
                continue
            successful_pairs.append((probe, slug.lower()))

    store_info_cache: dict[tuple[str, str], dict[str, Any]] = {}
    geo_cache: dict[str, dict[str, Any]] = {}
    context_rows: list[dict[str, Any]] = []
    total = len(successful_pairs)
    for index, (probe, slug) in enumerate(successful_pairs, start=1):
        base = pharmacy_by_slug.get(slug) or {}
        store_url = str((store_map.get(slug) or {}).get("store_url") or base.get("store_url") or "").strip()
        display_names = sorted((store_map.get(slug) or {}).get("display_names") or [])
        fallback_name = str(base.get("name") or "").strip() or (display_names[0] if display_names else slug)
        place = {
            "placeId": probe.get("place_id") or "",
            "latitude": probe.get("latitude") or "",
            "longitude": probe.get("longitude") or "",
            "title": probe.get("resolved_title") or "",
            "subtitle": probe.get("resolved_subtitle") or "",
            "fullAddress": probe.get("resolved_full_address") or "",
            "countryCode": "KZ",
            "cityCode": "ALA",
        }
        row = {
            "probe_label": probe.get("probe_label") or "",
            "probe_type": probe.get("probe_type") or "",
            "probe_input_address": probe.get("input_address") or "",
            "probe_resolved_title": probe.get("resolved_title") or "",
            "probe_resolved_subtitle": probe.get("resolved_subtitle") or "",
            "probe_resolved_full_address": probe.get("resolved_full_address") or "",
            "probe_place_id": probe.get("place_id") or "",
            "probe_latitude": probe.get("latitude") or "",
            "probe_longitude": probe.get("longitude") or "",
            "slug": slug,
            "store_url": store_url,
            "display_name": fallback_name,
            "context_store_name": fallback_name,
            "context_store_id": base.get("store_id") or "",
            "context_address_id": base.get("address_id") or "",
            "context_address": base.get("address") or "",
            "context_geo_lat": base.get("store_geo_lat") or "",
            "context_geo_lon": base.get("store_geo_lon") or "",
            "context_geo_place_id": base.get("store_geo_place_id") or "",
            "context_geo_address": base.get("store_geo_address") or "",
            "branch_key": "",
            "fetch_status": "error",
            "error": "",
        }
        try:
            html = fetch_store_page_html_for_place(session, store_url=store_url, place=place, timeout=timeout)
            contextual_row = parse_store_page(html, store_url)
            context_name = str(contextual_row.get("name") or "").strip()
            if is_bad_store_name(context_name, slug):
                context_name = fallback_name
            context_store_id = str(contextual_row.get("store_id") or "").strip()
            context_address_id = str(contextual_row.get("address_id") or "").strip()
            info_key = (context_store_id, context_address_id)
            store_info = store_info_cache.get(info_key)
            if store_info is None:
                store_info = fetch_store_info(
                    session,
                    store_id=context_store_id,
                    address_id=context_address_id,
                    timeout=timeout,
                )
                store_info_cache[info_key] = store_info
            context_address = str(store_info.get("address") or "").strip() or str(base.get("address") or "").strip()
            geo_key = context_address
            geo = geo_cache.get(geo_key)
            if geo is None:
                geo = geocode_store_address(session, context_address, timeout)
                geo_cache[geo_key] = geo
            row.update(
                {
                    "display_name": context_name or fallback_name,
                    "context_store_name": context_name or fallback_name,
                    "context_store_id": context_store_id,
                    "context_address_id": context_address_id,
                    "context_address": context_address,
                    "context_geo_lat": geo.get("store_geo_lat") or "",
                    "context_geo_lon": geo.get("store_geo_lon") or "",
                    "context_geo_place_id": geo.get("store_geo_place_id") or "",
                    "context_geo_address": geo.get("store_geo_address") or "",
                    "branch_key": _branch_key(
                        store_id=context_store_id,
                        address_id=context_address_id,
                        address=context_address,
                        slug=slug,
                    ),
                    "fetch_status": "ok",
                    "error": "",
                }
            )
        except Exception as exc:  # noqa: BLE001
            row["branch_key"] = _branch_key(
                store_id=str(row.get("context_store_id") or ""),
                address_id=str(row.get("context_address_id") or ""),
                address=str(row.get("context_address") or ""),
                slug=slug,
            )
            row["error"] = str(exc)
        context_rows.append(row)
        print(
            f"[branch {index}/{total}] {row.get('display_name') or slug} | "
            f"{row.get('probe_label') or ''} -> {row.get('context_address') or row.get('fetch_status')}"
        )
        if sleep_ms > 0:
            time.sleep(sleep_ms / 1000.0)
    return context_rows


def main() -> None:
    try:
        code = run()
    except Exception as exc:  # noqa: BLE001
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(2)
    sys.exit(code)


def run() -> int:
    args = parse_args()
    results_dir = Path(args.results_dir).expanduser().resolve()
    results_dir.mkdir(parents=True, exist_ok=True)
    catalog_path = Path(args.catalog_path).expanduser().resolve()
    catalog_path.parent.mkdir(parents=True, exist_ok=True)
    wolt_catalog_path = Path(args.wolt_catalog_path).expanduser().resolve()
    catalog_snapshot = load_catalog(catalog_path)

    session = build_session()
    source_urls = resolve_source_urls(args.category_url, args.source_url)
    store_map = discover_stores_from_sources(session, source_urls=source_urls, timeout=args.timeout)

    manual_addresses = [*DEFAULT_MANUAL_ADDRESSES, *args.manual_address]
    probe_points = build_manual_points(manual_addresses)
    wolt_rows = load_wolt_locations(wolt_catalog_path)
    probe_points.extend(
        build_grid_points(
            wolt_rows,
            rows=max(0, args.grid_rows),
            cols=max(0, args.grid_cols),
            padding=max(0.0, args.grid_padding),
        )
    )
    probe_map, probe_rows = discover_stores_from_probes(
        session,
        category_url=args.category_url,
        points=probe_points,
        timeout=args.timeout,
        sleep_ms=max(0, args.sleep_ms),
        language=args.language,
        country=args.country,
        city_slug=args.city_slug,
    )

    for slug, meta in probe_map.items():
        existing = store_map.setdefault(
            slug,
            {
                "store_url": meta.get("store_url") or slug_to_store_url(slug, language=args.language, country=args.country, city_slug=args.city_slug),
                "display_names": set(),
                "source_urls": set(),
                "probe_hits": 0,
                "probe_labels": set(),
                "sample_contexts": set(),
                "sample_probe_addresses": set(),
                "methods": set(),
            },
        )
        existing["store_url"] = meta.get("store_url") or existing.get("store_url") or ""
        existing["display_names"].update(meta.get("display_names") or set())
        existing["source_urls"].update(meta.get("source_urls") or set())
        existing["probe_hits"] = int(existing.get("probe_hits") or 0) + int(meta.get("probe_hits") or 0)
        existing["probe_labels"].update(meta.get("probe_labels") or set())
        existing["sample_contexts"].update(meta.get("sample_contexts") or set())
        existing["sample_probe_addresses"].update(meta.get("sample_probe_addresses") or set())
        existing["methods"].update(meta.get("methods") or set())

    if not store_map:
        raise RuntimeError("No stores discovered from Glovo sources or address probes")

    store_summary_rows = summarize_store_map(store_map)
    now = datetime.now()
    timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")
    checked_at = now.isoformat(timespec="seconds")

    grid_path = results_dir / f"glovo_almaty_address_grid_{timestamp}.csv"
    write_csv(
        grid_path,
        probe_rows,
        [
            "probe_label",
            "probe_type",
            "input_address",
            "resolved_title",
            "resolved_subtitle",
            "resolved_full_address",
            "place_id",
            "latitude",
            "longitude",
            "store_count",
            "store_slugs",
            "status",
            "error",
        ],
    )

    store_summary_path = results_dir / f"glovo_almaty_store_visibility_{timestamp}.csv"
    write_csv(
        store_summary_path,
        store_summary_rows,
        [
            "slug",
            "store_url",
            "display_names",
            "discovery_methods",
            "grid_probe_hits",
            "source_urls",
            "probe_labels",
            "sample_contexts",
            "sample_probe_addresses",
        ],
    )

    rows: list[dict[str, Any]] = []
    store_urls = [meta.get("store_url") or slug_to_store_url(slug, language=args.language, country=args.country, city_slug=args.city_slug) for slug, meta in sorted(store_map.items())]
    for index, store_url in enumerate(store_urls, start=1):
        slug = store_url.rstrip("/").rsplit("/", 1)[-1]
        meta = store_map.get(slug, {})
        display_names = sorted(meta.get("display_names") or [])
        display_name = display_names[0] if display_names else ""
        catalog_row = catalog_snapshot.get(slug) or {}
        context_address = ""
        raw_context = " || ".join(sorted(meta.get("sample_probe_addresses") or []))
        if raw_context:
            context_address = raw_context.split(" || ", 1)[0].strip()
        try:
            plain_html = fetch_store_page_html(session, store_url=store_url, timeout=args.timeout)
            row = parse_store_page(plain_html, store_url)
            if (
                not row.get("store_id")
                or not row.get("address_id")
                or row.get("name") == slug
                or is_bad_store_name(str(row.get("name") or ""), slug)
            ) and context_address:
                contextual_html = fetch_store_page_html(
                    session,
                    store_url=store_url,
                    timeout=args.timeout,
                    context_address=context_address,
                )
                contextual_row = parse_store_page(contextual_html, store_url)
                if contextual_row.get("name") and not is_bad_store_name(str(contextual_row.get("name") or ""), slug):
                    row["name"] = contextual_row["name"]
                if contextual_row.get("store_id"):
                    row["store_id"] = contextual_row["store_id"]
                if contextual_row.get("address_id"):
                    row["address_id"] = contextual_row["address_id"]
                if not row.get("address") and contextual_row.get("address"):
                    row["address"] = contextual_row["address"]
            if display_name and (
                not row.get("name")
                or row.get("name") == slug
                or is_bad_store_name(str(row.get("name") or ""), slug)
            ):
                row["name"] = display_name
            row.update(fetch_store_info(session, store_id=row.get("store_id", ""), address_id=row.get("address_id", ""), timeout=args.timeout))
            row.update(geocode_store_address(session, row.get("address", ""), args.timeout))
            row["fetch_error"] = ""
        except Exception as exc:  # noqa: BLE001
            row = {
                "name": display_name or str(catalog_row.get("name") or "") or slug,
                "slug": slug,
                "store_url": store_url,
                "store_id": str(catalog_row.get("store_id") or ""),
                "address_id": str(catalog_row.get("address_id") or ""),
                "address": str(catalog_row.get("address") or ""),
                "store_geo_lat": "",
                "store_geo_lon": "",
                "store_geo_place_id": "",
                "store_geo_address": "",
                "fetch_error": str(exc),
            }
        row["is_pharmacy"], row["store_type"], row["skip_reason"] = classify_store(row["name"], row["slug"])
        meta = store_map.get(row["slug"], meta)
        row["category_url"] = args.category_url
        row["discovery_sources"] = " || ".join(sorted(meta.get("source_urls") or []))
        row["discovery_methods"] = " || ".join(sorted(meta.get("methods") or []))
        row["grid_probe_hits"] = int(meta.get("probe_hits") or 0)
        row["probe_labels"] = " || ".join(sorted(meta.get("probe_labels") or []))
        row["sample_contexts"] = " || ".join(sorted(meta.get("sample_contexts") or []))
        row["sample_probe_addresses"] = " || ".join(sorted(meta.get("sample_probe_addresses") or []))
        row["city_slug"] = args.city_slug
        row["language"] = args.language
        row["country"] = args.country
        rows.append(row)
        print(
            f"[store {index}/{len(store_urls)}] {row['name']} | slug={row['slug']} | "
            f"methods={row['discovery_methods']} | grid_hits={row['grid_probe_hits']} | is_pharmacy={row['is_pharmacy']}"
        )
        if args.sleep_ms > 0:
            time.sleep(args.sleep_ms / 1000.0)

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
            "store_geo_lat",
            "store_geo_lon",
            "store_geo_place_id",
            "store_geo_address",
            "fetch_error",
            "is_pharmacy",
            "store_type",
            "skip_reason",
            "category_url",
            "discovery_sources",
            "discovery_methods",
            "grid_probe_hits",
            "probe_labels",
            "sample_contexts",
            "sample_probe_addresses",
            "city_slug",
            "language",
            "country",
        ],
    )
    update_catalog(catalog_path=catalog_path, rows=rows, checked_at=checked_at)

    compare_rows = compare_with_wolt(rows, wolt_rows, max_distance_m=max(0.0, args.compare_distance_m))
    compare_path = results_dir / f"glovo_almaty_vs_wolt_{timestamp}.csv"
    write_csv(
        compare_path,
        compare_rows,
        [
            "glovo_name",
            "glovo_slug",
            "glovo_address",
            "glovo_chain",
            "glovo_store_url",
            "glovo_lat",
            "glovo_lon",
            "match_status",
            "distance_m",
            "wolt_name",
            "wolt_slug",
            "wolt_address",
            "wolt_lat",
            "wolt_lon",
        ],
    )

    branch_context_rows = resolve_branch_contexts(
        session,
        store_map=store_map,
        probe_rows=probe_rows,
        pharmacy_rows=rows,
        timeout=args.timeout,
        sleep_ms=max(0, args.sleep_ms),
    )
    branch_contexts_path = results_dir / f"glovo_almaty_branch_contexts_{timestamp}.csv"
    write_csv(
        branch_contexts_path,
        branch_context_rows,
        [
            "probe_label",
            "probe_type",
            "probe_input_address",
            "probe_resolved_title",
            "probe_resolved_subtitle",
            "probe_resolved_full_address",
            "probe_place_id",
            "probe_latitude",
            "probe_longitude",
            "slug",
            "store_url",
            "display_name",
            "context_store_name",
            "context_store_id",
            "context_address_id",
            "context_address",
            "context_geo_lat",
            "context_geo_lon",
            "context_geo_place_id",
            "context_geo_address",
            "branch_key",
            "fetch_status",
            "error",
        ],
    )

    status_counts: dict[str, int] = {}
    for row in compare_rows:
        key = str(row.get("match_status") or "unknown")
        status_counts[key] = status_counts.get(key, 0) + 1

    summary_path = results_dir / f"glovo_almaty_discovery_summary_{timestamp}.json"
    summary_payload = {
        "generated_at": checked_at,
        "category_url": args.category_url,
        "probe_points": len(probe_points),
        "probe_success": sum(1 for row in probe_rows if row.get("status") == "ok"),
        "unique_storefronts": len(rows),
        "pharmacies": sum(1 for row in rows if normalize_text(str(row.get("is_pharmacy") or "")) in {"1", "true", "yes"}),
        "stores_visible_in_grid": sum(1 for row in rows if int(row.get("grid_probe_hits") or 0) > 0),
        "comparison_status_counts": status_counts,
        "files": {
            "pharmacies": output_path.name,
            "grid": grid_path.name,
            "visibility": store_summary_path.name,
            "compare_wolt": compare_path.name,
            "branch_contexts": branch_contexts_path.name,
        },
    }
    summary_path.write_text(json.dumps(summary_payload, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Pharmacies discovered: {len(rows)}")
    print(f"Pharmacy output: {output_path}")
    print(f"Address grid: {grid_path}")
    print(f"Visibility matrix: {store_summary_path}")
    print(f"Glovo vs Wolt: {compare_path}")
    print(f"Summary: {summary_path}")
    return 0


if __name__ == "__main__":
    main()
