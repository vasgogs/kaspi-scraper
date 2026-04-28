#!/usr/bin/env python3
"""Discover Wolt pharmacies in Almaty using a geographic grid."""

from __future__ import annotations

import argparse
import csv
import json
import math
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


BASE_DIR = Path(__file__).resolve().parent
WOLT_PROJECT_DIR = BASE_DIR / "wolt_project"
DEFAULT_RESULTS_DIR = WOLT_PROJECT_DIR / "RESULTS"
DEFAULT_PHARMACIES_CATALOG = WOLT_PROJECT_DIR / "state" / "wolt_pharmacies_catalog.csv"
CITY_API = "https://restaurant-api.wolt.com/v1/cities"
DISTRICTS_API_FMT = "https://restaurant-api.wolt.com/v1/cities/{city_id}/districts"
RETAIL_API = "https://restaurant-api.wolt.com/v1/pages/retail"

PHARMACY_NAME_MARKERS = (
    "pharmacy",
    "apteka",
    "аптека",
    "аптек",
    "apothecary",
    "pharma",
    "фарма",
)


@dataclass
class GridPoint:
    point_id: str
    lat: float
    lon: float
    source: str  # grid or district


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Find all Wolt pharmacies in a city by scanning a coordinate grid."
    )
    parser.add_argument("--city-slug", default="almaty", help="City slug in Wolt")
    parser.add_argument(
        "--country-alpha2",
        default="KZ",
        help="Country alpha2 code filter for city lookup (default: %(default)s)",
    )
    parser.add_argument(
        "--language",
        default="en",
        help="Language parameter for retail page request (default: %(default)s)",
    )
    parser.add_argument(
        "--step-km",
        type=float,
        default=2.5,
        help="Grid step in km (default: %(default)s)",
    )
    parser.add_argument(
        "--padding-km",
        type=float,
        default=1.5,
        help="Padding around district bbox in km (default: %(default)s)",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=20.0,
        help="HTTP timeout in seconds (default: %(default)s)",
    )
    parser.add_argument(
        "--sleep-ms",
        type=int,
        default=120,
        help="Pause between point requests in ms (default: %(default)s)",
    )
    parser.add_argument(
        "--results-dir",
        default=str(DEFAULT_RESULTS_DIR),
        help="Directory for output files (default: %(default)s)",
    )
    parser.add_argument(
        "--output-prefix",
        default="wolt_almaty_pharmacies",
        help="Prefix for output files (default: %(default)s)",
    )
    parser.add_argument(
        "--pharmacies-catalog",
        default=str(DEFAULT_PHARMACIES_CATALOG),
        help="Cumulative pharmacies catalog CSV path (default: %(default)s)",
    )
    parser.add_argument(
        "--limit-points",
        type=int,
        default=0,
        help="Optional debug limit of points to scan (0 = all)",
    )
    return parser.parse_args()


def build_session() -> requests.Session:
    retry = Retry(
        total=3,
        connect=3,
        read=3,
        backoff_factor=0.4,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
    )
    adapter = HTTPAdapter(max_retries=retry)
    session = requests.Session()
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
            ),
            "Accept": "application/json",
        }
    )
    return session


def fetch_city(session: requests.Session, city_slug: str, country_alpha2: str, timeout: float) -> dict[str, Any]:
    response = session.get(CITY_API, timeout=timeout)
    response.raise_for_status()
    payload = response.json()
    results = payload.get("results") if isinstance(payload, dict) else None
    if not isinstance(results, list):
        raise ValueError("Unexpected cities payload")

    city_slug = city_slug.strip().lower()
    country_alpha2 = country_alpha2.strip().upper()
    for city in results:
        if not isinstance(city, dict):
            continue
        if str(city.get("slug", "")).lower() != city_slug:
            continue
        if country_alpha2 and str(city.get("country_code_alpha2", "")).upper() != country_alpha2:
            continue
        return city
    raise ValueError(f"City not found: slug={city_slug}, country={country_alpha2}")


def fetch_district_points(session: requests.Session, city_id: str, timeout: float) -> list[GridPoint]:
    url = DISTRICTS_API_FMT.format(city_id=city_id)
    response = session.get(url, timeout=timeout)
    response.raise_for_status()
    payload = response.json()
    districts = payload.get("city_districts") if isinstance(payload, dict) else None
    if not isinstance(districts, list):
        return []

    points: list[GridPoint] = []
    for idx, district in enumerate(districts, start=1):
        if not isinstance(district, dict):
            continue
        loc = district.get("location")
        if not (isinstance(loc, list) and len(loc) >= 2):
            continue
        lon, lat = loc[0], loc[1]
        if not isinstance(lat, (int, float)) or not isinstance(lon, (int, float)):
            continue
        points.append(
            GridPoint(
                point_id=f"D{idx:02d}",
                lat=float(lat),
                lon=float(lon),
                source="district",
            )
        )
    return points


def km_to_lat_deg(km: float) -> float:
    return km / 110.574


def km_to_lon_deg(km: float, lat: float) -> float:
    cos_lat = max(0.15, abs(math.cos(math.radians(lat))))
    return km / (111.320 * cos_lat)


def generate_grid_points(
    min_lat: float,
    max_lat: float,
    min_lon: float,
    max_lon: float,
    step_km: float,
) -> list[GridPoint]:
    center_lat = (min_lat + max_lat) / 2.0
    lat_step = max(0.003, km_to_lat_deg(step_km))
    lon_step = max(0.003, km_to_lon_deg(step_km, center_lat))

    points: list[GridPoint] = []
    row = 0
    lat = min_lat
    while lat <= max_lat + 1e-9:
        row += 1
        col = 0
        lon = min_lon
        while lon <= max_lon + 1e-9:
            col += 1
            points.append(
                GridPoint(
                    point_id=f"G{row:02d}_{col:02d}",
                    lat=round(lat, 6),
                    lon=round(lon, 6),
                    source="grid",
                )
            )
            lon += lon_step
        lat += lat_step
    return points


def is_pharmacy_venue(venue: dict[str, Any]) -> tuple[bool, str]:
    tags = venue.get("tags")
    tags_lower = [str(tag).strip().lower() for tag in tags] if isinstance(tags, list) else []
    if "pharmacy" in tags_lower:
        return True, "tag:pharmacy"

    name = str(venue.get("name") or "").lower()
    slug = str(venue.get("slug") or "").lower()
    short_description = str(venue.get("short_description") or "").lower()
    haystack = " ".join([name, slug, short_description])
    for marker in PHARMACY_NAME_MARKERS:
        if marker in haystack:
            return True, f"name:{marker}"
    return False, ""


def parse_retail_response(payload: dict[str, Any]) -> list[dict[str, Any]]:
    sections = payload.get("sections")
    if not isinstance(sections, list):
        return []
    venues: list[dict[str, Any]] = []
    for section in sections:
        if not isinstance(section, dict):
            continue
        items = section.get("items")
        if not isinstance(items, list):
            continue
        for item in items:
            if not isinstance(item, dict):
                continue
            venue = item.get("venue")
            if isinstance(venue, dict):
                venues.append(venue)
    return venues


def fetch_retail_page(
    session: requests.Session,
    lat: float,
    lon: float,
    language: str,
    timeout: float,
    max_attempts: int = 5,
) -> dict[str, Any]:
    last_error: Exception | None = None
    params = {
        "lat": f"{lat:.6f}",
        "lon": f"{lon:.6f}",
        "language": language,
    }
    for attempt in range(1, max_attempts + 1):
        try:
            response = session.get(RETAIL_API, params=params, timeout=timeout)
            if response.status_code == 429 and attempt < max_attempts:
                time.sleep(min(4.0, 0.7 * attempt))
                continue
            response.raise_for_status()
            payload = response.json()
            if not isinstance(payload, dict):
                raise ValueError("Unexpected retail payload")
            return payload
        except Exception as exc:
            last_error = exc
            if attempt >= max_attempts:
                break
            time.sleep(min(4.0, 0.7 * attempt))
    raise RuntimeError(f"Retail request failed after {max_attempts} attempts: {last_error}")


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({name: row.get(name, "") for name in fieldnames})


def load_pharmacy_catalog(path: Path) -> dict[str, dict[str, Any]]:
    if not path.exists():
        return {}
    catalog: dict[str, dict[str, Any]] = {}
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            slug = (row.get("slug") or "").strip()
            if not slug:
                continue
            seen_runs_raw = str(row.get("seen_runs") or "0").strip() or "0"
            try:
                seen_runs = int(seen_runs_raw)
            except Exception:
                seen_runs = 0
            catalog[slug] = {
                "slug": slug,
                "name": (row.get("name") or "").strip(),
                "address": (row.get("address") or "").strip(),
                "city": (row.get("city") or "").strip(),
                "country": (row.get("country") or "").strip(),
                "lat": (row.get("lat") or "").strip(),
                "lon": (row.get("lon") or "").strip(),
                "tags": (row.get("tags") or "").strip(),
                "venue_url": (row.get("venue_url") or "").strip(),
                "first_seen": (row.get("first_seen") or "").strip(),
                "last_seen": (row.get("last_seen") or "").strip(),
                "seen_runs": seen_runs,
                "last_seen_city_slug": (row.get("last_seen_city_slug") or "").strip(),
                "last_seen_detection_reason": (row.get("last_seen_detection_reason") or "").strip(),
                "last_seen_points_count": (row.get("last_seen_points_count") or "").strip(),
                "last_seen_points": (row.get("last_seen_points") or "").strip(),
                "last_online": (row.get("last_online") or "").strip(),
                "last_rating_score": (row.get("last_rating_score") or "").strip(),
                "last_rating_volume": (row.get("last_rating_volume") or "").strip(),
                "last_delivery_price_int": (row.get("last_delivery_price_int") or "").strip(),
            }
    return catalog


def save_pharmacy_catalog(path: Path, catalog: dict[str, dict[str, Any]]) -> None:
    rows = sorted(
        catalog.values(),
        key=lambda row: (
            str(row.get("city") or "").lower(),
            str(row.get("name") or "").lower(),
            str(row.get("slug") or "").lower(),
        ),
    )
    write_csv(
        path,
        rows,
        [
            "slug",
            "name",
            "address",
            "city",
            "country",
            "lat",
            "lon",
            "tags",
            "venue_url",
            "first_seen",
            "last_seen",
            "seen_runs",
            "last_seen_city_slug",
            "last_seen_detection_reason",
            "last_seen_points_count",
            "last_seen_points",
            "last_online",
            "last_rating_score",
            "last_rating_volume",
            "last_delivery_price_int",
        ],
    )


def update_pharmacy_catalog(
    catalog_path: Path,
    pharmacy_rows: list[dict[str, Any]],
    checked_at: str,
    city_slug: str,
) -> tuple[int, list[dict[str, Any]], int]:
    catalog = load_pharmacy_catalog(catalog_path)
    new_rows: list[dict[str, Any]] = []
    updated_count = 0

    for row in pharmacy_rows:
        slug = str(row.get("slug") or "").strip()
        if not slug:
            continue
        existing = catalog.get(slug)
        if existing is None:
            catalog[slug] = {
                "slug": slug,
                "name": row.get("name", ""),
                "address": row.get("address", ""),
                "city": row.get("city", ""),
                "country": row.get("country", ""),
                "lat": row.get("lat", ""),
                "lon": row.get("lon", ""),
                "tags": row.get("tags", ""),
                "venue_url": row.get("venue_url", ""),
                "first_seen": checked_at,
                "last_seen": checked_at,
                "seen_runs": 1,
                "last_seen_city_slug": city_slug,
                "last_seen_detection_reason": row.get("detection_reason", ""),
                "last_seen_points_count": row.get("seen_points_count", ""),
                "last_seen_points": row.get("seen_points", ""),
                "last_online": row.get("online", ""),
                "last_rating_score": row.get("rating_score", ""),
                "last_rating_volume": row.get("rating_volume", ""),
                "last_delivery_price_int": row.get("delivery_price_int", ""),
            }
            new_rows.append(
                {
                    "slug": slug,
                    "name": row.get("name", ""),
                    "address": row.get("address", ""),
                    "city": row.get("city", ""),
                    "country": row.get("country", ""),
                    "first_seen": checked_at,
                    "venue_url": row.get("venue_url", ""),
                }
            )
            continue

        existing["name"] = row.get("name", existing.get("name", ""))
        existing["address"] = row.get("address", existing.get("address", ""))
        existing["city"] = row.get("city", existing.get("city", ""))
        existing["country"] = row.get("country", existing.get("country", ""))
        existing["lat"] = row.get("lat", existing.get("lat", ""))
        existing["lon"] = row.get("lon", existing.get("lon", ""))
        existing["tags"] = row.get("tags", existing.get("tags", ""))
        existing["venue_url"] = row.get("venue_url", existing.get("venue_url", ""))
        existing["last_seen"] = checked_at
        existing["seen_runs"] = int(existing.get("seen_runs", 0)) + 1
        existing["last_seen_city_slug"] = city_slug
        existing["last_seen_detection_reason"] = row.get("detection_reason", "")
        existing["last_seen_points_count"] = row.get("seen_points_count", "")
        existing["last_seen_points"] = row.get("seen_points", "")
        existing["last_online"] = row.get("online", "")
        existing["last_rating_score"] = row.get("rating_score", "")
        existing["last_rating_volume"] = row.get("rating_volume", "")
        existing["last_delivery_price_int"] = row.get("delivery_price_int", "")
        updated_count += 1

    save_pharmacy_catalog(catalog_path, catalog)
    return len(catalog), new_rows, updated_count


def run() -> int:
    args = parse_args()
    session = build_session()
    checked_at = datetime.now().isoformat(timespec="seconds")

    city = fetch_city(
        session=session,
        city_slug=args.city_slug,
        country_alpha2=args.country_alpha2,
        timeout=args.timeout,
    )
    city_id = str(city.get("id"))
    district_points = fetch_district_points(session=session, city_id=city_id, timeout=args.timeout)
    if not district_points:
        raise ValueError(f"No district points found for city_id={city_id}")

    min_lat = min(point.lat for point in district_points)
    max_lat = max(point.lat for point in district_points)
    min_lon = min(point.lon for point in district_points)
    max_lon = max(point.lon for point in district_points)

    lat_pad = km_to_lat_deg(args.padding_km)
    lon_pad = km_to_lon_deg(args.padding_km, (min_lat + max_lat) / 2.0)
    bbox = {
        "min_lat": min_lat - lat_pad,
        "max_lat": max_lat + lat_pad,
        "min_lon": min_lon - lon_pad,
        "max_lon": max_lon + lon_pad,
    }

    grid_points = generate_grid_points(
        min_lat=bbox["min_lat"],
        max_lat=bbox["max_lat"],
        min_lon=bbox["min_lon"],
        max_lon=bbox["max_lon"],
        step_km=args.step_km,
    )

    all_points: list[GridPoint] = grid_points + district_points
    if args.limit_points > 0:
        all_points = all_points[: args.limit_points]

    unique_pharmacies: dict[str, dict[str, Any]] = {}
    point_rows: list[dict[str, Any]] = []
    start = time.time()

    for idx, point in enumerate(all_points, start=1):
        row = {
            "point_id": point.point_id,
            "source": point.source,
            "lat": point.lat,
            "lon": point.lon,
            "request_ok": False,
            "total_venues": 0,
            "pharmacies_found": 0,
            "error": "",
        }
        try:
            payload = fetch_retail_page(
                session=session,
                lat=point.lat,
                lon=point.lon,
                language=args.language,
                timeout=args.timeout,
            )
            venues = parse_retail_response(payload)
            row["request_ok"] = True
            row["total_venues"] = len(venues)

            point_pharmacy_hits = 0
            for venue in venues:
                is_pharmacy, reason = is_pharmacy_venue(venue)
                if not is_pharmacy:
                    continue
                point_pharmacy_hits += 1

                slug = str(venue.get("slug") or "")
                if not slug:
                    continue
                location = venue.get("location") if isinstance(venue.get("location"), list) else []
                venue_lon = location[0] if len(location) >= 2 else ""
                venue_lat = location[1] if len(location) >= 2 else ""

                item = unique_pharmacies.get(slug)
                if item is None:
                    item = {
                        "slug": slug,
                        "name": venue.get("name", ""),
                        "address": venue.get("address", ""),
                        "city": venue.get("city", ""),
                        "country": venue.get("country", ""),
                        "lat": venue_lat,
                        "lon": venue_lon,
                        "tags": ",".join(str(x) for x in (venue.get("tags") or [])),
                        "rating_score": (venue.get("rating") or {}).get("score", ""),
                        "rating_volume": (venue.get("rating") or {}).get("volume", ""),
                        "delivery_price_int": venue.get("delivery_price_int", ""),
                        "online": venue.get("online", ""),
                        "detection_reason": reason,
                        "seen_points": set(),
                    }
                    unique_pharmacies[slug] = item
                item["seen_points"].add(point.point_id)
                item["online"] = bool(venue.get("online", False))

            row["pharmacies_found"] = point_pharmacy_hits
        except Exception as exc:
            row["error"] = str(exc)
        point_rows.append(row)

        if args.sleep_ms > 0:
            time.sleep(args.sleep_ms / 1000.0)

        if idx % 15 == 0 or idx == len(all_points):
            elapsed = time.time() - start
            print(
                f"[{idx}/{len(all_points)}] unique pharmacies={len(unique_pharmacies)} "
                f"elapsed={elapsed:.1f}s"
            )

    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    results_dir = Path(args.results_dir).expanduser().resolve()
    pharmacies_csv = results_dir / f"{args.output_prefix}_{stamp}.csv"
    points_csv = results_dir / f"{args.output_prefix}_points_{stamp}.csv"
    summary_json = results_dir / f"{args.output_prefix}_summary_{stamp}.json"

    pharmacy_rows: list[dict[str, Any]] = []
    for slug, data in unique_pharmacies.items():
        seen_points = sorted(data["seen_points"])
        pharmacy_rows.append(
            {
                "slug": slug,
                "name": data["name"],
                "address": data["address"],
                "city": data["city"],
                "country": data["country"],
                "lat": data["lat"],
                "lon": data["lon"],
                "tags": data["tags"],
                "rating_score": data["rating_score"],
                "rating_volume": data["rating_volume"],
                "delivery_price_int": data["delivery_price_int"],
                "online": data["online"],
                "detection_reason": data["detection_reason"],
                "seen_points_count": len(seen_points),
                "seen_points": ",".join(seen_points),
                "venue_url": f"https://wolt.com/en/kaz/{args.city_slug}/venue/{slug}",
            }
        )

    pharmacy_rows.sort(key=lambda row: str(row.get("name", "")).lower())

    write_csv(
        pharmacies_csv,
        pharmacy_rows,
        [
            "slug",
            "name",
            "address",
            "city",
            "country",
            "lat",
            "lon",
            "tags",
            "rating_score",
            "rating_volume",
            "delivery_price_int",
            "online",
            "detection_reason",
            "seen_points_count",
            "seen_points",
            "venue_url",
        ],
    )

    write_csv(
        points_csv,
        point_rows,
        [
            "point_id",
            "source",
            "lat",
            "lon",
            "request_ok",
            "total_venues",
            "pharmacies_found",
            "error",
        ],
    )

    catalog_path = Path(args.pharmacies_catalog).expanduser().resolve()
    catalog_total, new_pharmacies_rows, updated_existing = update_pharmacy_catalog(
        catalog_path=catalog_path,
        pharmacy_rows=pharmacy_rows,
        checked_at=checked_at,
        city_slug=args.city_slug,
    )
    new_pharmacies_csv = results_dir / f"{args.output_prefix}_new_{stamp}.csv"
    write_csv(
        new_pharmacies_csv,
        new_pharmacies_rows,
        ["slug", "name", "address", "city", "country", "first_seen", "venue_url"],
    )

    summary = {
        "run_at": checked_at,
        "city_slug": args.city_slug,
        "city_id": city_id,
        "country_alpha2": args.country_alpha2,
        "step_km": args.step_km,
        "padding_km": args.padding_km,
        "language": args.language,
        "grid_points": len(grid_points),
        "district_points": len(district_points),
        "scanned_points": len(all_points),
        "unique_pharmacies": len(pharmacy_rows),
        "new_pharmacies_this_run": len(new_pharmacies_rows),
        "updated_existing_this_run": updated_existing,
        "catalog_total_pharmacies": catalog_total,
        "bbox": bbox,
        "files": {
            "pharmacies_csv": str(pharmacies_csv),
            "points_csv": str(points_csv),
            "summary_json": str(summary_json),
            "new_pharmacies_csv": str(new_pharmacies_csv),
            "catalog_csv": str(catalog_path),
        },
    }
    summary_json.parent.mkdir(parents=True, exist_ok=True)
    with open(summary_json, "w", encoding="utf-8") as fh:
        json.dump(summary, fh, ensure_ascii=False, indent=2)

    print(f"City: {args.city_slug} ({city_id})")
    print(f"Scanned points: {len(all_points)} (grid={len(grid_points)}, district={len(district_points)})")
    print(f"Unique pharmacies found: {len(pharmacy_rows)}")
    print(f"New pharmacies this run: {len(new_pharmacies_rows)}")
    print(f"Cumulative catalog total: {catalog_total}")
    print(f"Pharmacies CSV: {pharmacies_csv}")
    print(f"Points CSV: {points_csv}")
    print(f"New pharmacies CSV: {new_pharmacies_csv}")
    print(f"Catalog CSV: {catalog_path}")
    print(f"Summary JSON: {summary_json}")

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
