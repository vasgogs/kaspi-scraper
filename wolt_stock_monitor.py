#!/usr/bin/env python3
"""Track product stock status in Wolt venues (e.g., pharmacies)."""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


BASE_DIR = Path(__file__).resolve().parent
WOLT_PROJECT_DIR = BASE_DIR / "wolt_project"
DEFAULT_CONFIG_PATH = WOLT_PROJECT_DIR / "config" / "wolt_positions.csv"
DEFAULT_RESULTS_DIR = WOLT_PROJECT_DIR / "RESULTS"
DEFAULT_STATE_PATH = WOLT_PROJECT_DIR / "state" / "wolt_stock_state.json"

SOLD_OUT_MARKERS = (
    "sold out",
    "out of stock",
    "unavailable",
    "нет в наличии",
    "законч",
    "временно недоступ",
)


@dataclass
class Position:
    row_index: int
    position_id: str
    pharmacy: str
    venue_url: str
    venue_slug: str
    search_query: str
    product_name: str
    item_id: str
    strict_name: bool

    @property
    def state_key(self) -> str:
        if self.position_id:
            return self.position_id
        tail = self.item_id or self.product_name or self.search_query
        return f"{self.venue_slug}|{tail}".strip()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Monitor stock of configured positions in Wolt venue pages."
    )
    parser.add_argument(
        "--config",
        default=str(DEFAULT_CONFIG_PATH),
        help="CSV config path (default: %(default)s)",
    )
    parser.add_argument(
        "--language",
        default="ru",
        help="Wolt language for search endpoint, e.g. ru or en (default: %(default)s)",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=25.0,
        help="HTTP timeout in seconds (default: %(default)s)",
    )
    parser.add_argument(
        "--results-dir",
        default=str(DEFAULT_RESULTS_DIR),
        help="Folder for run reports (default: %(default)s)",
    )
    parser.add_argument(
        "--state-file",
        default=str(DEFAULT_STATE_PATH),
        help="JSON file for previous snapshot comparison (default: %(default)s)",
    )
    parser.add_argument(
        "--output-prefix",
        default="wolt_stock",
        help="Prefix for output csv/json files (default: %(default)s)",
    )
    return parser.parse_args()


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "").strip().lower())


def parse_bool(value: str) -> bool:
    return normalize_text(value) in {"1", "true", "yes", "y", "on", "да"}


def extract_venue_slug(venue_url: str) -> str:
    if not venue_url:
        return ""
    parsed = urlparse(venue_url.strip())
    if not parsed.path:
        return venue_url.strip().strip("/")
    parts = [part for part in parsed.path.split("/") if part]
    if "venue" in parts:
        idx = parts.index("venue")
        if idx + 1 < len(parts):
            return parts[idx + 1]
    if parts:
        return parts[-1]
    return ""


def build_session() -> requests.Session:
    retry = Retry(
        total=3,
        connect=3,
        read=3,
        backoff_factor=0.5,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "POST"],
    )
    adapter = HTTPAdapter(max_retries=retry)
    session = requests.Session()
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
            ),
            "Content-Type": "application/json",
            "Accept": "application/json",
        }
    )
    return session


def read_positions(config_path: Path) -> list[Position]:
    if not config_path.exists():
        raise FileNotFoundError(
            f"Config file not found: {config_path}. "
            "Create it from wolt_positions.example.csv"
        )

    with open(config_path, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        if reader.fieldnames is None:
            raise ValueError(f"CSV has no header: {config_path}")

        positions: list[Position] = []
        errors: list[str] = []
        for row_index, row in enumerate(reader, start=2):
            row = {k: (v or "").strip() for k, v in row.items()}
            if not any(row.values()):
                continue

            venue_slug = row.get("venue_slug") or extract_venue_slug(row.get("venue_url", ""))
            search_query = row.get("search_query") or row.get("product_name") or row.get("item_id")

            if not venue_slug:
                errors.append(f"Line {row_index}: missing venue_slug/venue_url")
                continue
            if not search_query:
                errors.append(
                    f"Line {row_index}: missing search_query (or product_name/item_id as fallback)"
                )
                continue

            position = Position(
                row_index=row_index,
                position_id=row.get("position_id", ""),
                pharmacy=row.get("pharmacy", ""),
                venue_url=row.get("venue_url", ""),
                venue_slug=venue_slug,
                search_query=search_query,
                product_name=row.get("product_name", ""),
                item_id=row.get("item_id", ""),
                strict_name=parse_bool(row.get("strict_name", "")),
            )
            positions.append(position)

        if errors:
            joined = "\n".join(errors)
            raise ValueError(f"Config validation failed:\n{joined}")
        if not positions:
            raise ValueError(f"No positions found in {config_path}")
        return positions


def fetch_items(
    session: requests.Session,
    venue_slug: str,
    search_query: str,
    language: str,
    timeout: float,
) -> dict[str, Any]:
    url = (
        "https://consumer-api.wolt.com/consumer-api/consumer-assortment/"
        f"v1/venues/slug/{venue_slug}/assortment/items/search"
    )
    response = session.post(
        url,
        params={"language": language},
        json={"q": search_query},
        timeout=timeout,
    )
    response.raise_for_status()
    payload = response.json()
    if not isinstance(payload, dict):
        raise ValueError("Unexpected API payload type")
    return payload


def pick_best_match(items: list[dict[str, Any]], product_name: str) -> dict[str, Any]:
    if not items:
        return {}
    if not product_name:
        return items[0]

    target = normalize_text(product_name)

    def rank(item: dict[str, Any]) -> tuple[int, int, int, int]:
        name = normalize_text(str(item.get("name", "")))
        is_exact = 0 if name == target else 1
        is_prefix = 0 if name.startswith(target) else 1
        contains_idx = name.find(target)
        contains_rank = contains_idx if contains_idx >= 0 else 1_000_000
        length_delta = abs(len(name) - len(target))
        return (is_exact, is_prefix, contains_rank, length_delta)

    return sorted(items, key=rank)[0]


def filter_matches(position: Position, items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    filtered = items

    if position.item_id:
        filtered = [item for item in filtered if str(item.get("id", "")) == position.item_id]

    if position.product_name:
        target = normalize_text(position.product_name)
        if position.strict_name:
            filtered = [
                item
                for item in filtered
                if normalize_text(str(item.get("name", ""))) == target
            ]
        else:
            filtered = [
                item
                for item in filtered
                if target in normalize_text(str(item.get("name", "")))
            ]

    return filtered


def parse_disable_info(item: dict[str, Any]) -> tuple[bool, str, str]:
    disabled_info = item.get("disabled_info")
    if not isinstance(disabled_info, dict):
        return False, "", ""

    disable_text = str(disabled_info.get("disable_text") or "")
    disable_reason = str(disabled_info.get("disable_reason") or "")
    lower = normalize_text(disable_text)
    sold_out = any(marker in lower for marker in SOLD_OUT_MARKERS)
    return sold_out, disable_text, disable_reason


def minor_to_major(value: Any) -> float | None:
    if isinstance(value, (int, float)):
        return round(float(value) / 100.0, 2)
    return None


def build_result_row(
    checked_at: str,
    position: Position,
    payload: dict[str, Any] | None,
    error: str = "",
) -> tuple[dict[str, Any], dict[str, Any]]:
    base_row: dict[str, Any] = {
        "checked_at": checked_at,
        "position_id": position.position_id,
        "state_key": position.state_key,
        "pharmacy": position.pharmacy,
        "venue_slug": position.venue_slug,
        "venue_url": position.venue_url,
        "search_query": position.search_query,
        "product_name_input": position.product_name,
        "item_id_input": position.item_id,
        "strict_name": position.strict_name,
        "status": "error",
        "in_stock": False,
        "search_items_count": 0,
        "matches_count": 0,
        "matched_item_id": "",
        "matched_name": "",
        "disable_text": "",
        "disable_reason": "",
        "price_minor": "",
        "price": "",
        "original_price_minor": "",
        "original_price": "",
        "purchasable_balance": "",
        "error": error,
    }

    if error:
        state = {
            "checked_at": checked_at,
            "status": "error",
            "in_stock": False,
            "matched_item_id": "",
            "price_minor": "",
            "purchasable_balance": "",
            "disable_text": "",
        }
        return base_row, state

    payload = payload or {}
    items = payload.get("items") if isinstance(payload.get("items"), list) else []
    matches = filter_matches(position, items)
    base_row["search_items_count"] = len(items)
    base_row["matches_count"] = len(matches)

    if not matches:
        base_row["status"] = "not_found"
        state = {
            "checked_at": checked_at,
            "status": "not_found",
            "in_stock": False,
            "matched_item_id": "",
            "price_minor": "",
            "purchasable_balance": "",
            "disable_text": "",
        }
        return base_row, state

    item = pick_best_match(matches, position.product_name)
    sold_out, disable_text, disable_reason = parse_disable_info(item)
    status = "out_of_stock" if sold_out else ("unavailable" if disable_text else "in_stock")
    in_stock = status == "in_stock"

    price_minor = item.get("price")
    original_price_minor = item.get("original_price")
    purchasable_balance = item.get("purchasable_balance")

    base_row.update(
        {
            "status": status,
            "in_stock": in_stock,
            "matched_item_id": str(item.get("id", "")),
            "matched_name": str(item.get("name", "")),
            "disable_text": disable_text,
            "disable_reason": disable_reason,
            "price_minor": price_minor if price_minor is not None else "",
            "price": minor_to_major(price_minor) if price_minor is not None else "",
            "original_price_minor": (
                original_price_minor if original_price_minor is not None else ""
            ),
            "original_price": (
                minor_to_major(original_price_minor) if original_price_minor is not None else ""
            ),
            "purchasable_balance": (
                purchasable_balance if purchasable_balance is not None else ""
            ),
        }
    )

    state = {
        "checked_at": checked_at,
        "status": status,
        "in_stock": in_stock,
        "matched_item_id": base_row["matched_item_id"],
        "price_minor": base_row["price_minor"],
        "purchasable_balance": base_row["purchasable_balance"],
        "disable_text": disable_text,
    }
    return base_row, state


def read_state(path: Path) -> dict[str, dict[str, Any]]:
    if not path.exists():
        return {}
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        if isinstance(data, dict):
            return data
    except Exception:
        return {}
    return {}


def write_state(path: Path, data: dict[str, dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)


def diff_states(
    previous: dict[str, dict[str, Any]],
    current: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    changes: list[dict[str, Any]] = []
    compared_fields = ("status", "matched_item_id", "price_minor", "purchasable_balance")

    for state_key, now_data in current.items():
        prev_data = previous.get(state_key)
        if prev_data is None:
            changes.append(
                {
                    "state_key": state_key,
                    "change_type": "new",
                    "prev_status": "",
                    "new_status": now_data.get("status", ""),
                    "prev_price_minor": "",
                    "new_price_minor": now_data.get("price_minor", ""),
                    "prev_purchasable_balance": "",
                    "new_purchasable_balance": now_data.get("purchasable_balance", ""),
                }
            )
            continue

        changed = any(prev_data.get(field) != now_data.get(field) for field in compared_fields)
        if changed:
            changes.append(
                {
                    "state_key": state_key,
                    "change_type": "updated",
                    "prev_status": prev_data.get("status", ""),
                    "new_status": now_data.get("status", ""),
                    "prev_price_minor": prev_data.get("price_minor", ""),
                    "new_price_minor": now_data.get("price_minor", ""),
                    "prev_purchasable_balance": prev_data.get("purchasable_balance", ""),
                    "new_purchasable_balance": now_data.get("purchasable_balance", ""),
                }
            )

    for state_key, prev_data in previous.items():
        if state_key not in current:
            changes.append(
                {
                    "state_key": state_key,
                    "change_type": "removed",
                    "prev_status": prev_data.get("status", ""),
                    "new_status": "",
                    "prev_price_minor": prev_data.get("price_minor", ""),
                    "new_price_minor": "",
                    "prev_purchasable_balance": prev_data.get("purchasable_balance", ""),
                    "new_purchasable_balance": "",
                }
            )

    return changes


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({key: row.get(key, "") for key in fieldnames})


def run_monitor(args: argparse.Namespace) -> int:
    config_path = Path(args.config).expanduser().resolve()
    results_dir = Path(args.results_dir).expanduser().resolve()
    state_path = Path(args.state_file).expanduser().resolve()

    positions = read_positions(config_path)
    checked_at = datetime.now().isoformat(timespec="seconds")
    session = build_session()

    results: list[dict[str, Any]] = []
    current_state: dict[str, dict[str, Any]] = {}

    for position in positions:
        try:
            payload = fetch_items(
                session=session,
                venue_slug=position.venue_slug,
                search_query=position.search_query,
                language=args.language,
                timeout=args.timeout,
            )
            row, state_row = build_result_row(checked_at, position, payload)
        except Exception as exc:
            row, state_row = build_result_row(checked_at, position, payload=None, error=str(exc))

        results.append(row)
        current_state[position.state_key] = state_row

    previous_state = read_state(state_path)
    changes = diff_states(previous_state, current_state)
    write_state(state_path, current_state)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    results_path = results_dir / f"{args.output_prefix}_{timestamp}.csv"
    changes_path = results_dir / f"{args.output_prefix}_changes_{timestamp}.csv"
    snapshot_path = results_dir / f"{args.output_prefix}_snapshot_{timestamp}.json"

    result_fields = [
        "checked_at",
        "position_id",
        "state_key",
        "pharmacy",
        "venue_slug",
        "venue_url",
        "search_query",
        "product_name_input",
        "item_id_input",
        "strict_name",
        "status",
        "in_stock",
        "search_items_count",
        "matches_count",
        "matched_item_id",
        "matched_name",
        "disable_text",
        "disable_reason",
        "price_minor",
        "price",
        "original_price_minor",
        "original_price",
        "purchasable_balance",
        "error",
    ]
    write_csv(results_path, results, result_fields)

    change_fields = [
        "state_key",
        "change_type",
        "prev_status",
        "new_status",
        "prev_price_minor",
        "new_price_minor",
        "prev_purchasable_balance",
        "new_purchasable_balance",
    ]
    write_csv(changes_path, changes, change_fields)

    with open(snapshot_path, "w", encoding="utf-8") as fh:
        json.dump(current_state, fh, ensure_ascii=False, indent=2)

    counts: dict[str, int] = {}
    for row in results:
        status = str(row.get("status", "unknown"))
        counts[status] = counts.get(status, 0) + 1

    counts_text = ", ".join(f"{k}={v}" for k, v in sorted(counts.items()))
    print(f"Checked {len(results)} positions at {checked_at}")
    print(f"Status summary: {counts_text}")
    print(f"Results CSV: {results_path}")
    print(f"Changes CSV: {changes_path} ({len(changes)} changes)")
    print(f"Snapshot JSON: {snapshot_path}")
    print(f"State file updated: {state_path}")

    return 1 if counts.get("error", 0) else 0


def main() -> None:
    try:
        args = parse_args()
        exit_code = run_monitor(args)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(2)
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
