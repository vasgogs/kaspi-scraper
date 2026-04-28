#!/usr/bin/env python3
"""Run Glovo brand monitor for all enabled brands from portfolio CSV."""

from __future__ import annotations

import argparse
import csv
import re
import subprocess
import sys
import time
from pathlib import Path


PHARMACY_REPORT_PATTERN = re.compile(
    r"^glovo_almaty_pharmacies_(?P<stamp>\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2})\.csv$",
    re.IGNORECASE,
)


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "").strip().lower())


def parse_bool(value: str) -> bool:
    return normalize_text(value) in {"1", "true", "yes", "y", "on", "да"}


def has_nonempty_rows(csv_path: Path) -> bool:
    if not csv_path.exists() or not csv_path.is_file():
        return False
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            if any((v or "").strip() for v in row.values()):
                return True
    return False


def stamp_sort_key(path: Path) -> tuple[str, float]:
    match = PHARMACY_REPORT_PATTERN.match(path.name)
    stamp = match.group("stamp") if match else ""
    try:
        mtime = path.stat().st_mtime
    except OSError:
        mtime = 0.0
    return stamp, mtime


def resolve_pharmacies_csv(results_dir: Path, explicit: str) -> Path:
    if explicit:
        path = Path(explicit).expanduser().resolve()
        if not has_nonempty_rows(path):
            raise FileNotFoundError(f"Pharmacies CSV is empty or missing: {path}")
        return path

    candidates = [
        path
        for path in results_dir.glob("glovo_almaty_pharmacies_*.csv")
        if path.is_file() and PHARMACY_REPORT_PATTERN.match(path.name)
    ]
    candidates.sort(key=stamp_sort_key, reverse=True)
    for path in candidates:
        if has_nonempty_rows(path):
            return path.resolve()
    raise FileNotFoundError(f"No non-empty pharmacies CSV found in {results_dir}")


def load_portfolio(portfolio_csv: Path, only_brand: str, stage: str) -> list[dict[str, object]]:
    if not portfolio_csv.exists():
        raise FileNotFoundError(f"Portfolio CSV not found: {portfolio_csv}")
    grouped: dict[str, dict[str, object]] = {}
    only_brand_key = normalize_text(only_brand)
    stage_key = normalize_text(stage)

    with open(portfolio_csv, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            brand = (row.get("brand") or "").strip()
            query = (row.get("query") or "").strip() or brand
            enabled = parse_bool(row.get("enabled") or "")
            row_stage = (row.get("stage") or "").strip()

            if not brand or not enabled:
                continue
            if only_brand_key and normalize_text(brand) != only_brand_key:
                continue
            if stage_key and normalize_text(row_stage) != stage_key:
                continue

            brand_key = normalize_text(brand)
            group = grouped.setdefault(
                brand_key,
                {"brand": brand, "queries": [], "stage": row_stage},
            )
            queries = group["queries"]
            if isinstance(queries, list) and query not in queries:
                queries.append(query)

    return list(grouped.values())


def parse_args() -> tuple[argparse.Namespace, list[str]]:
    default_root = Path(__file__).resolve().parents[2]
    default_portfolio = default_root / "glovo_project" / "config" / "portfolio.csv"
    default_results = default_root / "glovo_project" / "RESULTS"

    parser = argparse.ArgumentParser(
        description="Run glovo_brand_search_monitor.py for all enabled brands from portfolio.csv"
    )
    parser.add_argument("--portfolio-csv", default=str(default_portfolio), help="Portfolio CSV path")
    parser.add_argument("--results-dir", default=str(default_results), help="Results directory")
    parser.add_argument("--pharmacies-csv", default="", help="Explicit pharmacies CSV")
    parser.add_argument("--only-brand", default="", help="Run only this brand (exact match, case-insensitive)")
    parser.add_argument("--stage", default="", help="Filter portfolio rows by stage column")
    parser.add_argument("--send-telegram", action="store_true", help="Forward --send-telegram to monitor runs")
    parser.add_argument("--dry-run", action="store_true", help="Print resolved commands without execution")
    parser.add_argument("--continue-on-error", action="store_true", help="Continue if one brand fails")
    parser.add_argument(
        "--between-brands-sleep-ms",
        type=int,
        default=1200,
        help="Pause between brand runs to reduce Glovo rate limits",
    )
    args, extra = parser.parse_known_args()
    extra_args = [arg for arg in extra if arg != "--"]
    return args, extra_args


def main() -> int:
    args, extra_args = parse_args()
    root_dir = Path(__file__).resolve().parents[2]
    monitor_script = root_dir / "glovo_brand_search_monitor.py"

    if not monitor_script.exists():
        raise FileNotFoundError(f"Missing monitor script: {monitor_script}")

    portfolio_csv = Path(args.portfolio_csv).expanduser().resolve()
    results_dir = Path(args.results_dir).expanduser().resolve()
    pharmacies_csv = resolve_pharmacies_csv(results_dir=results_dir, explicit=args.pharmacies_csv)
    portfolio_rows = load_portfolio(portfolio_csv=portfolio_csv, only_brand=args.only_brand, stage=args.stage)
    if not portfolio_rows:
        raise ValueError("No enabled portfolio rows matched filters")

    print(f"Portfolio: {portfolio_csv}")
    print(f"Pharmacies CSV: {pharmacies_csv}")
    print(f"Brands to run: {len(portfolio_rows)}")
    if extra_args:
        print(f"Forwarded args: {' '.join(extra_args)}")

    errors: list[tuple[str, int]] = []
    for idx, row in enumerate(portfolio_rows, start=1):
        brand = row["brand"]
        queries = [str(value).strip() for value in row.get("queries", []) if str(value).strip()]
        stage = str(row.get("stage") or "")
        cmd = [
            sys.executable,
            str(monitor_script),
            "--pharmacies-csv",
            str(pharmacies_csv),
            "--brand",
            brand,
        ]
        for query in queries or [brand]:
            cmd.extend(["--query", query])
        if args.send_telegram:
            cmd.append("--send-telegram")
        cmd.extend(extra_args)

        query_label = " | ".join(queries or [brand])
        print(f"[{idx}/{len(portfolio_rows)}] brand={brand} query={query_label} stage={stage or '-'}")
        print("  " + " ".join(cmd))
        if args.dry_run:
            continue

        completed = subprocess.run(cmd)
        if completed.returncode != 0:
            errors.append((brand, completed.returncode))
            if not args.continue_on_error:
                break
        if idx < len(portfolio_rows) and args.between_brands_sleep_ms > 0:
            time.sleep(args.between_brands_sleep_ms / 1000.0)

    if errors:
        print("Failures:")
        for brand, code in errors:
            print(f"- {brand}: exit_code={code}")
        return 1

    print("Portfolio run complete")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
