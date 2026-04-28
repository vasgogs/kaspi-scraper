#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import re
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
RESULTS_DIR = BASE_DIR / "RESULTS"
SNAPSHOT_RE = re.compile(r"(\d{4}-\d{2}-\d{2})_(\d{2})-(\d{2})-(\d{2})")
PRODUCT_CODE_RE = re.compile(r"-(\d+)(?=/?(?:\?.*)?$)")
REGION_SLUGS = {
    "Алматы": "almaty",
    "Астана": "astana",
    "Шымкент": "shymkent",
}


def extract_product_code(value: object) -> str:
    if value is None:
        return ""
    raw = str(value).strip()
    if not raw:
        return ""
    if raw.isdigit() and len(raw) >= 6:
        return raw
    if "kaspi.kz" in raw:
        matches = PRODUCT_CODE_RE.findall(raw)
        return matches[-1] if matches else ""
    return ""


def parse_snapshot_dt(path: Path) -> datetime | None:
    match = SNAPSHOT_RE.search(path.name)
    if not match:
        return None
    return datetime.strptime(
        f"{match.group(1)} {match.group(2)}:{match.group(3)}:{match.group(4)}",
        "%Y-%m-%d %H:%M:%S",
    )


def slugify_region(region: str) -> str:
    raw = REGION_SLUGS.get(region)
    if raw:
        return raw
    safe = re.sub(r"[^a-z0-9]+", "_", region.lower())
    return safe.strip("_") or "region"


def select_morning_snapshots() -> list[Path]:
    by_day: dict[str, list[tuple[datetime, Path]]] = defaultdict(list)
    for path in RESULTS_DIR.glob("kaspi_prices*.xlsx"):
        dt = parse_snapshot_dt(path)
        if not dt:
            continue
        if not (7 <= dt.hour <= 11):
            continue
        by_day[dt.strftime("%Y-%m-%d")].append((dt, path))

    selected: list[Path] = []
    for day in sorted(by_day):
        options = sorted(
            by_day[day],
            key=lambda item: (
                item[0],
                "reviews_refreshed" in item[1].name,
                "almaty_march" in item[1].name,
                len(item[1].name),
            ),
        )
        selected.append(options[0][1])
    if not selected:
        raise FileNotFoundError("No morning kaspi_prices snapshots found")
    return selected


def load_portfolio_codes(path: Path) -> tuple[dict[str, str], dict[str, str]]:
    product_name_by_code: dict[str, str] = {}
    url_by_code: dict[str, str] = {}
    with open(path, encoding="utf-8", newline="") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            url = (row.get("product_link") or "").strip()
            code = extract_product_code(url)
            name = (row.get("product_name") or "").strip()
            if code:
                url_by_code[code] = url
                if name:
                    product_name_by_code[code] = name
    return product_name_by_code, url_by_code


def load_brand_map(path: Path) -> dict[str, str]:
    brand_by_code: dict[str, str] = {}
    with open(path, encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            code = (row.get("portfolio_product_code") or "").strip()
            brand = (row.get("portfolio_brand") or "").strip()
            if code and brand:
                brand_by_code[code] = brand
    return brand_by_code


def normalize_snapshot(df: pd.DataFrame, portfolio_codes: set[str], region: str) -> pd.DataFrame:
    work = df.copy()
    if "region" in work.columns:
        work["region"] = work["region"].fillna("").astype(str).str.strip()
        work = work[work["region"] == region]
    raw_product_code = work["product_code"] if "product_code" in work.columns else pd.Series("", index=work.index)
    work["product_code"] = ""
    if "product_url" in work.columns:
        work["product_code"] = work["product_url"].apply(extract_product_code)
    if "product_code" in df.columns:
        missing = work["product_code"] == ""
        work.loc[missing, "product_code"] = raw_product_code[missing].apply(extract_product_code)
    work["seller"] = work.get("seller", "").fillna("").astype(str).str.strip()
    work["price_kzt"] = pd.to_numeric(work.get("price_kzt"), errors="coerce")
    work["product_name_norm"] = (
        work.get("input_product", pd.Series(index=work.index, dtype=str)).fillna("").astype(str).str.strip()
    )
    if "product" in work.columns:
        missing_name = work["product_name_norm"] == ""
        work.loc[missing_name, "product_name_norm"] = (
            work.loc[missing_name, "product"].fillna("").astype(str).str.strip()
        )
    work = work[
        work["product_code"].isin(portfolio_codes)
        & work["seller"].ne("")
        & work["price_kzt"].notna()
    ].copy()
    return work


def aggregate(
    files: list[Path],
    product_name_by_code: dict[str, str],
    url_by_code: dict[str, str],
    brand_by_code: dict[str, str],
    region: str,
) -> tuple[list[dict], list[dict], list[dict]]:
    portfolio_codes = set(product_name_by_code)
    total_morning_days = len(files)
    first_seller_counts: dict[tuple[str, str], int] = defaultdict(int)
    seen_days_by_code: dict[str, set[str]] = defaultdict(set)
    price_ranges: dict[tuple[str, str], list[float]] = defaultdict(list)

    for path in files:
        dt = parse_snapshot_dt(path)
        if not dt:
            continue
        day = dt.strftime("%Y-%m-%d")
        df = pd.read_excel(path)
        df = normalize_snapshot(df, portfolio_codes=portfolio_codes, region=region)
        if df.empty:
            continue

        for product_code, group in df.groupby("product_code", sort=False):
            first_row = group.iloc[0]
            seller = str(first_row["seller"]).strip()
            first_seller_counts[(product_code, seller)] += 1
            seen_days_by_code[product_code].add(day)
            price_ranges[(product_code, seller)].append(float(first_row["price_kzt"]))

    detailed_rows: list[dict] = []
    summary_rows: list[dict] = []
    wide_rows: list[dict] = []
    for product_code in sorted(portfolio_codes):
        available_days = len(seen_days_by_code.get(product_code, set()))
        missing_days = total_morning_days - available_days
        product_name = product_name_by_code.get(product_code, "")
        product_url = url_by_code.get(product_code, "")
        brand = brand_by_code.get(product_code, "")

        sku_rows = []
        for (code, seller), count in first_seller_counts.items():
            if code != product_code:
                continue
            prices = price_ranges[(code, seller)]
            row = {
                "product_code": product_code,
                "product_name": product_name,
                "product_url": product_url,
                "brand": brand,
                "seller": seller,
                "first_price_days": count,
                "morning_days_total": total_morning_days,
                "morning_days_available": available_days,
                "morning_days_missing": missing_days,
                "share_of_total_days_pct": round(count / total_morning_days * 100, 2) if total_morning_days else 0.0,
                "share_of_available_days_pct": round(count / available_days * 100, 2) if available_days else 0.0,
                "first_price_min_kzt": int(min(prices)) if prices else "",
                "first_price_max_kzt": int(max(prices)) if prices else "",
                "first_price_avg_kzt": round(sum(prices) / len(prices), 2) if prices else "",
            }
            sku_rows.append(row)
            detailed_rows.append(row)

        sku_rows.sort(
            key=lambda row: (
                -row["share_of_total_days_pct"],
                -row["first_price_days"],
                row["seller"],
            )
        )
        for rank, row in enumerate(sku_rows[:3], start=1):
            summary_rows.append({"rank": rank, **row})

        wide_row = {
            "product_code": product_code,
            "product_name": product_name,
            "product_url": product_url,
            "brand": brand,
            "morning_days_total": total_morning_days,
            "morning_days_available": available_days,
            "morning_days_missing": missing_days,
        }
        for rank in range(1, 4):
            if rank <= len(sku_rows):
                row = sku_rows[rank - 1]
                wide_row[f"top_{rank}_seller"] = row["seller"]
                wide_row[f"top_{rank}_days"] = row["first_price_days"]
                wide_row[f"top_{rank}_share_total_pct"] = row["share_of_total_days_pct"]
                wide_row[f"top_{rank}_share_available_pct"] = row["share_of_available_days_pct"]
                wide_row[f"top_{rank}_price_min_kzt"] = row["first_price_min_kzt"]
                wide_row[f"top_{rank}_price_max_kzt"] = row["first_price_max_kzt"]
                wide_row[f"top_{rank}_price_avg_kzt"] = row["first_price_avg_kzt"]
            else:
                wide_row[f"top_{rank}_seller"] = ""
                wide_row[f"top_{rank}_days"] = 0
                wide_row[f"top_{rank}_share_total_pct"] = 0.0
                wide_row[f"top_{rank}_share_available_pct"] = 0.0
                wide_row[f"top_{rank}_price_min_kzt"] = ""
                wide_row[f"top_{rank}_price_max_kzt"] = ""
                wide_row[f"top_{rank}_price_avg_kzt"] = ""
        wide_rows.append(wide_row)

        if not sku_rows:
            summary_rows.append(
                {
                    "rank": 1,
                    "product_code": product_code,
                    "product_name": product_name,
                    "product_url": product_url,
                    "brand": brand,
                    "seller": "",
                    "first_price_days": 0,
                    "morning_days_total": total_morning_days,
                    "morning_days_available": 0,
                    "morning_days_missing": total_morning_days,
                    "share_of_total_days_pct": 0.0,
                    "share_of_available_days_pct": 0.0,
                    "first_price_min_kzt": "",
                    "first_price_max_kzt": "",
                    "first_price_avg_kzt": "",
                }
            )

    detailed_rows.sort(
        key=lambda row: (
            row["brand"],
            row["product_name"],
            -row["share_of_total_days_pct"],
            -row["first_price_days"],
            row["seller"],
        )
    )
    summary_rows.sort(
        key=lambda row: (
            row["brand"],
            row["product_name"],
            row["rank"],
        )
    )
    wide_rows.sort(key=lambda row: (row["brand"], row["product_name"]))
    return detailed_rows, summary_rows, wide_rows


def write_csv(path: Path, rows: list[dict], fieldnames: list[str]) -> None:
    with open(path, "w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    parser = argparse.ArgumentParser(description="Analyze morning first-price seller shares by SKU")
    parser.add_argument("--portfolio-file", type=Path, default=BASE_DIR / "my_products.csv")
    parser.add_argument("--brand-file", type=Path, default=None)
    parser.add_argument("--region", default="Алматы")
    args = parser.parse_args()

    files = select_morning_snapshots()
    product_name_by_code, url_by_code = load_portfolio_codes(args.portfolio_file)
    brand_file = args.brand_file or max(RESULTS_DIR.glob("portfolio_direct_competitors_*_v2.csv"), key=lambda p: p.stat().st_mtime)
    brand_by_code = load_brand_map(brand_file)
    detailed_rows, summary_rows, wide_rows = aggregate(
        files=files,
        product_name_by_code=product_name_by_code,
        url_by_code=url_by_code,
        brand_by_code=brand_by_code,
        region=args.region,
    )

    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    region_slug = slugify_region(args.region)
    detailed_path = RESULTS_DIR / f"morning_first_price_share_by_sku_{region_slug}_{stamp}_detailed.csv"
    summary_path = RESULTS_DIR / f"morning_first_price_share_by_sku_{region_slug}_{stamp}_top3.csv"
    wide_path = RESULTS_DIR / f"morning_first_price_share_by_sku_{region_slug}_{stamp}_wide.csv"

    fields = [
        "rank",
        "product_code",
        "product_name",
        "product_url",
        "brand",
        "seller",
        "first_price_days",
        "morning_days_total",
        "morning_days_available",
        "morning_days_missing",
        "share_of_total_days_pct",
        "share_of_available_days_pct",
        "first_price_min_kzt",
        "first_price_max_kzt",
        "first_price_avg_kzt",
    ]
    write_csv(detailed_path, detailed_rows, fields[1:])
    write_csv(summary_path, summary_rows, fields)
    write_csv(
        wide_path,
        wide_rows,
        [
            "product_code",
            "product_name",
            "product_url",
            "brand",
            "morning_days_total",
            "morning_days_available",
            "morning_days_missing",
            "top_1_seller",
            "top_1_days",
            "top_1_share_total_pct",
            "top_1_share_available_pct",
            "top_1_price_min_kzt",
            "top_1_price_max_kzt",
            "top_1_price_avg_kzt",
            "top_2_seller",
            "top_2_days",
            "top_2_share_total_pct",
            "top_2_share_available_pct",
            "top_2_price_min_kzt",
            "top_2_price_max_kzt",
            "top_2_price_avg_kzt",
            "top_3_seller",
            "top_3_days",
            "top_3_share_total_pct",
            "top_3_share_available_pct",
            "top_3_price_min_kzt",
            "top_3_price_max_kzt",
            "top_3_price_avg_kzt",
        ],
    )

    print(f"morning_files={len(files)}")
    print(f"portfolio_codes={len(product_name_by_code)}")
    print(f"region={args.region}")
    print(f"detailed={detailed_path}")
    print(f"summary={summary_path}")
    print(f"wide={wide_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
