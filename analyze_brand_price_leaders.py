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
PRODUCT_CODE_RE = re.compile(r"-(\d+)(?:/|\\?|$)")


def latest_file(pattern: str) -> Path:
    files = sorted(RESULTS_DIR.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError(f"No files found for pattern: {pattern}")
    return files[0]


def snapshot_files() -> list[Path]:
    files = sorted(RESULTS_DIR.glob("kaspi_prices_*.xlsx"), key=lambda p: p.name)
    if not files:
        raise FileNotFoundError("No kaspi_prices_*.xlsx files found")
    return files


def parse_snapshot_period(path: Path) -> tuple[str, str]:
    match = SNAPSHOT_RE.search(path.name)
    if not match:
        raise ValueError(f"Unexpected snapshot filename: {path.name}")
    dt = datetime.strptime(
        f"{match.group(1)} {match.group(2)}:{match.group(3)}:{match.group(4)}",
        "%Y-%m-%d %H:%M:%S",
    )
    return dt.strftime("%Y-%m"), dt.strftime("%Y-%m-%d %H:%M:%S")


def extract_product_code(value: object) -> str:
    if value is None:
        return ""
    raw = str(value).strip()
    if not raw:
        return ""
    if raw.isdigit():
        return raw
    match = PRODUCT_CODE_RE.search(raw)
    return match.group(1) if match else ""


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


def normalize_snapshot(df: pd.DataFrame, region: str, brand_by_code: dict[str, str]) -> pd.DataFrame:
    work = df.copy()
    if "region" in work.columns:
        work["region"] = work["region"].fillna("").astype(str).str.strip()
        work = work[work["region"] == region]
    if "product_code" in work.columns:
        work["product_code"] = work["product_code"].apply(extract_product_code)
    else:
        work["product_code"] = ""
    if "product_url" in work.columns:
        missing_code = work["product_code"] == ""
        work.loc[missing_code, "product_code"] = work.loc[missing_code, "product_url"].apply(extract_product_code)
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
        work["product_code"].isin(brand_by_code)
        & work["seller"].ne("")
        & work["price_kzt"].notna()
    ].copy()
    work["portfolio_brand"] = work["product_code"].map(brand_by_code)
    return work


def aggregate_history(files: list[Path], brand_by_code: dict[str, str], region: str) -> tuple[list[dict], list[dict], dict]:
    brand_period_meta: dict[tuple[str, str], dict[str, object]] = defaultdict(
        lambda: {"products": set(), "snapshot_files": set(), "snapshot_units": 0}
    )
    seller_stats: dict[tuple[str, str, str], dict[str, object]] = defaultdict(
        lambda: {
            "leader_units_fractional": 0.0,
            "leader_units_raw": 0,
            "solo_lead_units": 0,
            "tied_lead_units": 0,
            "presence_units": 0,
            "products_led": set(),
            "products_seen": set(),
        }
    )

    matched_files = 0
    matched_rows = 0

    for path in files:
        month, snapshot_label = parse_snapshot_period(path)
        df = pd.read_excel(path)
        df = normalize_snapshot(df, region=region, brand_by_code=brand_by_code)
        if df.empty:
            continue
        matched_files += 1
        matched_rows += len(df)

        for product_code, group in df.groupby("product_code"):
            brand = str(group["portfolio_brand"].iloc[0]).strip()
            product_name = str(group["product_name_norm"].iloc[0]).strip()
            min_price = float(group["price_kzt"].min())
            winners = sorted(set(group.loc[group["price_kzt"] == min_price, "seller"].astype(str)))
            if not winners:
                continue
            winner_weight = 1.0 / len(winners)
            all_sellers = sorted(set(group["seller"].astype(str)))

            for period in (month, "all"):
                meta = brand_period_meta[(period, brand)]
                meta["products"].add(product_code)
                meta["snapshot_files"].add(snapshot_label)
                meta["snapshot_units"] = int(meta["snapshot_units"]) + 1

                for seller in all_sellers:
                    stat = seller_stats[(period, brand, seller)]
                    stat["presence_units"] = int(stat["presence_units"]) + 1
                    stat["products_seen"].add(product_code)
                    if seller in winners:
                        stat["leader_units_fractional"] = float(stat["leader_units_fractional"]) + winner_weight
                        stat["leader_units_raw"] = int(stat["leader_units_raw"]) + 1
                        if len(winners) == 1:
                            stat["solo_lead_units"] = int(stat["solo_lead_units"]) + 1
                        else:
                            stat["tied_lead_units"] = int(stat["tied_lead_units"]) + 1
                        stat["products_led"].add(product_code)
                    stat["product_name_last"] = product_name

    detailed_rows: list[dict] = []
    grouped_rows: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for (period, brand, seller), stat in seller_stats.items():
        meta = brand_period_meta[(period, brand)]
        denominator = int(meta["snapshot_units"])
        if denominator <= 0:
            continue
        row = {
            "period": period,
            "brand": brand,
            "seller": seller,
            "leader_units_fractional": round(float(stat["leader_units_fractional"]), 4),
            "leader_units_raw": int(stat["leader_units_raw"]),
            "leader_share_pct": round(float(stat["leader_units_fractional"]) / denominator * 100, 2),
            "solo_lead_units": int(stat["solo_lead_units"]),
            "tied_lead_units": int(stat["tied_lead_units"]),
            "presence_units": int(stat["presence_units"]),
            "presence_share_pct": round(int(stat["presence_units"]) / denominator * 100, 2),
            "brand_snapshot_units": denominator,
            "products_led_count": len(stat["products_led"]),
            "products_seen_count": len(stat["products_seen"]),
            "brand_products_count": len(meta["products"]),
            "snapshot_files_count": len(meta["snapshot_files"]),
        }
        detailed_rows.append(row)
        grouped_rows[(period, brand)].append(row)

    detailed_rows.sort(
        key=lambda row: (
            row["period"] != "all",
            row["period"],
            row["brand"],
            -row["leader_share_pct"],
            -row["leader_units_fractional"],
            row["seller"],
        )
    )

    summary_rows: list[dict] = []
    for (period, brand), brand_rows in grouped_rows.items():
        brand_rows.sort(
            key=lambda row: (
                -row["leader_share_pct"],
                -row["leader_units_fractional"],
                -row["products_led_count"],
                row["seller"],
            )
        )
        for rank, row in enumerate(brand_rows, start=1):
            summary_rows.append({"period": period, "brand": brand, "rank": rank, **row})

    summary_rows.sort(
        key=lambda row: (
            row["period"] != "all",
            row["period"],
            row["brand"],
            row["rank"],
        )
    )

    meta = {
        "matched_files": matched_files,
        "matched_rows": matched_rows,
        "brands_count": len({brand for _, brand in brand_period_meta}),
        "products_count": len(brand_by_code),
    }
    return detailed_rows, summary_rows, meta


def write_csv(path: Path, rows: list[dict], fieldnames: list[str]) -> None:
    with open(path, "w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    parser = argparse.ArgumentParser(description="Analyze historical best-price sellers by portfolio brand")
    parser.add_argument("--portfolio-file", type=Path, default=None, help="CSV with portfolio_brand mapping")
    parser.add_argument("--region", default="Алматы", help="Region to analyze")
    parser.add_argument("--top-n", type=int, default=3, help="How many sellers keep in top summary")
    args = parser.parse_args()

    portfolio_path = args.portfolio_file or latest_file("portfolio_direct_competitors_*_v2.csv")
    brand_by_code = load_brand_map(portfolio_path)
    files = snapshot_files()
    detailed_rows, summary_rows, meta = aggregate_history(files, brand_by_code=brand_by_code, region=args.region)

    top_n = max(1, args.top_n)
    summary_rows = [row for row in summary_rows if int(row["rank"]) <= top_n]

    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    detailed_path = RESULTS_DIR / f"brand_price_leader_history_{stamp}_detailed.csv"
    summary_path = RESULTS_DIR / f"brand_price_leader_history_{stamp}_top{top_n}.csv"

    detailed_fields = [
        "period",
        "brand",
        "seller",
        "leader_units_fractional",
        "leader_units_raw",
        "leader_share_pct",
        "solo_lead_units",
        "tied_lead_units",
        "presence_units",
        "presence_share_pct",
        "brand_snapshot_units",
        "products_led_count",
        "products_seen_count",
        "brand_products_count",
        "snapshot_files_count",
    ]
    summary_fields = [
        "period",
        "brand",
        "rank",
        "seller",
        "leader_units_fractional",
        "leader_units_raw",
        "leader_share_pct",
        "solo_lead_units",
        "tied_lead_units",
        "presence_units",
        "presence_share_pct",
        "brand_snapshot_units",
        "products_led_count",
        "products_seen_count",
        "brand_products_count",
        "snapshot_files_count",
    ]

    write_csv(detailed_path, detailed_rows, detailed_fields)
    write_csv(summary_path, summary_rows, summary_fields)

    print(f"portfolio_file={portfolio_path}")
    print(f"snapshot_files={len(files)}")
    print(f"matched_files={meta['matched_files']}")
    print(f"matched_rows={meta['matched_rows']}")
    print(f"brands={meta['brands_count']}")
    print(f"portfolio_codes={meta['products_count']}")
    print(f"detailed={detailed_path}")
    print(f"summary={summary_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
