#!/usr/bin/env python3
"""Build historical Wolt assortment coverage analysis for Almaty pharmacies."""

from __future__ import annotations

import argparse
import csv
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd


ROOT_DIR = Path(__file__).resolve().parents[2]
WOLT_PROJECT_DIR = ROOT_DIR / "wolt_project"
DEFAULT_RESULTS_DIR = WOLT_PROJECT_DIR / "RESULTS"
DEFAULT_STATE_DIR = WOLT_PROJECT_DIR / "state"
DEFAULT_PORTFOLIO_CSV = WOLT_PROJECT_DIR / "config" / "portfolio.csv"
DEFAULT_PHARMACIES_CATALOG = DEFAULT_STATE_DIR / "wolt_pharmacies_catalog.csv"
STAMP_FMT = "%Y-%m-%d_%H-%M-%S"
REPORT_PATTERN = re.compile(
    r"^wolt_brand_(?P<brand>[a-z0-9_-]+)_(?P<stamp>\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2})_items\.csv$",
    re.IGNORECASE,
)

CYR_TO_LAT = str.maketrans(
    {
        "а": "a",
        "б": "b",
        "в": "v",
        "г": "g",
        "д": "d",
        "е": "e",
        "ё": "e",
        "ж": "zh",
        "з": "z",
        "и": "i",
        "й": "i",
        "к": "k",
        "л": "l",
        "м": "m",
        "н": "n",
        "о": "o",
        "п": "p",
        "р": "r",
        "с": "s",
        "т": "t",
        "у": "u",
        "ф": "f",
        "х": "h",
        "ц": "ts",
        "ч": "ch",
        "ш": "sh",
        "щ": "sch",
        "ъ": "",
        "ы": "y",
        "ь": "",
        "э": "e",
        "ю": "yu",
        "я": "ya",
    }
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build Wolt historical assortment gap report for portfolio brands."
    )
    parser.add_argument("--city-slug", default="almaty", help="City slug filter (default: %(default)s)")
    parser.add_argument(
        "--results-dir",
        default=str(DEFAULT_RESULTS_DIR),
        help="Directory with Wolt item reports and output artifacts (default: %(default)s)",
    )
    parser.add_argument(
        "--state-dir",
        default=str(DEFAULT_STATE_DIR),
        help="Directory with canonical catalogs (default: %(default)s)",
    )
    parser.add_argument(
        "--portfolio-csv",
        default=str(DEFAULT_PORTFOLIO_CSV),
        help="Portfolio CSV path (default: %(default)s)",
    )
    parser.add_argument(
        "--pharmacies-catalog",
        default=str(DEFAULT_PHARMACIES_CATALOG),
        help="Wolt pharmacies catalog CSV path (default: %(default)s)",
    )
    parser.add_argument(
        "--output-prefix",
        default="wolt_almaty_assortment_gap",
        help="Output file prefix (default: %(default)s)",
    )
    return parser.parse_args()


def normalize_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def slugify_token(value: str) -> str:
    text = normalize_text(value).translate(CYR_TO_LAT)
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text


def humanize_brand_slug(slug: str, brand_display_map: dict[str, str]) -> str:
    label = brand_display_map.get(slug)
    if label:
        return label
    return str(slug or "").replace("_", " ").replace("-", " ").title()


def parse_bool(value: str) -> bool:
    return normalize_text(value) in {"1", "true", "yes", "y", "on", "да"}


def safe_float(value: Any) -> float | None:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    text = str(value).strip()
    if not text:
        return None
    try:
        return float(text)
    except Exception:
        return None


def safe_int(value: Any) -> int | None:
    number = safe_float(value)
    if number is None:
        return None
    try:
        return int(number)
    except Exception:
        return None


def safe_iso(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    if isinstance(value, pd.Timestamp):
        if value.tzinfo:
            return value.isoformat()
        return value.isoformat(timespec="seconds")
    if isinstance(value, datetime):
        return value.isoformat(timespec="seconds")
    text = str(value).strip()
    return text


def load_enabled_portfolio(portfolio_csv: Path) -> tuple[set[str], dict[str, str]]:
    if not portfolio_csv.exists():
        raise FileNotFoundError(f"Portfolio CSV not found: {portfolio_csv}")
    enabled: set[str] = set()
    display_map: dict[str, str] = {}
    with open(portfolio_csv, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            brand = str(row.get("brand") or "").strip()
            if not brand or not parse_bool(row.get("enabled") or ""):
                continue
            slug = slugify_token(brand)
            if not slug:
                continue
            enabled.add(slug)
            display_map.setdefault(slug, brand)
    if not enabled:
        raise ValueError("No enabled brands found in portfolio.csv")
    return enabled, display_map


def load_pharmacies_catalog(path: Path, city_slug: str) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Pharmacies catalog not found: {path}")
    df = pd.read_csv(path, dtype=str).fillna("")
    if df.empty:
        raise ValueError(f"Pharmacies catalog is empty: {path}")
    city_key = normalize_text(city_slug)
    if "last_seen_city_slug" in df.columns:
        mask = df["last_seen_city_slug"].map(normalize_text).eq(city_key)
        if mask.any():
            df = df.loc[mask].copy()
    if df.empty:
        raise ValueError(f"No pharmacies for city={city_slug} in {path}")
    df["slug"] = df["slug"].map(str).str.strip()
    df = df[df["slug"] != ""].copy()
    for col in ("name", "address", "venue_url", "lat", "lon"):
        if col not in df.columns:
            df[col] = ""
    df["lat"] = pd.to_numeric(df["lat"], errors="coerce")
    df["lon"] = pd.to_numeric(df["lon"], errors="coerce")
    df = df.sort_values(["name", "slug"]).drop_duplicates("slug", keep="last")
    return df


def report_meta(path: Path) -> tuple[str, datetime, str] | None:
    match = REPORT_PATTERN.match(path.name)
    if not match:
        return None
    brand_slug = str(match.group("brand") or "").strip().lower()
    stamp_raw = str(match.group("stamp") or "").strip()
    stamp_dt = datetime.strptime(stamp_raw, STAMP_FMT)
    return brand_slug, stamp_dt, stamp_raw


def discover_report_files(results_dir: Path, enabled_brand_slugs: set[str]) -> list[tuple[Path, str, datetime, str]]:
    files: list[tuple[Path, str, datetime, str]] = []
    for path in results_dir.glob("wolt_brand_*_items.csv"):
        meta = report_meta(path)
        if not meta:
            continue
        brand_slug, stamp_dt, stamp_raw = meta
        if brand_slug not in enabled_brand_slugs:
            continue
        files.append((path, brand_slug, stamp_dt, stamp_raw))
    files.sort(key=lambda item: item[2])
    if not files:
        raise FileNotFoundError(f"No Wolt item reports found in {results_dir}")
    return files


def latest_snapshot_files(files: list[tuple[Path, str, datetime, str]]) -> list[tuple[Path, str, datetime, str]]:
    latest_day_by_brand: dict[str, str] = {}
    for path, brand_slug, stamp_dt, stamp_raw in files:
        day = stamp_raw[:10]
        prev = latest_day_by_brand.get(brand_slug)
        if prev is None or day > prev:
            latest_day_by_brand[brand_slug] = day
    selected = [
        item
        for item in files
        if latest_day_by_brand.get(item[1]) == item[3][:10]
    ]
    selected.sort(key=lambda item: (item[1], item[2], item[0].name))
    return selected


def load_sku_catalog(state_dir: Path, enabled_brand_slugs: set[str], brand_display_map: dict[str, str]) -> pd.DataFrame:
    rows: list[pd.DataFrame] = []
    for slug in sorted(enabled_brand_slugs):
        path = state_dir / f"wolt_{slug}_canonical_catalog.csv"
        if not path.exists():
            continue
        df = pd.read_csv(path, dtype=str).fillna("")
        if df.empty or "canonical_sku" not in df.columns:
            continue
        keep_cols = ["canonical_sku", "canonical_name", "product_line", "dosage_or_volume", "pack_size", "form_factor"]
        for col in keep_cols:
            if col not in df.columns:
                df[col] = ""
        df = df[keep_cols].copy()
        df["brand_slug"] = slug
        df["brand"] = humanize_brand_slug(slug, brand_display_map)
        rows.append(df)
    if not rows:
        return pd.DataFrame(
            columns=["canonical_sku", "canonical_name", "product_line", "dosage_or_volume", "pack_size", "form_factor", "brand_slug", "brand"]
        )
    catalog = pd.concat(rows, ignore_index=True)
    catalog["canonical_sku"] = catalog["canonical_sku"].map(str).str.strip()
    catalog = catalog[catalog["canonical_sku"] != ""].copy()
    catalog = catalog.sort_values(["brand_slug", "canonical_name", "canonical_sku"]).drop_duplicates("canonical_sku", keep="first")
    return catalog


def load_item_reports(
    files: list[tuple[Path, str, datetime, str]],
    brand_display_map: dict[str, str],
) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for path, brand_slug, stamp_dt, stamp_raw in files:
        try:
            df = pd.read_csv(path)
        except Exception:
            continue
        if df.empty or "canonical_sku" not in df.columns:
            continue
        df = df.copy()
        df["canonical_sku"] = df["canonical_sku"].fillna("").astype(str).str.strip()
        df = df[df["canonical_sku"] != ""].copy()
        if df.empty:
            continue
        for col in (
            "pharmacy",
            "venue_slug",
            "venue_url",
            "status",
            "item_name",
            "item_id",
            "canonical_name",
            "checked_at",
            "price",
            "price_minor",
            "product_link",
        ):
            if col not in df.columns:
                df[col] = ""
        df["report_file"] = path.name
        df["brand_slug"] = brand_slug
        df["brand"] = humanize_brand_slug(brand_slug, brand_display_map)
        df["report_stamp"] = stamp_raw
        df["report_day"] = stamp_raw[:10]
        df["report_dt"] = pd.Timestamp(stamp_dt)
        frames.append(df)
    if not frames:
        raise ValueError("No canonicalized Wolt item rows found in historical reports")
    out = pd.concat(frames, ignore_index=True)
    out["venue_slug"] = out["venue_slug"].fillna("").astype(str).str.strip()
    out["pharmacy"] = out["pharmacy"].fillna("").astype(str).str.strip()
    out["status"] = out["status"].fillna("").astype(str).str.strip().str.lower()
    out["canonical_name"] = out["canonical_name"].fillna("").astype(str).str.strip()
    out["checked_at_dt"] = pd.to_datetime(out["checked_at"], errors="coerce")
    out.loc[out["checked_at_dt"].isna(), "checked_at_dt"] = out.loc[out["checked_at_dt"].isna(), "report_dt"]
    out["price"] = pd.to_numeric(out["price"], errors="coerce")
    out["price_minor"] = pd.to_numeric(out["price_minor"], errors="coerce")
    out = out[out["venue_slug"] != ""].copy()
    return out


def dedupe_current_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.copy()
    ranked = df.copy()
    status_rank = {"in_stock": 0, "out_of_stock": 1, "unavailable": 2}
    ranked["status_rank"] = ranked["status"].map(status_rank).fillna(9)
    ranked = ranked.sort_values(
        ["venue_slug", "canonical_sku", "status_rank", "checked_at_dt", "price"],
        ascending=[True, True, True, False, True],
    )
    ranked = ranked.drop_duplicates(["venue_slug", "canonical_sku", "item_id"], keep="first")
    ranked = ranked.drop(columns=["status_rank"], errors="ignore")
    return ranked


def aggregate_current_presence(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(
            columns=[
                "venue_slug",
                "pharmacy",
                "venue_url",
                "canonical_sku",
                "canonical_name",
                "brand_slug",
                "brand",
                "listed_now",
                "in_stock_now",
                "status_primary",
                "current_best_price",
                "current_any_price",
                "checked_at_dt",
            ]
        )

    def summarize(group: pd.DataFrame) -> pd.Series:
        in_stock = group.loc[group["status"].eq("in_stock")]
        current_best_price = in_stock["price"].min() if not in_stock.empty else pd.NA
        current_any_price = group["price"].min() if group["price"].notna().any() else pd.NA
        primary = group.sort_values(
            ["checked_at_dt", "status"],
            ascending=[False, True],
        ).iloc[0]
        return pd.Series(
            {
                "pharmacy": primary.get("pharmacy") or "",
                "venue_url": primary.get("venue_url") or "",
                "canonical_name": primary.get("canonical_name") or "",
                "brand_slug": primary.get("brand_slug") or "",
                "brand": primary.get("brand") or "",
                "listed_now": True,
                "in_stock_now": bool((group["status"] == "in_stock").any()),
                "status_primary": primary.get("status") or "",
                "current_best_price": current_best_price,
                "current_any_price": current_any_price,
                "checked_at_dt": group["checked_at_dt"].max(),
            }
        )

    aggregated = (
        df.groupby(["venue_slug", "canonical_sku"], as_index=False, dropna=False)
        .apply(summarize, include_groups=False)
        .reset_index(drop=True)
    )
    return aggregated


def aggregate_historical_presence(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(
            columns=[
                "venue_slug",
                "canonical_sku",
                "pharmacy",
                "venue_url",
                "canonical_name",
                "brand_slug",
                "brand",
                "first_seen_at",
                "last_seen_at",
                "historical_best_price",
                "historical_any_price",
                "historical_runs",
                "ever_in_stock",
            ]
        )

    def summarize(group: pd.DataFrame) -> pd.Series:
        in_stock = group.loc[group["status"].eq("in_stock")]
        hist_best_price = in_stock["price"].min() if not in_stock.empty else pd.NA
        hist_any_price = group["price"].min() if group["price"].notna().any() else pd.NA
        first = group.sort_values("checked_at_dt", ascending=True).iloc[0]
        last = group.sort_values("checked_at_dt", ascending=False).iloc[0]
        return pd.Series(
            {
                "pharmacy": last.get("pharmacy") or first.get("pharmacy") or "",
                "venue_url": last.get("venue_url") or first.get("venue_url") or "",
                "canonical_name": last.get("canonical_name") or first.get("canonical_name") or "",
                "brand_slug": last.get("brand_slug") or "",
                "brand": last.get("brand") or "",
                "first_seen_at": group["checked_at_dt"].min(),
                "last_seen_at": group["checked_at_dt"].max(),
                "historical_best_price": hist_best_price,
                "historical_any_price": hist_any_price,
                "historical_runs": int(group["report_file"].nunique()),
                "ever_in_stock": bool((group["status"] == "in_stock").any()),
            }
        )

    aggregated = (
        df.groupby(["venue_slug", "canonical_sku"], as_index=False, dropna=False)
        .apply(summarize, include_groups=False)
        .reset_index(drop=True)
    )
    return aggregated


def compute_city_sku_stats(current_presence: pd.DataFrame, historical_presence: pd.DataFrame) -> pd.DataFrame:
    current = current_presence.copy()
    historical = historical_presence.copy()

    current_rows: list[dict[str, Any]] = []
    if not current.empty:
        for canonical_sku, group in current.groupby("canonical_sku", dropna=False):
            current_rows.append(
                {
                    "canonical_sku": canonical_sku,
                    "current_pharmacies_count": int(group["venue_slug"].nunique()),
                    "current_in_stock_pharmacies_count": int(group.loc[group["in_stock_now"], "venue_slug"].nunique()),
                    "current_min_price": group.loc[group["in_stock_now"], "current_best_price"].min() if group.loc[group["in_stock_now"], "current_best_price"].notna().any() else pd.NA,
                    "current_any_min_price": group["current_any_price"].min() if group["current_any_price"].notna().any() else pd.NA,
                    "brand_slug": str(group["brand_slug"].iloc[0] or ""),
                    "brand": str(group["brand"].iloc[0] or ""),
                    "canonical_name": str(group["canonical_name"].iloc[0] or ""),
                }
            )
    current_df = pd.DataFrame(current_rows)

    historical_rows: list[dict[str, Any]] = []
    if not historical.empty:
        for canonical_sku, group in historical.groupby("canonical_sku", dropna=False):
            historical_rows.append(
                {
                    "canonical_sku": canonical_sku,
                    "historical_pharmacies_count": int(group["venue_slug"].nunique()),
                    "historical_best_price": group["historical_best_price"].min() if group["historical_best_price"].notna().any() else pd.NA,
                    "historical_any_price": group["historical_any_price"].min() if group["historical_any_price"].notna().any() else pd.NA,
                    "first_seen_at": group["first_seen_at"].min(),
                    "last_seen_at": group["last_seen_at"].max(),
                    "brand_slug": str(group["brand_slug"].iloc[0] or ""),
                    "brand": str(group["brand"].iloc[0] or ""),
                    "canonical_name": str(group["canonical_name"].iloc[0] or ""),
                }
            )
    historical_df = pd.DataFrame(historical_rows)

    if current_df.empty and historical_df.empty:
        return pd.DataFrame(
            columns=[
                "canonical_sku",
                "current_pharmacies_count",
                "current_in_stock_pharmacies_count",
                "current_min_price",
                "current_any_min_price",
                "historical_pharmacies_count",
                "historical_best_price",
                "historical_any_price",
                "first_seen_at",
                "last_seen_at",
                "brand_slug",
                "brand",
                "canonical_name",
            ]
        )

    merged = pd.merge(
        historical_df,
        current_df,
        on="canonical_sku",
        how="outer",
        suffixes=("_hist", "_current"),
    )
    merged["brand_slug"] = merged["brand_slug_current"].where(merged["brand_slug_current"].notna() & (merged["brand_slug_current"] != ""), merged["brand_slug_hist"])
    merged["brand"] = merged["brand_current"].where(merged["brand_current"].notna() & (merged["brand_current"] != ""), merged["brand_hist"])
    merged["canonical_name"] = merged["canonical_name_current"].where(
        merged["canonical_name_current"].notna() & (merged["canonical_name_current"] != ""),
        merged["canonical_name_hist"],
    )
    keep_cols = [
        "canonical_sku",
        "brand_slug",
        "brand",
        "canonical_name",
        "current_pharmacies_count",
        "current_in_stock_pharmacies_count",
        "current_min_price",
        "current_any_min_price",
        "historical_pharmacies_count",
        "historical_best_price",
        "historical_any_price",
        "first_seen_at",
        "last_seen_at",
    ]
    for col in keep_cols:
        if col not in merged.columns:
            merged[col] = pd.NA
    return merged[keep_cols].copy()


def build_gap_payload(
    *,
    pharmacies_df: pd.DataFrame,
    sku_catalog_df: pd.DataFrame,
    current_presence: pd.DataFrame,
    historical_presence: pd.DataFrame,
    city_sku_stats: pd.DataFrame,
    latest_files: list[tuple[Path, str, datetime, str]],
    all_history_files: list[tuple[Path, str, datetime, str]],
    output_prefix: str,
    catalog_file: Path,
) -> dict[str, Any]:
    pharmacies_index = {
        str(row["slug"]): row.to_dict()
        for _, row in pharmacies_df.iterrows()
    }
    portfolio_skus = set(sku_catalog_df["canonical_sku"].dropna().astype(str).str.strip())
    historical_city_skus = set(historical_presence["canonical_sku"].dropna().astype(str).str.strip())
    active_city_skus = {sku for sku in historical_city_skus if sku}
    current_available_city_skus = {
        str(row["canonical_sku"])
        for _, row in city_sku_stats.iterrows()
        if safe_float(row.get("current_min_price")) is not None
    }

    current_map = {
        (str(row["venue_slug"]), str(row["canonical_sku"])): row.to_dict()
        for _, row in current_presence.iterrows()
    }
    historical_map = {
        (str(row["venue_slug"]), str(row["canonical_sku"])): row.to_dict()
        for _, row in historical_presence.iterrows()
    }
    city_sku_map = {
        str(row["canonical_sku"]): row.to_dict()
        for _, row in city_sku_stats.iterrows()
    }
    sku_catalog_map = {
        str(row["canonical_sku"]): row.to_dict()
        for _, row in sku_catalog_df.iterrows()
    }

    pharmacy_rows: list[dict[str, Any]] = []
    gap_rows: list[dict[str, Any]] = []

    for slug, meta in pharmacies_index.items():
        current_skus = {sku for venue_slug, sku in current_map if venue_slug == slug}
        current_in_stock_skus = {
            sku
            for (venue_slug, sku), row in current_map.items()
            if venue_slug == slug and bool(row.get("in_stock_now"))
        }
        historical_skus = {sku for venue_slug, sku in historical_map if venue_slug == slug}
        missing_active_skus = active_city_skus - current_skus
        lost_historical_skus = historical_skus - current_skus
        never_listed_skus = missing_active_skus - lost_historical_skus
        out_of_stock_skus = current_skus - current_in_stock_skus
        missing_available_now_skus = current_available_city_skus - current_skus

        potential_loss_now = 0.0
        lost_value_now = 0.0
        top_missing: list[dict[str, Any]] = []

        for sku in sorted(missing_active_skus):
            sku_stat = city_sku_map.get(sku)
            sku_meta = sku_catalog_map.get(sku)
            current_ref = safe_float(sku_stat.get("current_min_price") if sku_stat is not None else None)
            historical_ref = safe_float(sku_stat.get("historical_best_price") if sku_stat is not None else None)
            reference_price = current_ref if current_ref is not None else historical_ref
            available_now_in_city = sku in current_available_city_skus
            gap_type = "lost_historical" if sku in lost_historical_skus else "never_listed_here"
            historical_row = historical_map.get((slug, sku))
            last_seen_here = historical_row.get("last_seen_at") if historical_row is not None else pd.NaT
            current_city_pharmacies = safe_int(sku_stat.get("current_pharmacies_count") if sku_stat is not None else None) or 0
            current_city_in_stock = safe_int(sku_stat.get("current_in_stock_pharmacies_count") if sku_stat is not None else None) or 0
            if available_now_in_city and current_ref is not None:
                potential_loss_now += current_ref
            if gap_type == "lost_historical" and reference_price is not None:
                lost_value_now += reference_price
            gap_row = {
                "pharmacy": str(meta.get("name") or ""),
                "venue_slug": slug,
                "venue_url": str(meta.get("venue_url") or ""),
                "address": str(meta.get("address") or ""),
                "canonical_sku": sku,
                "canonical_name": str((sku_meta or {}).get("canonical_name") or (sku_stat or {}).get("canonical_name") or sku),
                "brand_slug": str((sku_meta or {}).get("brand_slug") or (sku_stat or {}).get("brand_slug") or ""),
                "brand": str((sku_meta or {}).get("brand") or (sku_stat or {}).get("brand") or ""),
                "gap_type": gap_type,
                "available_now_in_city": available_now_in_city,
                "current_city_pharmacies_count": current_city_pharmacies,
                "current_city_in_stock_pharmacies_count": current_city_in_stock,
                "reference_price_kzt": round(reference_price, 2) if reference_price is not None else None,
                "last_seen_here_at": safe_iso(last_seen_here),
            }
            gap_rows.append(gap_row)
            top_missing.append(gap_row)

        top_missing.sort(
            key=lambda row: (
                0 if row["gap_type"] == "lost_historical" else 1,
                -(row.get("reference_price_kzt") or 0),
                str(row.get("canonical_name") or ""),
            )
        )

        pharmacy_rows.append(
            {
                "pharmacy": str(meta.get("name") or ""),
                "venue_slug": slug,
                "venue_url": str(meta.get("venue_url") or ""),
                "address": str(meta.get("address") or ""),
                "lat": safe_float(meta.get("lat")),
                "lon": safe_float(meta.get("lon")),
                "portfolio_skus_count": len(portfolio_skus),
                "active_city_skus_count": len(active_city_skus),
                "current_listed_skus_count": len(current_skus),
                "current_in_stock_skus_count": len(current_in_stock_skus),
                "historical_skus_count": len(historical_skus),
                "out_of_stock_skus_count": len(out_of_stock_skus),
                "missing_active_skus_count": len(missing_active_skus),
                "missing_available_now_skus_count": len(missing_available_now_skus),
                "lost_historical_skus_count": len(lost_historical_skus),
                "never_listed_active_skus_count": len(never_listed_skus),
                "coverage_active_pct": round((len(current_skus) / len(active_city_skus) * 100.0), 2) if active_city_skus else None,
                "potential_loss_now_kzt": round(potential_loss_now, 2),
                "lost_historical_value_kzt": round(lost_value_now, 2),
                "top_missing_skus": top_missing[:12],
            }
        )

    pharmacy_rows.sort(
        key=lambda row: (
            -(row.get("potential_loss_now_kzt") or 0),
            -(row.get("lost_historical_skus_count") or 0),
            str(row.get("pharmacy") or ""),
        )
    )

    all_pharmacy_slugs = {str(row["slug"]) for _, row in pharmacies_df.iterrows()}
    sku_rows: list[dict[str, Any]] = []
    for sku in sorted(portfolio_skus | active_city_skus):
        sku_meta = sku_catalog_map.get(sku)
        sku_stat = city_sku_map.get(sku)
        current_pharmacies = {venue_slug for venue_slug, canonical_sku in current_map if canonical_sku == sku}
        historical_pharmacies = {venue_slug for venue_slug, canonical_sku in historical_map if canonical_sku == sku}
        lost_pharmacies = historical_pharmacies - current_pharmacies
        missing_pharmacies_now = all_pharmacy_slugs - current_pharmacies
        top_missing_pharmacies = []
        for slug in sorted(lost_pharmacies if lost_pharmacies else missing_pharmacies_now):
            meta = pharmacies_index.get(slug, {})
            top_missing_pharmacies.append(
                {
                    "pharmacy": str(meta.get("name") or slug),
                    "venue_slug": slug,
                    "address": str(meta.get("address") or ""),
                }
            )
        current_min = safe_float(sku_stat.get("current_min_price") if sku_stat is not None else None)
        sku_rows.append(
            {
                "canonical_sku": sku,
                "canonical_name": str((sku_meta or {}).get("canonical_name") or (sku_stat or {}).get("canonical_name") or sku),
                "brand_slug": str((sku_meta or {}).get("brand_slug") or (sku_stat or {}).get("brand_slug") or ""),
                "brand": str((sku_meta or {}).get("brand") or (sku_stat or {}).get("brand") or ""),
                "current_pharmacies_count": len(current_pharmacies),
                "historical_pharmacies_count": len(historical_pharmacies),
                "lost_pharmacies_count": len(lost_pharmacies),
                "missing_pharmacies_now_count": len(missing_pharmacies_now),
                "current_min_price_kzt": round(current_min, 2) if current_min is not None else None,
                "first_seen_at": safe_iso((sku_stat or {}).get("first_seen_at")),
                "last_seen_at": safe_iso((sku_stat or {}).get("last_seen_at")),
                "top_missing_pharmacies": top_missing_pharmacies[:12],
            }
        )

    sku_rows.sort(
        key=lambda row: (
            -(row.get("lost_pharmacies_count") or 0),
            -(row.get("missing_pharmacies_now_count") or 0),
            str(row.get("canonical_name") or ""),
        )
    )

    latest_days = {
        brand_slug: max(stamp_raw[:10] for _, slug, _, stamp_raw in latest_files if slug == brand_slug)
        for brand_slug in sorted({slug for _, slug, _, _ in latest_files})
    }
    summary = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "city_slug": "almaty",
        "pharmacies_total": len(pharmacy_rows),
        "portfolio_skus_total": len(portfolio_skus),
        "active_city_skus_total": len(active_city_skus),
        "currently_available_city_skus_total": len(current_available_city_skus),
        "current_listed_pairs_total": int(len(current_presence)),
        "historical_pairs_total": int(len(historical_presence)),
        "missing_active_pairs_total": int(sum(row["missing_active_skus_count"] for row in pharmacy_rows)),
        "missing_available_now_pairs_total": int(sum(row["missing_available_now_skus_count"] for row in pharmacy_rows)),
        "lost_historical_pairs_total": int(sum(row["lost_historical_skus_count"] for row in pharmacy_rows)),
        "potential_loss_now_total_kzt": round(sum(row["potential_loss_now_kzt"] for row in pharmacy_rows), 2),
        "catalog_file": catalog_file.name,
        "latest_brand_days": latest_days,
        "latest_files": [path.name for path, _, _, _ in latest_files],
        "history_files_count": len(all_history_files),
        "output_prefix": output_prefix,
    }
    return {
        "summary": summary,
        "pharmacy_rows": pharmacy_rows,
        "sku_rows": sku_rows,
        "gap_rows": gap_rows,
    }


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({field: row.get(field, "") for field in fieldnames})


def run() -> int:
    args = parse_args()
    results_dir = Path(args.results_dir).expanduser().resolve()
    state_dir = Path(args.state_dir).expanduser().resolve()
    portfolio_csv = Path(args.portfolio_csv).expanduser().resolve()
    pharmacies_catalog = Path(args.pharmacies_catalog).expanduser().resolve()

    enabled_brand_slugs, brand_display_map = load_enabled_portfolio(portfolio_csv)
    pharmacies_df = load_pharmacies_catalog(pharmacies_catalog, city_slug=args.city_slug)
    history_files = discover_report_files(results_dir, enabled_brand_slugs)
    latest_files = latest_snapshot_files(history_files)
    sku_catalog_df = load_sku_catalog(state_dir, enabled_brand_slugs, brand_display_map)
    historical_df = load_item_reports(history_files, brand_display_map)
    current_df = load_item_reports(latest_files, brand_display_map)
    if not sku_catalog_df.empty:
        portfolio_sku_set = set(sku_catalog_df["canonical_sku"].dropna().astype(str).str.strip())
        historical_df = historical_df[historical_df["canonical_sku"].isin(portfolio_sku_set)].copy()
        current_df = current_df[current_df["canonical_sku"].isin(portfolio_sku_set)].copy()
    current_df = dedupe_current_rows(current_df)
    if sku_catalog_df.empty:
        fallback = historical_df[["canonical_sku", "canonical_name", "brand_slug", "brand"]].copy()
        fallback["product_line"] = ""
        fallback["dosage_or_volume"] = ""
        fallback["pack_size"] = ""
        fallback["form_factor"] = ""
        sku_catalog_df = fallback.drop_duplicates("canonical_sku", keep="first")

    current_presence = aggregate_current_presence(current_df)
    historical_presence = aggregate_historical_presence(historical_df)
    city_sku_stats = compute_city_sku_stats(current_presence, historical_presence)
    payload = build_gap_payload(
        pharmacies_df=pharmacies_df,
        sku_catalog_df=sku_catalog_df,
        current_presence=current_presence,
        historical_presence=historical_presence,
        city_sku_stats=city_sku_stats,
        latest_files=latest_files,
        all_history_files=history_files,
        output_prefix=args.output_prefix,
        catalog_file=pharmacies_catalog,
    )

    stamp = datetime.now().strftime(STAMP_FMT)
    output_dir = results_dir
    json_path = output_dir / f"{args.output_prefix}_{stamp}.json"
    pharmacies_csv = output_dir / f"{args.output_prefix}_{stamp}_pharmacies.csv"
    skus_csv = output_dir / f"{args.output_prefix}_{stamp}_skus.csv"
    gaps_csv = output_dir / f"{args.output_prefix}_{stamp}_gaps.csv"

    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    write_csv(
        pharmacies_csv,
        payload["pharmacy_rows"],
        [
            "pharmacy",
            "venue_slug",
            "venue_url",
            "address",
            "lat",
            "lon",
            "portfolio_skus_count",
            "active_city_skus_count",
            "current_listed_skus_count",
            "current_in_stock_skus_count",
            "historical_skus_count",
            "out_of_stock_skus_count",
            "missing_active_skus_count",
            "missing_available_now_skus_count",
            "lost_historical_skus_count",
            "never_listed_active_skus_count",
            "coverage_active_pct",
            "potential_loss_now_kzt",
            "lost_historical_value_kzt",
        ],
    )
    write_csv(
        skus_csv,
        payload["sku_rows"],
        [
            "canonical_sku",
            "canonical_name",
            "brand_slug",
            "brand",
            "current_pharmacies_count",
            "historical_pharmacies_count",
            "lost_pharmacies_count",
            "missing_pharmacies_now_count",
            "current_min_price_kzt",
            "first_seen_at",
            "last_seen_at",
        ],
    )
    write_csv(
        gaps_csv,
        payload["gap_rows"],
        [
            "pharmacy",
            "venue_slug",
            "venue_url",
            "address",
            "canonical_sku",
            "canonical_name",
            "brand_slug",
            "brand",
            "gap_type",
            "available_now_in_city",
            "current_city_pharmacies_count",
            "current_city_in_stock_pharmacies_count",
            "reference_price_kzt",
            "last_seen_here_at",
        ],
    )

    print(f"Pharmacies: {len(payload['pharmacy_rows'])}")
    print(f"Portfolio SKUs: {payload['summary']['portfolio_skus_total']}")
    print(f"Active city SKUs: {payload['summary']['active_city_skus_total']}")
    print(f"Missing active pairs: {payload['summary']['missing_active_pairs_total']}")
    print(f"Potential loss now: {payload['summary']['potential_loss_now_total_kzt']:.2f} KZT")
    print(f"JSON: {json_path}")
    print(f"Pharmacies CSV: {pharmacies_csv}")
    print(f"SKUs CSV: {skus_csv}")
    print(f"Gaps CSV: {gaps_csv}")
    return 0


if __name__ == "__main__":
    raise SystemExit(run())
