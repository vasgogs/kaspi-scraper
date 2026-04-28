#!/usr/bin/env python3
"""Build time-weighted campaign quality KPIs from mission snapshots."""

from __future__ import annotations

import argparse
import json
import os
import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd

PRICE_TOLERANCE_KZT = int(os.environ.get("PRICE_TOLERANCE_KZT", "5"))
STAMP_RE = re.compile(r"^(?P<prefix>.+)_(?P<date>\d{4}-\d{2}-\d{2})_(?P<h>\d{2})-(?P<m>\d{2})-(?P<s>\d{2})\.xlsx$")
PRODUCT_CODE_RE = re.compile(r"/p/[^/]*?-(\d+)(?:/|$)")
ANALYSIS_OFFER_RE = re.compile(
    r"^(?:(?:Следующий|После нас):\s*)?(?P<seller>.+?)\s+—\s+(?P<price>[\d\s]+)\s*₸",
    re.IGNORECASE,
)
MSP_SELLER_PATTERN = re.compile(r"^аптека\s*msp(?:\s+(?:алматы|астана|шымкент))?$", re.IGNORECASE)
MISSION_PARTNER_SELLERS_DEFAULT = (
    "Аптека от А до Я;"
    "Аптека MSP;"
    "Аптека MSP Алматы;"
    "Аптека MSP Астана;"
    "Аптека MSP Шымкент;"
    "ФАРМАКОМ"
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Analyze mission campaign quality by time.")
    parser.add_argument("--results-dir", default="RESULTS", help="Directory with mission XLSX files")
    parser.add_argument("--prefix", default=os.environ.get("MISSION_FILE_PREFIX", "mission_april"), help="Mission file prefix")
    parser.add_argument("--campaign-start", required=True, help="Campaign start date YYYY-MM-DD")
    parser.add_argument("--campaign-end", required=True, help="Campaign end date YYYY-MM-DD")
    parser.add_argument(
        "--output-prefix",
        default="mission_campaign_quality",
        help="Output file prefix inside RESULTS",
    )
    return parser.parse_args()


def parse_campaign_date(value: str) -> datetime:
    return datetime.strptime(value.strip(), "%Y-%m-%d")


def parse_snapshot_stamp(path: Path, prefix: str) -> datetime | None:
    match = STAMP_RE.match(path.name)
    if not match:
        return None
    if match.group("prefix") != prefix:
        return None
    return datetime.strptime(
        f"{match.group('date')} {match.group('h')}:{match.group('m')}:{match.group('s')}",
        "%Y-%m-%d %H:%M:%S",
    )


def normalize_text(value: Any) -> str:
    return " ".join(str(value or "").replace("\u00a0", " ").strip().split())


def normalize_seller_name(name: Any) -> str:
    text = normalize_text(name)
    if not text:
        return ""
    lookalikes = str.maketrans(
        {
            "a": "а", "A": "а",
            "e": "е", "E": "е",
            "o": "о", "O": "о",
            "p": "р", "P": "р",
            "c": "с", "C": "с",
            "x": "х", "X": "х",
            "y": "у", "Y": "у",
            "k": "к", "K": "к",
            "h": "н", "H": "н",
            "b": "в", "B": "в",
            "m": "м", "M": "м",
            "t": "т", "T": "т",
        }
    )
    return re.sub(r"\s+", " ", text.translate(lookalikes)).lower()


def seller_matches(actual_norm: str, expected_norm: str) -> bool:
    if not expected_norm:
        return bool(actual_norm)
    if actual_norm == expected_norm:
        return True
    return expected_norm in actual_norm or actual_norm in expected_norm


def parse_seller_list(raw: str | None, default_raw: str = "") -> list[str]:
    text = raw if raw is not None else default_raw
    return [s.strip() for s in re.split(r"[;,/|]", text or "") if s.strip()]


def seller_in_list(seller_name: Any, sellers: list[str] | tuple[str, ...] | set[str]) -> bool:
    actual_norm = normalize_seller_name(seller_name)
    if not actual_norm:
        return False
    if MSP_SELLER_PATTERN.match(normalize_text(seller_name)):
        actual_norm = normalize_seller_name("Аптека MSP")
    for candidate in sellers:
        candidate_norm = normalize_seller_name(candidate)
        if seller_matches(actual_norm, candidate_norm):
            return True
    return False


def normalize_status(value: Any) -> str:
    return normalize_text(value).lower()


def extract_product_code(url: Any) -> str:
    text = str(url or "").strip()
    if not text:
        return ""
    match = PRODUCT_CODE_RE.search(text)
    return match.group(1) if match else ""


def safe_div(num: float, den: float) -> float | None:
    if not den:
        return None
    return num / den


def pct(num: float, den: float) -> float | None:
    value = safe_div(num, den)
    if value is None:
        return None
    return round(value * 100.0, 2)


def format_iso(dt: datetime | None) -> str:
    return dt.strftime("%Y-%m-%d %H:%M:%S") if dt else ""


def parse_offer_analysis(value: Any) -> tuple[str | None, int | None]:
    text = normalize_text(value)
    if not text:
        return (None, None)
    first_line = str(value or "").splitlines()[0].strip()
    if not first_line:
        return (None, None)
    match = ANALYSIS_OFFER_RE.search(first_line)
    if not match:
        return (None, None)
    digits = re.sub(r"[^\d]", "", match.group("price") or "")
    return (normalize_text(match.group("seller")), int(digits) if digits else None)


def resolve_external_market_offer(row: pd.Series, partner_sellers: list[str]) -> tuple[int | None, str | None]:
    actual_seller = row.get("seller") or ""
    candidates = [
        parse_offer_analysis(row.get("better_price_analysis")),
        parse_offer_analysis(row.get("second_price_analysis")),
    ]
    for seller_name, price_kzt in candidates:
        if not seller_name or price_kzt is None:
            continue
        if seller_in_list(seller_name, partner_sellers):
            continue
        return (price_kzt, seller_name)
    fallback_price = pd.to_numeric(row.get("best_price_kzt"), errors="coerce")
    fallback_seller, _fallback_price_from_text = parse_offer_analysis(row.get("better_price_analysis"))
    if pd.notna(fallback_price) and fallback_seller and not seller_in_list(fallback_seller, partner_sellers):
        return (int(fallback_price), fallback_seller)
    if pd.notna(fallback_price) and actual_seller and not seller_in_list(actual_seller, partner_sellers):
        return (int(fallback_price), normalize_text(actual_seller))
    return (None, None)


def collect_files(results_dir: Path, prefix: str, start_dt: datetime, end_exclusive: datetime) -> list[tuple[Path, datetime]]:
    rows: list[tuple[Path, datetime]] = []
    for path in results_dir.glob(f"{prefix}_*.xlsx"):
        stamp = parse_snapshot_stamp(path, prefix)
        if not stamp:
            continue
        if stamp < start_dt or stamp >= end_exclusive:
            continue
        rows.append((path, stamp))
    rows.sort(key=lambda item: item[1])
    return rows


def build_metric_flags(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()
    partner_sellers = parse_seller_list(os.environ.get("MISSION_PARTNER_SELLERS"), MISSION_PARTNER_SELLERS_DEFAULT)
    work["status_norm"] = work["status"].map(normalize_status)
    work["price_kzt"] = pd.to_numeric(work["price_kzt"], errors="coerce")
    work["actual_price_kzt"] = pd.to_numeric(work["actual_price_kzt"], errors="coerce")
    work["best_price_kzt"] = pd.to_numeric(work["best_price_kzt"], errors="coerce")
    external_market = work.apply(lambda row: resolve_external_market_offer(row, partner_sellers), axis=1)
    work["market_best_price_kzt"] = [item[0] for item in external_market]
    work["market_best_seller"] = [item[1] for item in external_market]
    unavailable_statuses = {
        "продавец отсутствует",
        "нет продавцов на карточке",
        "нет цены",
        "ошибка скрейпа",
    }
    work["is_live"] = work["actual_price_kzt"].notna() & ~work["status_norm"].isin(unavailable_statuses)
    work["has_external_market"] = work["market_best_price_kzt"].notna()
    work["is_correct_price"] = (
        work["is_live"]
        & work["price_kzt"].notna()
        & ((work["actual_price_kzt"] - work["price_kzt"]).abs() <= PRICE_TOLERANCE_KZT)
    )
    work["is_best_price"] = (
        work["is_live"]
        & work["market_best_price_kzt"].notna()
        & ((work["actual_price_kzt"] - work["market_best_price_kzt"]) <= PRICE_TOLERANCE_KZT)
    )
    work["is_attention"] = (
        work["is_live"]
        & work["market_best_price_kzt"].notna()
        & ((work["actual_price_kzt"] - work["market_best_price_kzt"]) > PRICE_TOLERANCE_KZT)
    )
    work["is_problem"] = work["status_norm"].map(
        lambda text: text.startswith("проблема")
        or text.startswith("дороже")
        or text.startswith("дешевле")
        or "нет" in text
        or "отсутств" in text
        or "ошибка" in text
    )
    return work


def add_duration_columns(work: pd.DataFrame) -> pd.DataFrame:
    work = work.copy()
    work["observed_hours"] = work["duration_hours"]
    work["live_hours"] = work["duration_hours"] * work["is_live"].astype(float)
    work["market_hours"] = work["duration_hours"] * work["has_external_market"].astype(float)
    work["correct_price_hours"] = work["duration_hours"] * work["is_correct_price"].astype(float)
    work["best_price_hours"] = work["duration_hours"] * work["is_best_price"].astype(float)
    work["attention_hours"] = work["duration_hours"] * work["is_attention"].astype(float)
    work["problem_hours"] = work["duration_hours"] * work["is_problem"].astype(float)
    work["snapshots"] = 1
    work["live_snapshots"] = work["is_live"].astype(int)
    work["market_snapshots"] = work["has_external_market"].astype(int)
    work["correct_price_snapshots"] = work["is_correct_price"].astype(int)
    work["best_price_snapshots"] = work["is_best_price"].astype(int)
    return work


def summarize_group(work: pd.DataFrame, group_cols: list[str], label_cols: list[str]) -> pd.DataFrame:
    aggregations: dict[str, str] = {
        "observed_hours": "sum",
        "live_hours": "sum",
        "market_hours": "sum",
        "correct_price_hours": "sum",
        "best_price_hours": "sum",
        "attention_hours": "sum",
        "problem_hours": "sum",
        "snapshots": "sum",
        "live_snapshots": "sum",
        "market_snapshots": "sum",
        "correct_price_snapshots": "sum",
        "best_price_snapshots": "sum",
        "snapshot_ts": "min",
        "snapshot_end_ts": "max",
    }
    for col in label_cols:
        aggregations[col] = "first"
    grouped = work.groupby(group_cols, dropna=False).agg(aggregations).reset_index()
    grouped["availability_pct_observed"] = grouped.apply(lambda row: pct(row["live_hours"], row["observed_hours"]), axis=1)
    grouped["correct_price_pct_observed"] = grouped.apply(lambda row: pct(row["correct_price_hours"], row["observed_hours"]), axis=1)
    grouped["correct_price_pct_live"] = grouped.apply(lambda row: pct(row["correct_price_hours"], row["live_hours"]), axis=1)
    grouped["best_price_pct_observed"] = grouped.apply(lambda row: pct(row["best_price_hours"], row["observed_hours"]), axis=1)
    grouped["best_price_pct_live"] = grouped.apply(lambda row: pct(row["best_price_hours"], row["live_hours"]), axis=1)
    grouped["best_price_pct_market"] = grouped.apply(lambda row: pct(row["best_price_hours"], row["market_hours"]), axis=1)
    grouped["attention_pct_observed"] = grouped.apply(lambda row: pct(row["attention_hours"], row["observed_hours"]), axis=1)
    grouped["problem_pct_observed"] = grouped.apply(lambda row: pct(row["problem_hours"], row["observed_hours"]), axis=1)
    grouped["first_seen_snapshot"] = grouped["snapshot_ts"].map(format_iso)
    grouped["last_seen_snapshot"] = grouped["snapshot_end_ts"].map(format_iso)
    grouped = grouped.drop(columns=["snapshot_ts", "snapshot_end_ts"])
    return grouped


def load_snapshot_rows(files: list[tuple[Path, datetime]]) -> tuple[pd.DataFrame, list[datetime]]:
    snapshot_times = [stamp for _path, stamp in files]
    next_map: dict[datetime, datetime] = {}
    for idx, stamp in enumerate(snapshot_times[:-1]):
        next_map[stamp] = snapshot_times[idx + 1]
    frames: list[pd.DataFrame] = []
    for path, stamp in files:
        next_stamp = next_map.get(stamp)
        duration_hours = max(0.0, ((next_stamp - stamp).total_seconds() / 3600.0) if next_stamp else 0.0)
        df = pd.read_excel(path)
        required = {"region", "seller", "input_product", "product_url", "price_kzt", "actual_price_kzt", "status", "best_price_kzt"}
        required_with_analysis = required | {"better_price_analysis", "second_price_analysis"}
        if not required.issubset(df.columns):
            continue
        keep_cols = required_with_analysis | {"product", "scraped_at"}
        frame = df.loc[:, [col for col in df.columns if col in keep_cols]].copy()
        frame["snapshot_ts"] = stamp
        frame["snapshot_end_ts"] = next_stamp or stamp
        frame["duration_hours"] = duration_hours
        frame["region"] = frame["region"].map(normalize_text)
        frame["seller"] = frame["seller"].map(normalize_text)
        frame["input_product"] = frame["input_product"].map(normalize_text)
        frame["product"] = frame.get("product", "").map(normalize_text)
        frame["product_url"] = frame["product_url"].fillna("").astype(str).str.strip()
        frame["product_code"] = frame["product_url"].map(extract_product_code)
        frame["row_key"] = frame.apply(
            lambda row: "|".join(
                [
                    normalize_text(row.get("region")),
                    normalize_text(row.get("seller")),
                    str(row.get("product_url") or "").strip() or normalize_text(row.get("input_product")),
                ]
            ),
            axis=1,
        )
        frames.append(frame)
    if not frames:
        return pd.DataFrame(), snapshot_times
    return pd.concat(frames, ignore_index=True), snapshot_times


def build_summary_payload(
    work: pd.DataFrame,
    files: list[tuple[Path, datetime]],
    requested_start: datetime,
    requested_end: datetime,
) -> dict[str, Any]:
    partner_sellers = parse_seller_list(os.environ.get("MISSION_PARTNER_SELLERS"), MISSION_PARTNER_SELLERS_DEFAULT)
    coverage_start = files[0][1] if files else None
    coverage_end = files[-1][1] if files else None
    observed_hours = (coverage_end - coverage_start).total_seconds() / 3600.0 if coverage_start and coverage_end else 0.0
    total_track_hours = float(work["observed_hours"].sum()) if not work.empty else 0.0
    live_hours = float(work["live_hours"].sum()) if not work.empty else 0.0
    market_hours = float(work["market_hours"].sum()) if not work.empty else 0.0
    correct_price_hours = float(work["correct_price_hours"].sum()) if not work.empty else 0.0
    best_price_hours = float(work["best_price_hours"].sum()) if not work.empty else 0.0
    attention_hours = float(work["attention_hours"].sum()) if not work.empty else 0.0
    problem_hours = float(work["problem_hours"].sum()) if not work.empty else 0.0
    return {
        "campaign_requested_start": requested_start.strftime("%Y-%m-%d"),
        "campaign_requested_end": requested_end.strftime("%Y-%m-%d"),
        "data_coverage_start": format_iso(coverage_start),
        "data_coverage_end": format_iso(coverage_end),
        "snapshots_used": len(files),
        "tracked_rows": int(work["row_key"].nunique()) if not work.empty else 0,
        "observed_window_hours": round(observed_hours, 2),
        "row_hours_total": round(total_track_hours, 2),
        "live_hours_total": round(live_hours, 2),
        "market_hours_total": round(market_hours, 2),
        "correct_price_hours_total": round(correct_price_hours, 2),
        "best_price_hours_total": round(best_price_hours, 2),
        "attention_hours_total": round(attention_hours, 2),
        "problem_hours_total": round(problem_hours, 2),
        "availability_pct_observed": pct(live_hours, total_track_hours),
        "correct_price_pct_observed": pct(correct_price_hours, total_track_hours),
        "correct_price_pct_live": pct(correct_price_hours, live_hours),
        "best_price_pct_observed": pct(best_price_hours, total_track_hours),
        "best_price_pct_live": pct(best_price_hours, live_hours),
        "best_price_pct_market": pct(best_price_hours, market_hours),
        "attention_pct_observed": pct(attention_hours, total_track_hours),
        "problem_pct_observed": pct(problem_hours, total_track_hours),
        "price_tolerance_kzt": PRICE_TOLERANCE_KZT,
        "best_price_scope": "external_market_only",
        "partner_sellers": partner_sellers,
        "files": [path.name for path, _stamp in files],
    }


def main():
    args = parse_args()
    results_dir = Path(args.results_dir).resolve()
    campaign_start = parse_campaign_date(args.campaign_start)
    campaign_end = parse_campaign_date(args.campaign_end)
    campaign_end_exclusive = campaign_end + timedelta(days=1)
    files = collect_files(results_dir, args.prefix, campaign_start, campaign_end_exclusive)
    if len(files) < 2:
        raise SystemExit("Недостаточно mission snapshot-файлов в выбранном окне для time-weighted анализа.")

    raw_df, _snapshot_times = load_snapshot_rows(files)
    if raw_df.empty:
        raise SystemExit("Не удалось собрать строки из mission snapshot-файлов.")

    work = add_duration_columns(build_metric_flags(raw_df))
    per_row = summarize_group(
        work,
        group_cols=["row_key"],
        label_cols=["region", "seller", "input_product", "product", "product_url", "product_code", "price_kzt"],
    ).sort_values(
        by=["problem_pct_observed", "correct_price_pct_observed", "best_price_pct_observed", "seller", "region", "input_product"],
        ascending=[False, False, False, True, True, True],
    )
    per_seller = summarize_group(
        work,
        group_cols=["seller"],
        label_cols=[],
    ).sort_values(by=["problem_pct_observed", "seller"], ascending=[False, True])
    per_region = summarize_group(
        work,
        group_cols=["region"],
        label_cols=[],
    ).sort_values(by=["region"])
    per_seller_region = summarize_group(
        work,
        group_cols=["seller", "region"],
        label_cols=[],
    ).sort_values(by=["seller", "region"])
    per_sku = summarize_group(
        work,
        group_cols=["product_code", "input_product"],
        label_cols=["product"],
    ).sort_values(by=["problem_pct_observed", "input_product"], ascending=[False, True])

    summary = build_summary_payload(work, files, campaign_start, campaign_end)

    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    base_name = f"{args.output_prefix}_{args.campaign_start}_{args.campaign_end}_{stamp}"
    summary_path = results_dir / f"{base_name}_summary.json"
    per_row_path = results_dir / f"{base_name}_per_row.csv"
    per_seller_path = results_dir / f"{base_name}_per_seller.csv"
    per_region_path = results_dir / f"{base_name}_per_region.csv"
    per_seller_region_path = results_dir / f"{base_name}_per_seller_region.csv"
    per_sku_path = results_dir / f"{base_name}_per_sku.csv"

    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    per_row.to_csv(per_row_path, index=False, encoding="utf-8-sig")
    per_seller.to_csv(per_seller_path, index=False, encoding="utf-8-sig")
    per_region.to_csv(per_region_path, index=False, encoding="utf-8-sig")
    per_seller_region.to_csv(per_seller_region_path, index=False, encoding="utf-8-sig")
    per_sku.to_csv(per_sku_path, index=False, encoding="utf-8-sig")

    print(json.dumps({
        "summary_file": summary_path.name,
        "per_row_file": per_row_path.name,
        "per_seller_file": per_seller_path.name,
        "per_region_file": per_region_path.name,
        "per_seller_region_file": per_seller_region_path.name,
        "per_sku_file": per_sku_path.name,
        "summary": summary,
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
