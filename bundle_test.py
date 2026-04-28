#!/usr/bin/env python3
"""Тестовый расчёт выгодности бандлов из CSV без запуска бота."""

from __future__ import annotations

import math
import os
from pathlib import Path
from typing import Any

import pandas as pd
from playwright.sync_api import sync_playwright

from Scraper_Kaspi import (
    scrape_single_product,
    normalize_seller_name,
    seller_matches,
)


def _safe_int(val: Any) -> int | None:
    try:
        if pd.isna(val):
            return None
        s = str(val).replace(" ", "")
        return int(float(s))
    except Exception:
        return None


def fetch_base_price(
    page,
    url: str,
    city: str,
    seller_hint: str,
    exclude_sellers: list[str] | None = None,
) -> tuple[int | None, tuple[int | None, str | None], tuple[int | None, str | None]]:
    """Возвращаем (цену продавца, лучший по всем, лучший по исключениям)."""
    recs = scrape_single_product(page, url, input_name="base", city=city) or []
    target_norm = normalize_seller_name(seller_hint)
    exclude_norms = {normalize_seller_name(s) for s in (exclude_sellers or []) if s}
    prices = []
    best_overall = []
    best_filtered = []
    for rec in recs:
        seller = normalize_seller_name(rec.get("seller", ""))
        price = _safe_int(rec.get("price_kzt"))
        if price:
            best_overall.append((price, rec.get("seller")))
            if seller not in exclude_norms:
                best_filtered.append((price, rec.get("seller")))
        if seller_matches(seller, target_norm) and price:
            prices.append(price)
    best_overall_pair = min(best_overall, key=lambda x: x[0]) if best_overall else (None, None)
    best_filtered_pair = min(best_filtered, key=lambda x: x[0]) if best_filtered else (None, None)
    return (min(prices) if prices else None), best_overall_pair, best_filtered_pair


def run_bundle_check(
    csv_path: Path,
    limit_rows: int | None = None,
    output_path: Path | None = None,
    write_excel: bool = True,
) -> tuple[Path | None, int, pd.DataFrame]:
    """Scrape base SKU prices for bundle rows and (optionally) write an Excel report."""
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV not found: {csv_path}")

    df = pd.read_csv(csv_path)
    if "bundle_of" not in df.columns or "bundle_qty" not in df.columns:
        raise ValueError("Нет колонок bundle_of / bundle_qty")

    if limit_rows and limit_rows > 0:
        df = df.head(limit_rows)
    rows = df.to_dict(orient="records")
    results: list[dict[str, Any]] = []

    exclude_raw = os.environ.get("BUNDLE_EXCLUDE_SELLERS", "Аптека от А до Я;Аптека MSP")
    exclude_sellers = [s.strip() for s in exclude_raw.replace(",", ";").split(";") if s.strip()]

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--disable-dev-shm-usage"])
        context = browser.new_context(
            viewport={"width": 1600, "height": 900},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
        )
        for row in rows:
            bundle_url = str(row.get("bundle_of") or "").strip()
            bundle_qty = _safe_int(row.get("bundle_qty")) or 1
            if not bundle_url:
                fallback_url = str(row.get("product_url") or row.get("product_link") or "").strip()
                if fallback_url:
                    bundle_url = fallback_url
                    bundle_qty = 1
            seller_hint = str(row.get("seller") or "")
            city = str(row.get("region") or "Алматы")
            bundle_price = _safe_int(row.get("price_kzt"))
            base_price = None
            best_price = None
            best_seller = None
            note = ""
            if bundle_url:
                page = context.new_page()
                try:
                    base_price, best_overall, best_filtered = fetch_base_price(
                        page,
                        bundle_url,
                        city,
                        seller_hint,
                        exclude_sellers=exclude_sellers,
                    )
                    best_pair = best_filtered if best_filtered[0] is not None else best_overall
                    best_price, best_seller = best_pair
                    if base_price is None:
                        note = "Базовый SKU не найден"
                except Exception as exc:
                    exc_text = str(exc)
                    if "Kaspi returned error page" in exc_text or "ERR_NAME_NOT_RESOLVED" in exc_text:
                        note = "Каспи не дает парсить"
                    else:
                        note = "Нет данных по бандлам"
                    print(f"⚠️ Bundle base error for {bundle_url} ({city}): {exc}")
                finally:
                    page.close()
            else:
                note = "Нет данных по бандлам"
            bundle_unit = None
            delta = None
            delta_best = None
            if bundle_price is not None and bundle_qty:
                bundle_unit = math.ceil(bundle_price / bundle_qty)
            if bundle_unit is not None and base_price is not None:
                delta = bundle_unit - base_price
            if bundle_unit is not None and best_price is not None:
                best_total = best_price * bundle_qty
                delta_best = bundle_price - best_total
            out_row = dict(row)
            out_row.update({
                "bundle_qty": bundle_qty,
                "bundle_1_sku_price_kzt": base_price,
                "bundle_price_kzt": bundle_price,
                "bundle_unit_price_kzt": bundle_unit,
                "bundle_vs_single_kzt": delta,
                "bundle_best_seller": best_seller,
                "1_sku_best_price_kzt": best_price,
                "bundle_best_price_kzt": (best_price * bundle_qty) if best_price is not None else None,
                "bundle_vs_best_kzt": delta_best,
            })
            results.append(out_row)
        context.close()
        browser.close()

    out_df = pd.DataFrame(results)
    desired_order = [
        "region",
        "seller",
        "input_product",
        "product_url",
        "price_kzt",
        "bundle_of",
        "bundle_qty",
        "bundle_1_sku_price_kzt",
        "bundle_price_kzt",
        "bundle_unit_price_kzt",
        "bundle_vs_single_kzt",
        "bundle_best_seller",
        "1_sku_best_price_kzt",
        "bundle_best_price_kzt",
        "bundle_vs_best_kzt",
    ]
    for col in desired_order:
        if col not in out_df.columns:
            out_df[col] = None
    out_df = out_df.reindex(columns=desired_order)

    def _discount(row):
        best_total = row.get("bundle_best_price_kzt")
        bundle_price = row.get("bundle_price_kzt")
        try:
            if best_total and bundle_price:
                return round((best_total - bundle_price) / best_total * 100, 2)
        except Exception:
            return None
        return None

    out_df["bundle_discount_pct"] = out_df.apply(_discount, axis=1)
    out_df["bundle_discount_pct"] = out_df["bundle_discount_pct"].map(
        lambda v: (f"{float(v):.2f}%") if pd.notna(v) else None
    )

    out_path: Path | None
    if write_excel:
        out_path = output_path or Path("RESULTS/bundle_test.xlsx")
        if not out_path.is_absolute():
            out_path = Path(__file__).resolve().parent / out_path
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_df.to_excel(out_path, index=False)
    else:
        out_path = None if output_path is None else output_path
    return out_path, len(out_df), out_df


def _parse_int_env(raw: str | None) -> int | None:
    if not raw:
        return None
    raw = raw.strip()
    if not raw:
        return None
    try:
        value = int(raw)
        return value if value > 0 else None
    except Exception:
        return None


def main():
    base_dir = Path(__file__).resolve().parent
    csv_env = os.environ.get("BUNDLE_TEST_CSV", "миссия февраль_test1sku.csv")
    output_env = os.environ.get("BUNDLE_TEST_OUTPUT", "RESULTS/bundle_test.xlsx")
    limit_env = os.environ.get("BUNDLE_TEST_LIMIT_ROWS", "4")

    csv_path = Path(csv_env)
    if not csv_path.is_absolute():
        csv_path = base_dir / csv_path

    output_path = Path(output_env)
    if not output_path.is_absolute():
        output_path = base_dir / output_path

    limit_rows = _parse_int_env(limit_env)
    out_path, rows_count, _ = run_bundle_check(
        csv_path,
        limit_rows=limit_rows,
        output_path=output_path,
        write_excel=True,
    )
    print(f"✅ Done. Saved {out_path} (rows: {rows_count})")


if __name__ == "__main__":
    main()
