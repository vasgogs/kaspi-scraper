#!/usr/bin/env python3
"""Rebuild Wolt brand references/canonical catalogs from existing item_id catalog."""

from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parents[2]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import wolt_brand_search_monitor as monitor


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Rebuild Wolt references from existing catalog")
    parser.add_argument(
        "--catalog",
        default=str(monitor.DEFAULT_ITEMS_CATALOG),
        help="Path to wolt_item_ids_catalog.csv",
    )
    parser.add_argument(
        "--portfolio",
        default=str(monitor.WOLT_PROJECT_DIR / "config" / "portfolio.csv"),
        help="Path to portfolio.csv with brand list",
    )
    return parser.parse_args()


def load_brands(path: Path) -> list[str]:
    if not path.exists():
        raise FileNotFoundError(f"Portfolio CSV not found: {path}")
    seen: set[str] = set()
    ordered: list[str] = []
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            brand = (row.get("brand") or "").strip()
            if not brand:
                continue
            key = monitor.normalize_text(brand)
            if key in seen:
                continue
            seen.add(key)
            ordered.append(brand)
    if not ordered:
        raise ValueError(f"No brands found in {path}")
    return ordered


def main() -> int:
    args = parse_args()
    catalog_path = Path(args.catalog).expanduser().resolve()
    portfolio_path = Path(args.portfolio).expanduser().resolve()
    state_dir = catalog_path.parent

    catalog = monitor.load_catalog(catalog_path)
    brands = load_brands(portfolio_path)
    monitor.export_active_ingredient_reference(state_dir / monitor.DEFAULT_ACTIVE_INGREDIENT_REFERENCE.name)

    for rec in catalog.values():
        rec_brand = str(rec.get("brand") or "").strip()
        rec_name = str(rec.get("item_name") or "").strip()
        if not rec_brand or not rec_name:
            continue
        canonical = monitor.canonicalize_brand_name(rec_brand, rec_name)
        rec["canonical_sku"] = canonical.get("canonical_sku", "")
        rec["canonical_name"] = canonical.get("canonical_name", "")
        rec["canonical_line"] = canonical.get("product_line", "")
        rec["canonical_pack_size"] = canonical.get("pack_size", "")
        rec["canonical_form_factor"] = canonical.get("form_factor", "")
        rec["canonical_dosage_or_volume"] = canonical.get("dosage_or_volume", "")
        rec["canonical_flavor"] = canonical.get("flavor", "")
        rec["canonical_active_ingredient"] = canonical.get("active_ingredient", "")
        rec["canonical_confidence"] = canonical.get("confidence", "")
        rec["canonical_rule"] = canonical.get("rule", "")

    total_rows = 0
    for brand in brands:
        brand_norm = monitor.normalize_text(brand)
        if brand_norm == "vitrum" or "витрум" in brand_norm:
            out = monitor.export_vitrum_reference(
                catalog=catalog,
                reference_path=state_dir / "wolt_vitrum_item_reference.csv",
                canonical_path=state_dir / "wolt_vitrum_canonical_catalog.csv",
                unmapped_path=state_dir / "wolt_vitrum_unmapped.csv",
            )
        else:
            safe_brand = monitor.slugify_token(brand) or "brand"
            out = monitor.export_brand_reference(
                catalog=catalog,
                brand=brand,
                reference_path=state_dir / f"wolt_{safe_brand}_item_reference.csv",
                canonical_path=state_dir / f"wolt_{safe_brand}_canonical_catalog.csv",
                unmapped_path=state_dir / f"wolt_{safe_brand}_unmapped.csv",
            )
        total_rows += int(out.get("detailed_count") or 0)
        print(
            f"{brand}: rows={out['detailed_count']} sku={out['canonical_count']} unmapped={out['unmapped_count']}"
        )

    monitor.save_catalog(catalog_path, catalog)
    print(f"Updated catalog: {catalog_path}")
    print(f"Total reference rows: {total_rows}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
