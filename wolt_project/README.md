# Wolt Project Workspace

This folder is the dedicated workspace for all Wolt tasks.

## Current policy
- Active scope: Vitrum + expansion brands from `config/portfolio.csv`.
- Portfolio can include multiple query variants for one brand (for better recall).
- Cumulative catalogs in `state/` are continuously updated by each run.

## Layout
- `wolt_project/config/`:
  - `portfolio.csv` - active brand list (currently only Vitrum)
  - `wolt_positions.csv` - current monitoring positions
  - `wolt_positions.example.csv` - template
- `wolt_project/RESULTS/`:
  - all Wolt discovery/search/monitor outputs
- `wolt_project/state/`:
  - cumulative catalogs and references
- `wolt_project/scripts/`:
  - helper scripts for daily runs

## Quick start
1. Discover/refresh pharmacies (Almaty):
   - `bash wolt_project/scripts/run_vitrum_discovery.sh`
2. Run brand monitor (Vitrum-only) and update catalogs:
   - `bash wolt_project/scripts/run_vitrum_brand_monitor.sh`
3. Run item stock monitor from config:
   - `bash wolt_project/scripts/run_vitrum_stock_monitor.sh`
4. Run portfolio monitor (all enabled brands from `portfolio.csv`):
   - `bash wolt_project/scripts/run_portfolio_brand_monitor.sh`

## Portfolio expansion rule
- Add a new line to `wolt_project/config/portfolio.csv` with `enabled=1`.
- The next run of `run_portfolio_brand_monitor.sh` will include this brand automatically.
- For search variants (e.g., hyphen/no-hyphen), add extra rows with the same `brand` and different `query`.

## Generated references
After each brand run:
- detailed mapping: `wolt_project/state/wolt_<brand>_item_reference.csv`
- canonical SKU catalog: `wolt_project/state/wolt_<brand>_canonical_catalog.csv`
- low-confidence/unmapped rows: `wolt_project/state/wolt_<brand>_unmapped.csv`

Additionally for Vitrum (`run_vitrum_brand_monitor.sh`):
- detailed mapping: `wolt_project/state/wolt_vitrum_item_reference.csv`
- canonical SKU catalog: `wolt_project/state/wolt_vitrum_canonical_catalog.csv`
- unmapped items: `wolt_project/state/wolt_vitrum_unmapped.csv`
