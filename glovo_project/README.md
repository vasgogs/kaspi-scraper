# Glovo Project Workspace

This folder is the dedicated workspace for all Glovo tasks.

## Current policy
- Active scope: same brand portfolio as Wolt, but tracked through Glovo.
- Portfolio can include multiple query variants for one brand.
- Cumulative catalogs in `state/` are updated on each run.

## Layout
- `glovo_project/config/`:
  - `portfolio.csv` - active brand list
- `glovo_project/RESULTS/`:
  - pharmacy discovery reports
  - per-brand monitor outputs
- `glovo_project/state/`:
  - cumulative Glovo product catalogs and references
- `glovo_project/scripts/`:
  - helper scripts for portfolio runs

## Quick start
1. Discover/refresh pharmacies (Almaty):
   - `./venv/bin/python glovo_discover_pharmacies.py`
2. Run portfolio monitor:
   - `bash glovo_project/scripts/run_portfolio_brand_monitor.sh`

## Notes
- Pharmacy discovery uses the public Glovo category page for pharmacies in Almaty.
- Brand monitoring uses the Glovo store search endpoint per pharmacy.
- Canonical SKU references reuse the same brand normalization logic as the Wolt pipeline.
