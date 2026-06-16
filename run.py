#!/usr/bin/env python3
"""
run.py — ASX Coverage Toolkit entry point

Usage
-----
    python run.py                          # full pipeline (fetch + enrich + format)
    python run.py --no-fetch               # skip yfinance, use cache / Research Data only
    python run.py --refresh                # force re-fetch all tickers
    python run.py --tickers CBA,BHP,ANZ   # refresh specific tickers only
    python run.py --output my_output.xlsx  # custom output path
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

# ── Paths ─────────────────────────────────────────────────────────────────────
ROOT        = Path(__file__).parent
DATA_DIR    = ROOT / "data"
OUTPUT_DIR  = ROOT / "output"
SOURCE_CSV  = DATA_DIR / "ASX_Entities_Enriched.csv"
CACHE_PATH  = DATA_DIR / "yfinance_cache.json"
DEFAULT_OUT = OUTPUT_DIR / "ASX_Master_Database.xlsx"


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="ASX Coverage Toolkit — build the master database")
    p.add_argument("--no-fetch",  action="store_true", help="skip yfinance fetch, use cached / Research Data only")
    p.add_argument("--refresh",   action="store_true", help="force re-fetch from yfinance (bypass cache)")
    p.add_argument("--tickers",   type=str, default="", help="comma-separated tickers to refresh (default: all)")
    p.add_argument("--output",    type=Path, default=DEFAULT_OUT, help="output .xlsx path")
    return p.parse_args()


def main() -> int:
    args = parse_args()

    # Validate source file
    if not SOURCE_CSV.exists():
        logger.error("Source CSV not found: %s", SOURCE_CSV)
        logger.error("Place ASX_Entities_Enriched.csv in the data/ directory.")
        return 1

    # ── Step 1: Live data — fetch, or load the committed cache offline ────────
    live_data: dict | None = None

    if args.no_fetch:
        from src.fetch_live import load_cache
        live_data = load_cache(CACHE_PATH)
        logger.info("Loaded %d tickers from committed cache (no fetch)", len(live_data))
    else:
        try:
            import pandas as pd

            from src.fetch_live import get_live_data

            src = pd.read_csv(SOURCE_CSV)
            all_tickers = src["ASX code"].astype(str).str.strip().tolist()

            if args.tickers:
                target_tickers = [t.strip().upper() for t in args.tickers.split(",")]
                logger.info("Refreshing %d specific tickers: %s", len(target_tickers), target_tickers)
            else:
                target_tickers = all_tickers

            live_data = get_live_data(
                tickers=target_tickers,
                cache_path=CACHE_PATH,
                force_refresh=args.refresh,
            )
            logger.info("Live data loaded for %d tickers", len(live_data))

        except ImportError as e:
            logger.warning("yfinance not available (%s) — falling back to committed cache", e)
            from src.fetch_live import load_cache
            live_data = load_cache(CACHE_PATH)
            logger.info("Loaded %d tickers from committed cache", len(live_data))

    # ── Step 2: Enrich ────────────────────────────────────────────────────────
    logger.info("Running enrichment pipeline...")
    from src.enrich import enrich
    df = enrich(source_csv=SOURCE_CSV, live_data=live_data)
    logger.info("Enriched %d rows, %d columns", len(df), len(df.columns))

    # ── Step 3: Format Excel ──────────────────────────────────────────────────
    logger.info("Formatting Excel workbook...")
    from src.format_excel import build_workbook
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    build_workbook(df=df, output_path=args.output)
    logger.info("Saved to %s", args.output)

    return 0


if __name__ == "__main__":
    sys.exit(main())
