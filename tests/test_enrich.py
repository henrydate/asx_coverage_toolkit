"""
tests/test_enrich.py
─────────────────────
Smoke tests for the enrichment pipeline.
Run with: python -m pytest tests/
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

SOURCE_CSV = Path(__file__).parent.parent / "data" / "ASX_Entities_Enriched.csv"

VALID_BHS      = {"Buy", "Hold", "Sell", ""}
VALID_PROGRESS = {"TA", "FA", "TA + FA", ""}
VALID_SOURCES  = {"Research Data", "Live Data", "Sector Estimate", "Source CSV"}
EXPECTED_COLS  = {
    "Company", "ASX Ticker", "Entity Type", "Market Cap Category",
    "GICS Sector", "GICS Industry Group", "GICS Industry", "GICS Sub-Industry",
    "Market Cap (AUD)", "Market Cap Data Date", "Share Price (AUD)", "Shares Outstanding",
    "Revenue (AUD, Latest Annual)", "Net Profit/Loss (AUD)", "Dividend Yield", "Franking Level",
    "P/E Ratio", "Price/Book", "Beta (vs ASX 200)", "Return on Equity (ROE)", "Debt/Equity Ratio",
    "Reporting Currency", "Fiscal Year End", "CEO / MD", "Website", "Year Founded",
    "Business Description", "HQ State / Territory", "Primary Country of Operations", "Domicile",
    "Primary Exchange", "Dual Listed", "Dual Listing Exchange", "Listing Status",
    "ASX 20", "ASX 50", "ASX 100", "ASX 200", "All Ordinaries", "Index Membership",
    "Data Source", "BHS Rating", "Progress", "Notes",
}


@pytest.fixture(scope="module")
def enriched_df() -> pd.DataFrame:
    """Run the enrichment pipeline once for all tests."""
    if not SOURCE_CSV.exists():
        pytest.skip(f"Source CSV not found: {SOURCE_CSV}")
    from src.enrich import enrich
    return enrich(source_csv=SOURCE_CSV)


def test_row_count(enriched_df: pd.DataFrame) -> None:
    """All source entities must be present in output."""
    src = pd.read_csv(SOURCE_CSV)
    assert len(enriched_df) == len(src), (
        f"Expected {len(src)} rows, got {len(enriched_df)}"
    )


def test_no_completely_blank_rows(enriched_df: pd.DataFrame) -> None:
    """No row should be entirely empty."""
    blank_rows = enriched_df.isnull().all(axis=1)
    assert not blank_rows.any(), f"Found {blank_rows.sum()} completely blank rows"


def test_required_columns_present(enriched_df: pd.DataFrame) -> None:
    """All 44 expected columns must be present."""
    missing = EXPECTED_COLS - set(enriched_df.columns)
    assert not missing, f"Missing columns: {missing}"


def test_bhs_rating_values(enriched_df: pd.DataFrame) -> None:
    """BHS Rating column must only contain valid values."""
    invalid = set(enriched_df["BHS Rating"].unique()) - VALID_BHS
    assert not invalid, f"Invalid BHS Rating values: {invalid}"


def test_progress_values(enriched_df: pd.DataFrame) -> None:
    """Progress column must only contain valid values."""
    invalid = set(enriched_df["Progress"].unique()) - VALID_PROGRESS
    assert not invalid, f"Invalid Progress values: {invalid}"


def test_data_source_values(enriched_df: pd.DataFrame) -> None:
    """Data Source column must only contain valid values."""
    invalid = set(enriched_df["Data Source"].unique()) - VALID_SOURCES
    assert not invalid, f"Invalid Data Source values: {invalid}"


def test_cba_mega_cap(enriched_df: pd.DataFrame) -> None:
    """CBA must be classified as Mega Cap."""
    cba = enriched_df[enriched_df["ASX Ticker"] == "CBA"]
    assert not cba.empty, "CBA not found in output"
    assert cba.iloc[0]["Market Cap Category"] == "Mega Cap", (
        f"CBA Market Cap Category was {cba.iloc[0]['Market Cap Category']}"
    )


def test_cba_in_asx20(enriched_df: pd.DataFrame) -> None:
    """CBA must be in the ASX 20."""
    cba = enriched_df[enriched_df["ASX Ticker"] == "CBA"].iloc[0]
    assert cba["ASX 20"] == "Yes"
    assert cba["ASX 50"] == "Yes"
    assert cba["ASX 100"] == "Yes"
    assert cba["ASX 200"] == "Yes"


def test_index_membership_consistency(enriched_df: pd.DataFrame) -> None:
    """
    Any company in ASX 20 must also be in ASX 50, 100, 200, All Ords.
    Any company in ASX 50 must also be in ASX 100, 200.
    """
    in_20 = enriched_df[enriched_df["ASX 20"] == "Yes"]
    assert (in_20["ASX 50"] == "Yes").all(), "ASX 20 member not in ASX 50"
    assert (in_20["ASX 100"] == "Yes").all(), "ASX 20 member not in ASX 100"
    assert (in_20["ASX 200"] == "Yes").all(), "ASX 20 member not in ASX 200"


def test_gics_sector_values(enriched_df: pd.DataFrame) -> None:
    """GICS Sector must only contain valid values."""
    valid_sectors = {
        "Financials", "Materials", "Energy", "Information Technology",
        "Health Care", "Industrials", "Consumer Discretionary", "Consumer Staples",
        "Communication Services", "Real Estate", "Utilities", "Unclassified",
    }
    invalid = set(enriched_df["GICS Sector"].unique()) - valid_sectors
    assert not invalid, f"Invalid GICS Sector values: {invalid}"


def test_no_empty_ticker(enriched_df: pd.DataFrame) -> None:
    """No row should have an empty ticker."""
    empty = enriched_df["ASX Ticker"].str.strip() == ""
    assert not empty.any(), f"{empty.sum()} rows have empty ASX Ticker"


def test_research_data_coverage(enriched_df: pd.DataFrame) -> None:
    """At least 150 rows should have Research Data (not sector estimates)."""
    research = (enriched_df["Data Source"] == "Research Data").sum()
    assert research >= 150, f"Only {research} Research Data rows (expected ≥150)"
