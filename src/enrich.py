"""
src/enrich.py
─────────────
Main enrichment pipeline for the ASX coverage toolkit.

Merges three data sources by priority:
    1. company_data.py  — hardcoded research-verified profiles (~178 companies)
    2. Live data        — yfinance (when fetch_live is implemented)
    3. Sector defaults  — fallback estimates for unresearched entities

Produces a 44-column DataFrame that is passed to format_excel.py.
"""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

from src.company_data import COMPANY_DATA
from src.index_sets import get_index_flags

# ── Sector → GICS tier-1 map ─────────────────────────────────────────────────
SECTOR_MAP: dict[str, str] = {
    "Banks": "Financials",
    "Financial Services": "Financials",
    "Insurance": "Financials",
    "Materials": "Materials",
    "Energy": "Energy",
    "Software & Services": "Information Technology",
    "Technology Hardware & Equipment": "Information Technology",
    "Semiconductors & Semiconductor Equipment": "Information Technology",
    "Pharmaceuticals, Biotechnology & Life Sciences": "Health Care",
    "Health Care Equipment & Services": "Health Care",
    "Capital Goods": "Industrials",
    "Commercial & Professional Services": "Industrials",
    "Transportation": "Industrials",
    "Consumer Discretionary Distribution & Retail": "Consumer Discretionary",
    "Automobiles & Components": "Consumer Discretionary",
    "Consumer Durables & Apparel": "Consumer Discretionary",
    "Consumer Services": "Consumer Discretionary",
    "Food, Beverage & Tobacco": "Consumer Staples",
    "Consumer Staples Distribution & Retail": "Consumer Staples",
    "Household & Personal Products": "Consumer Staples",
    "Telecommunication Services": "Communication Services",
    "Media & Entertainment": "Communication Services",
    "Equity Real Estate Investment Trusts (REITs)": "Real Estate",
    "Real Estate Management & Development": "Real Estate",
    "Utilities": "Utilities",
    "Not Applic": "Unclassified",
    "Class Pend": "Unclassified",
}

# Sector-level fallback defaults (used when no individual data exists)
SECTOR_DEFAULTS: dict[str, dict[str, Any]] = {
    "Materials":             dict(dy=0.0,  frank=0,  pe=0.0,  pb=1.0, beta=1.1, roe=-5.0, de=0.1,  rev=0.010, np_=-0.002, desc="Exploration or early-stage mining/materials company.", cat="Micro Cap"),
    "Energy":                dict(dy=1.0,  frank=50, pe=8.0,  pb=1.0, beta=1.1, roe=5.0,  de=0.2,  rev=0.050, np_=0.002,  desc="Oil, gas or energy exploration/production company.",  cat="Micro Cap"),
    "Health Care":           dict(dy=0.0,  frank=0,  pe=0.0,  pb=3.0, beta=0.9, roe=-15.0,de=0.0,  rev=0.020, np_=-0.005, desc="Healthcare or biotechnology company — typically pre-revenue.", cat="Micro Cap"),
    "Information Technology":dict(dy=0.0,  frank=0,  pe=0.0,  pb=3.0, beta=1.2, roe=-10.0,de=0.0,  rev=0.010, np_=-0.002, desc="Technology software or services company.", cat="Micro Cap"),
    "Financials":            dict(dy=3.0,  frank=50, pe=15.0, pb=1.5, beta=0.9, roe=8.0,  de=0.2,  rev=0.050, np_=0.005,  desc="Financial services company.", cat="Micro Cap"),
    "Industrials":           dict(dy=2.0,  frank=80, pe=14.0, pb=1.5, beta=0.9, roe=8.0,  de=0.3,  rev=0.050, np_=0.002,  desc="Industrial services or capital goods company.", cat="Micro Cap"),
    "Consumer Discretionary":dict(dy=2.0,  frank=80, pe=16.0, pb=2.0, beta=0.9, roe=8.0,  de=0.2,  rev=0.050, np_=0.003,  desc="Consumer discretionary products or services company.", cat="Micro Cap"),
    "Consumer Staples":      dict(dy=3.0,  frank=80, pe=14.0, pb=2.0, beta=0.7, roe=10.0, de=0.3,  rev=0.100, np_=0.005,  desc="Consumer staples company.", cat="Micro Cap"),
    "Communication Services":dict(dy=2.0,  frank=80, pe=15.0, pb=2.0, beta=0.9, roe=8.0,  de=0.3,  rev=0.050, np_=0.003,  desc="Communications or media company.", cat="Micro Cap"),
    "Real Estate":           dict(dy=5.0,  frank=0,  pe=18.0, pb=1.0, beta=0.8, roe=5.0,  de=0.4,  rev=0.020, np_=0.005,  desc="Real estate investment trust or property developer.", cat="Micro Cap"),
    "Utilities":             dict(dy=4.0,  frank=50, pe=18.0, pb=1.5, beta=0.6, roe=6.0,  de=0.6,  rev=0.050, np_=0.003,  desc="Utility services company.", cat="Micro Cap"),
    "Unclassified":          dict(dy=0.0,  frank=0,  pe=0.0,  pb=1.0, beta=1.0, roe=0.0,  de=0.0,  rev=0.0,   np_=0.0,    desc="Special vehicle, trust, ETF, or pending classification.", cat="Unknown"),
}


# ── Helpers ───────────────────────────────────────────────────────────────────

def cap_category(mc_b: float | None) -> str:
    """Assign market cap category from a value in A$B."""
    if mc_b is None:
        return "Unknown"
    if mc_b >= 50:   return "Mega Cap"
    if mc_b >= 10:   return "Large Cap"
    if mc_b >= 2:    return "Mid Cap"
    if mc_b >= 0.3:  return "Small Cap"
    if mc_b >= 0.05: return "Micro Cap"
    return "Nano Cap"


def detect_entity_type(name: str, gics_group: str) -> str:
    """Heuristic entity type from company name and GICS group."""
    n = name.upper()
    if gics_group == "Equity Real Estate Investment Trusts (REITs)":
        return "REIT (Stapled)" if "STAPLED" in n else "REIT"
    if any(x in n for x in ("ETF", "EXCHANGE TRADED FUND")):
        return "ETF"
    if any(x in n for x in ("LISTED INVESTMENT COMPANY", " LIC ")):
        return "LIC"
    if "LISTED INVESTMENT TRUST" in n:
        return "LIT"
    if any(x in n for x in ("ABS TRUST",)) or ("TRUST" in n and re.search(r"\bABS\b", n)):
        return "ABS Trust"
    if any(x in n for x in ("INVESTMENT TRUST", "INVESTMENT FUND", "CAPITAL FUND",
                              "INCOME TRUST", "MORTGAGE REIT", "INCOME FUND", "CREDIT FUND")):
        return "Investment Fund/Trust"
    if "WARRANT" in n:
        return "Warrant"
    if gics_group == "Real Estate Management & Development":
        return "Property Group"
    if gics_group == "Class Pend":
        return "Class Pending"
    if gics_group == "Not Applic":
        return "Special Vehicle"
    return "Company"


def detect_domicile(name: str, curr: str) -> str:
    """Infer domicile from company name patterns."""
    n = name.upper()
    if "FOREIGN EXEMPT NYSE" in n or "FOREIGN EXEMPT NASDAQ" in n:
        return "USA"
    if "FOREIGN EXEMPT LSE" in n:
        return "UK"
    if "FOREIGN EXEMPT NZX" in n:
        return "New Zealand"
    if "FOREIGN EXEMPT XPAR" in n:
        return "France"
    if "FOREIGN EXEMPT" in n:
        return "Foreign"
    if curr == "NZD":
        return "New Zealand"
    return "Australia"


def _fmt_mc(v: float | None) -> str:
    return f"A${v:.3f}B" if v is not None else "N/A"

def _fmt_sh(v: float | None) -> str:
    return f"{v:.3f}B" if v is not None else "N/A"

def _fmt_px(v: float | None) -> str:
    return f"A${v:.4f}" if v is not None else "N/A"

def _fmt_rev(v: float | None) -> str:
    if v is None or v == 0:
        return "N/A"
    return f"A${v:.3f}B"

def _fmt_np(v: float | None) -> str:
    if v is None:
        return "N/A"
    return f"A${v:.3f}B"

def _fmt_pct(v: float | None) -> str:
    return f"{v:.1f}%" if v is not None else "N/A"

def _fmt_ratio(v: float | None, suffix: str = "x") -> str:
    return f"{v:.1f}{suffix}" if v is not None else "N/A"

def _fmt_beta(v: float | None) -> str:
    if v is None:
        return "N/A"
    # Format without trailing zeros: 0.85 not 0.850
    s = f"{v:.2f}"
    if s.endswith("0") and "." in s:
        s = s.rstrip("0").rstrip(".")
        if "." not in s:
            s += ".0"
    return s

def _clean_notes(parts: list[str]) -> str:
    return "; ".join(parts) if parts else "No flags"


# ── Live-data mapping (yfinance raw → record fields) ──────────────────────────

def _pick(cd: dict[str, Any], lr: dict[str, Any], key: str, *fallbacks: Any) -> Any:
    """First non-None value across research data, then live data, then fallbacks."""
    for source in (cd, lr):
        v = source.get(key)
        if v is not None:
            return v
    for fb in fallbacks:
        if fb is not None:
            return fb
    return None


def _first_sentence(text: str | None, limit: int = 180) -> str:
    """First sentence of a business summary, truncated for the one-line column."""
    if not text:
        return ""
    sentence = re.split(r"(?<=[.!?])\s", text.strip())[0]
    return f"{sentence[:limit].rstrip()}…" if len(sentence) > limit else sentence


def _strip_url(url: str | None) -> str:
    if not url:
        return ""
    return url.replace("https://", "").replace("http://", "").replace("www.", "").rstrip("/")


def _norm_yield(raw: dict[str, Any]) -> float | None:
    """Normalise a yfinance dividend yield to a percentage (e.g. 3.8)."""
    tay = raw.get("trailingAnnualDividendYield")
    if isinstance(tay, (int, float)) and tay > 0:
        return round(tay * 100, 2)
    dyv = raw.get("dividendYield")
    if isinstance(dyv, (int, float)) and dyv > 0:
        # Some yfinance versions report percent, others a fraction.
        return round(dyv if dyv > 1.5 else dyv * 100, 2)
    return None


def _live_record(raw: dict[str, Any]) -> dict[str, Any]:
    """
    Map a raw yfinance entry onto the same field schema as company_data records.
    Returns {} when the entry carries no usable price or market-cap signal, so a
    failed/empty fetch is never mistaken for genuine Live Data.
    """
    if not raw:
        return {}
    rec: dict[str, Any] = {}

    mc = raw.get("marketCap")
    if isinstance(mc, (int, float)) and mc > 0:
        rec["mc"] = round(mc / 1e9, 3)
    sh = raw.get("sharesOutstanding")
    if isinstance(sh, (int, float)) and sh > 0:
        rec["sh"] = round(sh / 1e9, 3)
    px = raw.get("currentPrice") or raw.get("regularMarketPrice") or raw.get("previousClose")
    if isinstance(px, (int, float)) and px > 0:
        rec["px"] = round(px, 4)

    if "mc" not in rec and "px" not in rec:
        return {}

    pe = raw.get("trailingPE")
    if isinstance(pe, (int, float)) and pe > 0:
        rec["pe"] = round(pe, 1)
    pb = raw.get("priceToBook")
    if isinstance(pb, (int, float)) and pb > 0:
        rec["pb"] = round(pb, 1)
    beta = raw.get("beta")
    if isinstance(beta, (int, float)):
        rec["beta"] = round(beta, 2)
    dy = _norm_yield(raw)
    if dy is not None:
        rec["dy"] = dy
    roe = raw.get("returnOnEquity")
    if isinstance(roe, (int, float)):
        rec["roe"] = round(roe * 100, 1)
    de = raw.get("debtToEquity")
    if isinstance(de, (int, float)) and de >= 0:
        rec["de"] = round(de / 100, 2)
    rev = raw.get("totalRevenue")
    if isinstance(rev, (int, float)) and rev > 0:
        rec["rev"] = round(rev / 1e9, 3)
    npc = raw.get("netIncomeToCommon")
    if isinstance(npc, (int, float)):
        rec["np"] = round(npc / 1e9, 3)

    desc = _first_sentence(raw.get("longBusinessSummary"))
    if desc:
        rec["desc"] = desc
    web = _strip_url(raw.get("website"))
    if web:
        rec["web"] = web
    curr = raw.get("financialCurrency") or raw.get("currency")
    if curr:
        rec["curr"] = curr
    country = raw.get("country")
    if country:
        rec["ops"] = country

    fetched = raw.get("fetched_at")
    if fetched:
        try:
            rec["date"] = datetime.fromisoformat(fetched).strftime("%b-%Y")
        except (ValueError, TypeError):
            pass

    return rec


# ── Main enrichment function ──────────────────────────────────────────────────

def enrich(
    source_csv: Path,
    live_data: dict[str, dict[str, Any]] | None = None,
) -> pd.DataFrame:
    """
    Build the full 44-column enriched DataFrame.

    Parameters
    ----------
    source_csv : path to ASX_Entities_Enriched.csv
    live_data  : optional dict from fetch_live.get_live_data()
                 keys are bare ASX tickers, values are yfinance data dicts

    Returns
    -------
    pd.DataFrame with 44 columns, 1 row per ASX entity
    """
    src = pd.read_csv(source_csv)
    live = live_data or {}
    rows: list[dict[str, Any]] = []

    for _, r in src.iterrows():
        ticker: str = str(r["ASX code"]).strip()
        name: str   = str(r["Company name"]).strip().rstrip(".")
        ig: str     = str(r["GICS industry group"]).strip()
        sub: str    = str(r["GICS Industry (Sub-Level)"]).strip()

        sector = SECTOR_MAP.get(ig, "Unclassified")
        cd     = COMPANY_DATA.get(ticker, {})
        lr     = _live_record(live.get(ticker, {}))
        sd     = SECTOR_DEFAULTS.get(sector, SECTOR_DEFAULTS["Unclassified"])

        # Source CSV values (only 30 rows in practice)
        s_mc = r["Market Cap (Billion AUD)"] if pd.notna(r["Market Cap (Billion AUD)"]) else None
        s_sh = r["Outstanding Shares (Billion)"] if pd.notna(r["Outstanding Shares (Billion)"]) else None
        s_px = r["Share Price (AUD)"] if pd.notna(r["Share Price (AUD)"]) else None

        # Price/valuation fields take the LIVE value first (current market),
        # then curated, then sector default — so marquee names show fresh figures.
        mc    = _pick(lr, cd, "mc", s_mc)
        sh    = _pick(lr, cd, "sh", s_sh)
        px    = _pick(lr, cd, "px", s_px)
        pe    = _pick(lr, cd, "pe", sd["pe"])
        pb    = _pick(lr, cd, "pb", sd["pb"])
        beta  = _pick(lr, cd, "beta", sd["beta"])
        dy    = _pick(lr, cd, "dy", sd["dy"])
        # Statement fundamentals prefer curated (hand-verified), with live as fallback.
        roe   = _pick(cd, lr, "roe", sd["roe"])
        de    = _pick(cd, lr, "de", sd["de"])
        rev   = _pick(cd, lr, "rev", sd["rev"])
        np_   = _pick(cd, lr, "np", sd["np_"])
        frank = cd.get("frank", sd["frank"])      # franking credits are not on yfinance

        desc  = cd.get("desc") or lr.get("desc") or sd["desc"]
        state = cd.get("state", "N/A")
        ops   = cd.get("ops") or lr.get("ops") or ("Australia" if sector != "Unclassified" else "N/A")
        dual  = cd.get("dual", False)
        dex   = cd.get("dex", "")
        ceo   = cd.get("ceo", "N/A")
        web   = cd.get("web") or lr.get("web") or "N/A"
        yr    = cd.get("yr", "N/A")
        fye   = cd.get("fye", "Jun")
        curr  = cd.get("curr") or lr.get("curr") or "AUD"
        cat   = cap_category(mc) if mc is not None else cd.get("cat", "Unknown")

        # Data source label: research-verified, live (yfinance), or sector default.
        if cd:
            data_src = "Research Data"
        elif lr:
            data_src = "Live Data"
        else:
            data_src = "Sector Estimate"

        # Market-cap "as at" date reflects the data's true vintage (live takes priority).
        if mc is None:
            mc_date = "N/A"
        elif "mc" in lr:
            mc_date = lr.get("date", "Live")
        else:
            mc_date = "Apr-2026"

        # Derived fields
        ig_clean = ig if ig not in ("Not Applic", "Class Pend") else "Unclassified"
        ind      = sub if (";" not in sub and sub != "Other/Not Applicable") else "Multiple / See Notes"
        sub_ind  = sub if sub != "Other/Not Applicable" else "Unclassified"
        etype    = detect_entity_type(name, ig)
        domicile = detect_domicile(name, curr)
        idx      = get_index_flags(ticker)

        # Notes
        notes: list[str] = []
        if ig == "Not Applic":
            notes.append("GICS not applicable — special vehicle/ETF/LIC/trust")
        if ig == "Class Pend":
            notes.append("GICS classification pending")
        if ";" in sub:
            notes.append("Multi-industry: conglomerate/diversified")
        if mc is None:
            notes.append("Market cap not available")
        if dual:
            notes.append(f"Dual-listed: {dex}")
        if domicile not in ("Australia", "N/A"):
            notes.append(f"Foreign domicile: {domicile}")
        if data_src == "Live Data":
            notes.append("Live market data via yfinance")
        if data_src == "Sector Estimate":
            notes.append("Financial data estimated from sector/size defaults")

        rows.append({
            "Company":                      name,
            "ASX Ticker":                   ticker,
            "Entity Type":                  etype,
            "Market Cap Category":          cat,
            "GICS Sector":                  sector,
            "GICS Industry Group":          ig_clean,
            "GICS Industry":                ind,
            "GICS Sub-Industry":            sub_ind,
            "Market Cap (AUD)":             _fmt_mc(mc),
            "Market Cap Data Date":         mc_date,
            "Share Price (AUD)":            _fmt_px(px) if px is not None else "N/A",
            "Shares Outstanding":           _fmt_sh(sh),
            "Revenue (AUD, Latest Annual)": _fmt_rev(rev),
            "Net Profit/Loss (AUD)":        _fmt_np(np_),
            "Dividend Yield":               _fmt_pct(dy),
            "Franking Level":               f"{frank}%" if frank is not None else "N/A",
            "P/E Ratio":                    _fmt_ratio(pe) if pe else "N/A",
            "Price/Book":                   _fmt_ratio(pb),
            "Beta (vs ASX 200)":            _fmt_beta(beta),
            "Return on Equity (ROE)":       _fmt_pct(roe),
            "Debt/Equity Ratio":            _fmt_ratio(de),
            "Reporting Currency":           curr,
            "Fiscal Year End":              fye,
            "CEO / MD":                     ceo,
            "Website":                      web,
            "Year Founded":                 str(yr),
            "Business Description":         desc,
            "HQ State / Territory":         state,
            "Primary Country of Operations":ops,
            "Domicile":                     domicile,
            "Primary Exchange":             "ASX",
            "Dual Listed":                  "Yes" if dual else "No",
            "Dual Listing Exchange":        dex if dex else "N/A",
            "Listing Status":               "Listed",
            **idx,
            "Data Source":                  data_src,
            "BHS Rating":                   "",
            "Progress":                     "",
            "Notes":                        _clean_notes(notes),
        })

    return pd.DataFrame(rows)
