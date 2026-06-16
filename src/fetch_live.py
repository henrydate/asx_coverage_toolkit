"""
src/fetch_live.py
─────────────────
Batch fetches live market data from yfinance for ASX tickers and caches the
result to data/yfinance_cache.json with a per-ticker 24-hour TTL.

ASX tickers are quoted in AUD on Yahoo Finance under the ``.AX`` suffix
(e.g. ``CBA`` → ``CBA.AX``). Fetches are fault-tolerant: a failed ticker is
logged and skipped, never crashing the pipeline.
"""

from __future__ import annotations

import json
import logging
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)

# Fields pulled from each yfinance ``Ticker.info`` dict.
YFINANCE_FIELDS: list[str] = [
    "currentPrice",
    "regularMarketPrice",
    "previousClose",
    "marketCap",
    "sharesOutstanding",
    "trailingPE",
    "forwardPE",
    "priceToBook",
    "beta",
    "dividendYield",
    "trailingAnnualDividendYield",
    "dividendRate",
    "exDividendDate",
    "fiftyTwoWeekHigh",
    "fiftyTwoWeekLow",
    "totalRevenue",
    "netIncomeToCommon",
    "returnOnEquity",
    "debtToEquity",
    "totalDebt",
    "totalCash",
    "revenueGrowth",
    "targetMeanPrice",
    "recommendationMean",
    "recommendationKey",
    "numberOfAnalystOpinions",
    "averageVolume",
    "averageVolume10days",
    "currency",
    "financialCurrency",
    "shortName",
    "longName",
    "sector",
    "industry",
    "country",
    "fullTimeEmployees",
    "website",
    "longBusinessSummary",
]

CACHE_TTL_HOURS = 24


# ── Cache I/O ──────────────────────────────────────────────────────────────────

def _is_cache_fresh(cache_path: Path) -> bool:
    """Return True if the cache file exists and is younger than the TTL."""
    if not cache_path.exists():
        return False
    mtime = datetime.fromtimestamp(cache_path.stat().st_mtime, tz=timezone.utc)
    age_hours = (datetime.now(tz=timezone.utc) - mtime).total_seconds() / 3600
    return age_hours < CACHE_TTL_HOURS


def load_cache(cache_path: Path) -> dict[str, dict[str, Any]]:
    """Load the yfinance cache from disk (empty dict if missing/corrupt)."""
    if not cache_path.exists():
        return {}
    try:
        with cache_path.open(encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError) as e:
        logger.warning("Cache read failed (%s) — starting fresh", e)
        return {}


def save_cache(cache: dict[str, dict[str, Any]], cache_path: Path) -> None:
    """Persist the yfinance cache to disk."""
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    with cache_path.open("w", encoding="utf-8") as f:
        json.dump(cache, f, indent=2, default=str)


def _entry_age_hours(entry: dict[str, Any]) -> float:
    """Age (hours) of a cached entry from its ``fetched_at`` stamp; inf if absent."""
    stamp = entry.get("fetched_at")
    if not stamp:
        return float("inf")
    try:
        fetched = datetime.fromisoformat(stamp)
    except (ValueError, TypeError):
        return float("inf")
    if fetched.tzinfo is None:
        fetched = fetched.replace(tzinfo=timezone.utc)
    return (datetime.now(tz=timezone.utc) - fetched).total_seconds() / 3600


# ── yfinance access ────────────────────────────────────────────────────────────

def to_yf_symbol(ticker: str) -> str:
    """Map a bare ASX code to its Yahoo Finance symbol (``CBA`` → ``CBA.AX``)."""
    t = ticker.strip().upper()
    return t if t.endswith(".AX") else f"{t}.AX"


def _extract(info: dict[str, Any]) -> dict[str, Any]:
    """Keep only the fields we care about, dropping None/empty values."""
    out: dict[str, Any] = {}
    for key in YFINANCE_FIELDS:
        val = info.get(key)
        if val is not None and val != "":
            out[key] = val
    return out


def _has_market_data(entry: dict[str, Any]) -> bool:
    """True if an entry carries a usable price or market-cap signal."""
    return any(
        entry.get(k) for k in ("marketCap", "currentPrice", "regularMarketPrice")
    )


def _fetch_one(yf_symbol: str, retries: int = 2, retry_delay: float = 1.5) -> dict[str, Any]:
    """
    Fetch a single ticker's info dict from yfinance with light retry.

    Returns the extracted field dict (possibly empty). Never raises.
    """
    import yfinance as yf

    for attempt in range(1, retries + 1):
        try:
            info = yf.Ticker(yf_symbol).info or {}
            data = _extract(info)
            if _has_market_data(data):
                return data
            # Empty/placeholder response — retry once more, then give up.
            if attempt < retries:
                time.sleep(retry_delay)
        except Exception as e:  # noqa: BLE001 — yfinance raises many error types
            logger.debug("Fetch failed for %s (attempt %d): %s", yf_symbol, attempt, e)
            if attempt < retries:
                time.sleep(retry_delay)
    return {}


def fetch_batch(
    tickers: list[str],
    cache_path: Path,
    force_refresh: bool = False,
    batch_size: int = 50,
    delay_seconds: float = 0.4,
) -> dict[str, dict[str, Any]]:
    """
    Fetch yfinance data for a list of bare ASX tickers.

    Parameters
    ----------
    tickers        : bare ASX codes, e.g. ['CBA', 'BHP'].
    cache_path     : path to the JSON cache file.
    force_refresh  : re-fetch every ticker, ignoring cached freshness.
    batch_size     : tickers per progress/sleep group.
    delay_seconds  : pause between groups to stay under rate limits.

    Returns
    -------
    Mapping of each bare ticker (that has usable data) to its yfinance field
    dict, augmented with ``fetched_at`` (UTC ISO) and ``yf_symbol``. Tickers
    that error or return nothing are omitted.
    """
    # yfinance logs an ERROR line per missing/delisted ticker; we summarise instead.
    logging.getLogger("yfinance").setLevel(logging.CRITICAL)

    cache = load_cache(cache_path)

    # Decide what actually needs a network call.
    stale: list[str] = []
    for raw in tickers:
        t = str(raw).strip().upper()
        cached = cache.get(t)
        fresh = (
            cached is not None
            and _has_market_data(cached)
            and _entry_age_hours(cached) < CACHE_TTL_HOURS
        )
        if force_refresh or not fresh:
            stale.append(t)

    if not stale:
        logger.info("All %d tickers fresh in cache — no fetch needed", len(tickers))
    else:
        logger.info(
            "Fetching %d/%d tickers from yfinance (%d cached fresh)",
            len(stale), len(tickers), len(tickers) - len(stale),
        )
        _fetch_into_cache(stale, cache, cache_path, batch_size, delay_seconds)
        save_cache(cache, cache_path)

    # Return only requested tickers that ended up with usable data.
    result: dict[str, dict[str, Any]] = {}
    for raw in tickers:
        t = str(raw).strip().upper()
        entry = cache.get(t)
        if entry and _has_market_data(entry):
            result[t] = entry
    return result


def _fetch_into_cache(
    stale: list[str],
    cache: dict[str, dict[str, Any]],
    cache_path: Path,
    batch_size: int,
    delay_seconds: float,
) -> None:
    """
    Fetch each stale ticker and write successes into ``cache`` in place.
    The cache is persisted every ``batch_size`` tickers so a long or
    interrupted run keeps its progress (re-running resumes where it left off).
    """
    try:
        from tqdm import tqdm
        progress = tqdm(stale, desc="yfinance", unit="ticker")
    except ImportError:
        progress = stale

    ok = 0
    for i, ticker in enumerate(progress, 1):
        yf_symbol = to_yf_symbol(ticker)
        data = _fetch_one(yf_symbol)
        if data:
            data["fetched_at"] = datetime.now(tz=timezone.utc).isoformat(timespec="seconds")
            data["yf_symbol"] = yf_symbol
            cache[ticker] = data
            ok += 1
        if i % batch_size == 0:
            logger.info("  …%d/%d fetched (%d with data)", i, len(stale), ok)
            save_cache(cache, cache_path)
            time.sleep(delay_seconds)
    logger.info("Fetch complete: %d/%d tickers returned data", ok, len(stale))


def get_live_data(
    tickers: list[str],
    cache_path: Path,
    force_refresh: bool = False,
) -> dict[str, dict[str, Any]]:
    """Public entry point: returns fresh-or-fetched live data for ``tickers``."""
    return fetch_batch(tickers, cache_path, force_refresh=force_refresh)


# ── Formatting helpers (raw yfinance value → display string) ───────────────────

def format_market_cap(value: float | None) -> str:
    """Raw market cap in dollars → ``A$X.XXXB`` string."""
    if value is None:
        return "N/A"
    return f"A${value / 1e9:.3f}B"


def format_price(value: float | None, currency: str = "AUD") -> str:
    """Share price → currency-prefixed string."""
    if value is None:
        return "N/A"
    prefix = "A$" if currency in ("AUD", "") else f"{currency}$"
    return f"{prefix}{value:.4f}"


def format_shares(value: float | None) -> str:
    """Raw share count → ``X.XXXB`` string."""
    if value is None:
        return "N/A"
    return f"{value / 1e9:.3f}B"


def format_percent(value: float | None) -> str:
    """0–1 ratio → percentage string."""
    if value is None:
        return "N/A"
    return f"{value * 100:.1f}%"


def format_ratio(value: float | None, suffix: str = "x") -> str:
    """Ratio → suffixed string."""
    if value is None:
        return "N/A"
    return f"{value:.1f}{suffix}"
