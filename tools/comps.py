"""Peer comparables analysis.

Builds a side-by-side multiples table for a set of tickers and derives an
implied valuation range for a target from peer medians/quartiles.

Why this matters: a DCF is highly assumption-sensitive. Triangulating against
trading comps anchors the valuation in what the market is paying *today* for
similar businesses.
"""

from __future__ import annotations

import json
import statistics
from typing import Any

import yfinance as yf
from agents import function_tool


METRIC_KEYS = {
    "market_cap": "marketCap",
    "enterprise_value": "enterpriseValue",
    "revenue": "totalRevenue",
    "ebitda": "ebitda",
    "net_income": "netIncomeToCommon",
    "trailing_pe": "trailingPE",
    "forward_pe": "forwardPE",
    "price_to_book": "priceToBook",
    "price_to_sales": "priceToSalesTrailing12Months",
    "ev_to_ebitda": "enterpriseToEbitda",
    "ev_to_revenue": "enterpriseToRevenue",
    "gross_margin": "grossMargins",
    "operating_margin": "operatingMargins",
    "net_margin": "profitMargins",
    "return_on_equity": "returnOnEquity",
    "revenue_growth": "revenueGrowth",
    "earnings_growth": "earningsGrowth",
    "debt_to_equity": "debtToEquity",
}


def _fetch_company_metrics(ticker: str) -> dict[str, Any]:
    """Pull a comparable metrics dict from yfinance for one ticker."""
    try:
        t = yf.Ticker(ticker.upper())
        info = t.info or {}
        out: dict[str, Any] = {
            "ticker": ticker.upper(),
            "name": info.get("longName") or info.get("shortName"),
            "sector": info.get("sector"),
            "industry": info.get("industry"),
            "currency": info.get("currency"),
        }
        for friendly, yf_key in METRIC_KEYS.items():
            out[friendly] = info.get(yf_key)
        return out
    except Exception as e:
        return {"ticker": ticker.upper(), "error": str(e)}


def _summary_stats(values: list[float | None]) -> dict[str, float | None]:
    """Compute count/mean/median/quartiles ignoring None and non-finite values."""
    clean = [float(v) for v in values if isinstance(v, (int, float)) and v == v]
    if not clean:
        return {"count": 0, "mean": None, "median": None, "p25": None, "p75": None, "min": None, "max": None}
    clean_sorted = sorted(clean)
    n = len(clean_sorted)
    def pct(p: float) -> float:
        if n == 1:
            return clean_sorted[0]
        idx = (n - 1) * p
        lo, hi = int(idx), min(int(idx) + 1, n - 1)
        return clean_sorted[lo] + (clean_sorted[hi] - clean_sorted[lo]) * (idx - lo)
    return {
        "count": n,
        "mean": statistics.fmean(clean_sorted),
        "median": statistics.median(clean_sorted),
        "p25": pct(0.25),
        "p75": pct(0.75),
        "min": clean_sorted[0],
        "max": clean_sorted[-1],
    }


@function_tool
def get_peer_comps(tickers: list[str]) -> str:
    """Build a side-by-side multiples table for a list of comparable companies.

    Args:
        tickers: 2-10 ticker symbols to compare (e.g. ["NVDA", "AMD", "INTC"]).

    Returns JSON with:
      - rows: one row per ticker with valuation, profitability, growth, leverage
      - summary: count / mean / median / p25 / p75 / min / max per metric
    Use the medians for a quick peer benchmark and the IQR to gauge dispersion.
    """
    if not tickers or len(tickers) < 2:
        return json.dumps({"error": "provide at least 2 tickers"})
    if len(tickers) > 15:
        return json.dumps({"error": "max 15 tickers per call to keep output manageable"})

    rows = [_fetch_company_metrics(t) for t in tickers]

    summary: dict[str, dict[str, Any]] = {}
    for metric in METRIC_KEYS:
        values = [r.get(metric) for r in rows]
        summary[metric] = _summary_stats(values)

    return json.dumps(
        {
            "tickers": [t.upper() for t in tickers],
            "rows": rows,
            "summary": summary,
        },
        default=str,
        indent=2,
    )


@function_tool
def value_target_from_comps(
    target_ticker: str,
    peer_tickers: list[str],
    multiples: list[str] | None = None,
) -> str:
    """Estimate the target's implied equity value using peer-median multiples.

    For each selected multiple, applies the peer-median to the target's
    corresponding line item and backs out an implied share price.

    Args:
        target_ticker: The company you want to value.
        peer_tickers: 2-10 comparable companies.
        multiples: Which multiples to apply. Defaults to
            ["ev_to_ebitda", "ev_to_revenue", "trailing_pe", "price_to_sales", "price_to_book"].

    Returns JSON with per-multiple implied prices and a summary range.
    """
    multiples = multiples or [
        "ev_to_ebitda",
        "ev_to_revenue",
        "trailing_pe",
        "price_to_sales",
        "price_to_book",
    ]

    target = _fetch_company_metrics(target_ticker)
    if "error" in target:
        return json.dumps({"error": target["error"], "ticker": target_ticker})

    peers = [_fetch_company_metrics(t) for t in peer_tickers]

    try:
        t = yf.Ticker(target_ticker.upper())
        info = t.info or {}
        shares = info.get("sharesOutstanding") or 0
        current_price = info.get("currentPrice") or info.get("regularMarketPrice")
        net_debt = (info.get("totalDebt") or 0) - (info.get("totalCash") or 0)
        revenue = target.get("revenue")
        ebitda = target.get("ebitda")
        net_income = target.get("net_income")
        book_value_per_share = info.get("bookValue")
    except Exception as e:
        return json.dumps({"error": str(e)})

    results: list[dict[str, Any]] = []
    for mult in multiples:
        peer_vals = [p.get(mult) for p in peers if not p.get("error")]
        stats = _summary_stats(peer_vals)
        peer_median = stats.get("median")
        if peer_median is None or shares == 0:
            results.append({"multiple": mult, "peer_median": None, "implied_price": None, "note": "insufficient data"})
            continue

        implied_price = None
        note = None

        if mult == "ev_to_ebitda" and ebitda:
            implied_ev = peer_median * ebitda
            implied_equity = implied_ev - net_debt
            implied_price = implied_equity / shares
        elif mult == "ev_to_revenue" and revenue:
            implied_ev = peer_median * revenue
            implied_equity = implied_ev - net_debt
            implied_price = implied_equity / shares
        elif mult == "trailing_pe" and net_income:
            implied_equity = peer_median * net_income
            implied_price = implied_equity / shares
        elif mult == "price_to_sales" and revenue:
            implied_equity = peer_median * revenue
            implied_price = implied_equity / shares
        elif mult == "price_to_book" and book_value_per_share:
            implied_price = peer_median * book_value_per_share
        else:
            note = f"target missing required input for {mult}"

        results.append(
            {
                "multiple": mult,
                "peer_median": peer_median,
                "peer_p25": stats.get("p25"),
                "peer_p75": stats.get("p75"),
                "peer_count": stats.get("count"),
                "implied_price": implied_price,
                "vs_current_pct": ((implied_price / current_price - 1) * 100) if (implied_price and current_price) else None,
                "note": note,
            }
        )

    implied_prices = [r["implied_price"] for r in results if r["implied_price"] is not None]
    range_stats = _summary_stats(implied_prices)

    return json.dumps(
        {
            "target": target_ticker.upper(),
            "current_price": current_price,
            "peers": [t.upper() for t in peer_tickers],
            "implied_by_multiple": results,
            "implied_price_range": {
                "low": range_stats.get("min"),
                "median": range_stats.get("median"),
                "mean": range_stats.get("mean"),
                "high": range_stats.get("max"),
                "vs_current_median_pct": (
                    (range_stats.get("median") / current_price - 1) * 100
                    if (range_stats.get("median") and current_price) else None
                ),
            },
            "inputs_used": {
                "shares_outstanding": shares,
                "net_debt": net_debt,
                "revenue": revenue,
                "ebitda": ebitda,
                "net_income": net_income,
                "book_value_per_share": book_value_per_share,
            },
        },
        default=str,
        indent=2,
    )
