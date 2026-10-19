"""Market data tools: stock quotes, financial statements, and SEC EDGAR filings.

Uses yfinance for prices/statements and the SEC EDGAR REST API for filing
metadata. Results are returned as JSON strings (the agent's tool-call protocol).
"""

from __future__ import annotations

import json
import os
from typing import Any

import pandas as pd
import requests
import yfinance as yf
from agents import function_tool


def _df_to_records(df: pd.DataFrame | None, max_cols: int = 6) -> list[dict[str, Any]]:
    if df is None or df.empty:
        return []
    df = df.copy()
    if len(df.columns) > max_cols:
        df = df.iloc[:, :max_cols]
    df.columns = [str(c)[:10] for c in df.columns]
    df.index = [str(i) for i in df.index]
    df = df.reset_index().rename(columns={"index": "line_item"})
    df = df.where(pd.notna(df), None)
    return df.to_dict(orient="records")


@function_tool
def get_stock_quote(ticker: str) -> str:
    """Get the current quote, market cap, and basic info for a stock ticker.

    Returns JSON with price, market cap, P/E, dividend yield, 52-week range,
    sector/industry, and beta. Useful as a quick-look snapshot.
    """
    try:
        t = yf.Ticker(ticker.upper())
        info = t.info or {}
        fast = getattr(t, "fast_info", None)
        price = None
        if fast is not None:
            price = getattr(fast, "last_price", None) or fast.get("last_price") if hasattr(fast, "get") else None
        if price is None:
            price = info.get("currentPrice") or info.get("regularMarketPrice")

        out = {
            "ticker": ticker.upper(),
            "name": info.get("longName") or info.get("shortName"),
            "exchange": info.get("exchange"),
            "currency": info.get("currency"),
            "price": price,
            "previous_close": info.get("previousClose"),
            "day_range": [info.get("dayLow"), info.get("dayHigh")],
            "52w_range": [info.get("fiftyTwoWeekLow"), info.get("fiftyTwoWeekHigh")],
            "market_cap": info.get("marketCap"),
            "enterprise_value": info.get("enterpriseValue"),
            "trailing_pe": info.get("trailingPE"),
            "forward_pe": info.get("forwardPE"),
            "price_to_book": info.get("priceToBook"),
            "ev_to_ebitda": info.get("enterpriseToEbitda"),
            "dividend_yield": info.get("dividendYield"),
            "beta": info.get("beta"),
            "sector": info.get("sector"),
            "industry": info.get("industry"),
            "shares_outstanding": info.get("sharesOutstanding"),
            "description": (info.get("longBusinessSummary") or "")[:600],
        }
        return json.dumps(out, default=str, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e), "ticker": ticker})


@function_tool
def get_price_history(ticker: str, period: str = "1y", interval: str = "1d") -> str:
    """Get historical price data for a ticker.

    Args:
        ticker: Stock symbol.
        period: One of 1d, 5d, 1mo, 3mo, 6mo, 1y, 2y, 5y, 10y, ytd, max.
        interval: One of 1d, 1wk, 1mo (daily/weekly/monthly bars).

    Returns JSON with summary stats and a sample of the most recent rows.
    """
    try:
        t = yf.Ticker(ticker.upper())
        df = t.history(period=period, interval=interval, auto_adjust=True)
        if df.empty:
            return json.dumps({"error": "no data", "ticker": ticker})
        df = df.tail(60)
        rec = df.reset_index().to_dict(orient="records")
        for row in rec:
            for k, v in list(row.items()):
                row[k] = str(v) if hasattr(v, "isoformat") else v
        full = t.history(period=period, interval=interval, auto_adjust=True)
        return json.dumps(
            {
                "ticker": ticker.upper(),
                "period": period,
                "interval": interval,
                "stats": {
                    "start": str(full.index[0].date()),
                    "end": str(full.index[-1].date()),
                    "start_close": float(full["Close"].iloc[0]),
                    "end_close": float(full["Close"].iloc[-1]),
                    "total_return_pct": float(full["Close"].iloc[-1] / full["Close"].iloc[0] - 1) * 100,
                    "high": float(full["High"].max()),
                    "low": float(full["Low"].min()),
                    "avg_volume": float(full["Volume"].mean()) if "Volume" in full else None,
                },
                "recent_bars": rec[-20:],
            },
            default=str,
            indent=2,
        )
    except Exception as e:
        return json.dumps({"error": str(e), "ticker": ticker})


@function_tool
def get_financial_statements(ticker: str, statement: str = "all", period: str = "annual") -> str:
    """Pull income statement, balance sheet, and/or cash flow statement.

    Args:
        ticker: Stock symbol.
        statement: One of "income", "balance", "cashflow", or "all".
        period: "annual" or "quarterly".

    Returns JSON with line items by fiscal period. Values are in the company's
    reporting currency, raw units (not millions).
    """
    try:
        t = yf.Ticker(ticker.upper())
        out: dict[str, Any] = {"ticker": ticker.upper(), "period": period}

        if statement in ("income", "all"):
            df = t.income_stmt if period == "annual" else t.quarterly_income_stmt
            out["income_statement"] = _df_to_records(df)
        if statement in ("balance", "all"):
            df = t.balance_sheet if period == "annual" else t.quarterly_balance_sheet
            out["balance_sheet"] = _df_to_records(df)
        if statement in ("cashflow", "all"):
            df = t.cashflow if period == "annual" else t.quarterly_cashflow
            out["cashflow_statement"] = _df_to_records(df)

        return json.dumps(out, default=str, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e), "ticker": ticker})


@function_tool
def get_analyst_estimates(ticker: str) -> str:
    """Get analyst price targets, recommendations, and earnings estimates."""
    try:
        t = yf.Ticker(ticker.upper())
        info = t.info or {}
        out = {
            "ticker": ticker.upper(),
            "target_mean": info.get("targetMeanPrice"),
            "target_median": info.get("targetMedianPrice"),
            "target_high": info.get("targetHighPrice"),
            "target_low": info.get("targetLowPrice"),
            "num_analysts": info.get("numberOfAnalystOpinions"),
            "recommendation": info.get("recommendationKey"),
            "recommendation_mean": info.get("recommendationMean"),
            "earnings_growth": info.get("earningsGrowth"),
            "revenue_growth": info.get("revenueGrowth"),
        }
        try:
            rec = t.recommendations
            if rec is not None and not rec.empty:
                out["recent_recommendations"] = _df_to_records(rec.tail(10))
        except Exception:
            pass
        return json.dumps(out, default=str, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e), "ticker": ticker})


# ---------------------------------------------------------------------------
# SEC EDGAR
# ---------------------------------------------------------------------------

SEC_BASE = "https://data.sec.gov"
SEC_WWW = "https://www.sec.gov"


def _sec_headers() -> dict[str, str]:
    ua = os.environ.get("SEC_USER_AGENT", "Finance Agent contact@example.com")
    return {"User-Agent": ua, "Accept": "application/json"}


def _ticker_to_cik(ticker: str) -> str | None:
    """Resolve a ticker to a zero-padded 10-digit CIK using SEC's mapping file."""
    url = f"{SEC_WWW}/files/company_tickers.json"
    r = requests.get(url, headers=_sec_headers(), timeout=15)
    r.raise_for_status()
    data = r.json()
    upper = ticker.upper()
    for _, row in data.items():
        if row.get("ticker", "").upper() == upper:
            return str(row["cik_str"]).zfill(10)
    return None


@function_tool
def get_sec_filings(ticker: str, form_type: str = "10-K", limit: int = 5) -> str:
    """List recent SEC EDGAR filings for a company.

    Args:
        ticker: Stock symbol.
        form_type: Filing type filter (e.g. "10-K", "10-Q", "8-K", "DEF 14A").
            Use "any" to list all forms.
        limit: Max filings to return (most recent first).

    Returns JSON with filing date, accession number, primary document URL,
    and a direct link to the filing index on EDGAR.
    """
    try:
        cik = _ticker_to_cik(ticker)
        if not cik:
            return json.dumps({"error": f"CIK not found for {ticker}"})

        url = f"{SEC_BASE}/submissions/CIK{cik}.json"
        r = requests.get(url, headers=_sec_headers(), timeout=15)
        r.raise_for_status()
        data = r.json()
        recent = data.get("filings", {}).get("recent", {})

        forms = recent.get("form", [])
        dates = recent.get("filingDate", [])
        accessions = recent.get("accessionNumber", [])
        primaries = recent.get("primaryDocument", [])

        results = []
        for form, date, acc, prim in zip(forms, dates, accessions, primaries):
            if form_type.lower() != "any" and form.upper() != form_type.upper():
                continue
            acc_nodash = acc.replace("-", "")
            results.append(
                {
                    "form": form,
                    "filing_date": date,
                    "accession_number": acc,
                    "primary_document": prim,
                    "filing_url": f"{SEC_WWW}/Archives/edgar/data/{int(cik)}/{acc_nodash}/{prim}",
                    "index_url": f"{SEC_WWW}/cgi-bin/browse-edgar?action=getcompany&CIK={cik}&type={form}&dateb=&owner=include&count=40",
                }
            )
            if len(results) >= limit:
                break

        return json.dumps(
            {
                "ticker": ticker.upper(),
                "cik": cik,
                "company": data.get("name"),
                "filings": results,
            },
            indent=2,
        )
    except Exception as e:
        return json.dumps({"error": str(e), "ticker": ticker})


@function_tool
def get_sec_company_facts(ticker: str, concept: str | None = None) -> str:
    """Pull structured XBRL company facts from SEC EDGAR.

    These are auditable, standardized fundamentals (e.g. "Revenues",
    "NetIncomeLoss", "Assets", "StockholdersEquity") going back many years.

    Args:
        ticker: Stock symbol.
        concept: Optional specific US-GAAP concept name. If omitted, returns a
            curated set of key concepts to keep output manageable.

    Returns JSON with annual values per concept.
    """
    try:
        cik = _ticker_to_cik(ticker)
        if not cik:
            return json.dumps({"error": f"CIK not found for {ticker}"})

        url = f"{SEC_BASE}/api/xbrl/companyfacts/CIK{cik}.json"
        r = requests.get(url, headers=_sec_headers(), timeout=20)
        r.raise_for_status()
        data = r.json()
        gaap = data.get("facts", {}).get("us-gaap", {})

        default_concepts = [
            "Revenues",
            "RevenueFromContractWithCustomerExcludingAssessedTax",
            "GrossProfit",
            "OperatingIncomeLoss",
            "NetIncomeLoss",
            "Assets",
            "Liabilities",
            "StockholdersEquity",
            "CashAndCashEquivalentsAtCarryingValue",
            "LongTermDebt",
            "NetCashProvidedByUsedInOperatingActivities",
        ]
        concepts = [concept] if concept else default_concepts

        out: dict[str, Any] = {"ticker": ticker.upper(), "cik": cik, "company": data.get("entityName"), "concepts": {}}
        for c in concepts:
            if c not in gaap:
                continue
            units = gaap[c].get("units", {})
            unit_key = next(iter(units), None)
            if not unit_key:
                continue
            rows = units[unit_key]
            annual = [r for r in rows if r.get("fp") == "FY" and r.get("form") == "10-K"]
            annual = sorted(annual, key=lambda r: r.get("end", ""), reverse=True)[:6]
            out["concepts"][c] = {
                "unit": unit_key,
                "label": gaap[c].get("label"),
                "annual": [{"end": r.get("end"), "value": r.get("val")} for r in annual],
            }
        return json.dumps(out, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e), "ticker": ticker})
