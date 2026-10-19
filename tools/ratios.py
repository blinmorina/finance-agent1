"""Financial ratio calculations.

Two tools:
  - compute_ratios_from_ticker: pulls statements via yfinance and computes ratios.
  - compute_ratios_from_inputs: lets the agent pass raw numbers (e.g. from a
    user-shared statement) and get back a clean ratio set.
"""

from __future__ import annotations

import json
from typing import Any

import pandas as pd
import yfinance as yf
from agents import function_tool


def _safe_div(num: float | None, den: float | None) -> float | None:
    try:
        if num is None or den is None or den == 0:
            return None
        return float(num) / float(den)
    except Exception:
        return None


def _get(df: pd.DataFrame | None, candidates: list[str], col_idx: int = 0) -> float | None:
    """Pick the first matching row label from a yfinance statement DataFrame."""
    if df is None or df.empty:
        return None
    for label in candidates:
        if label in df.index:
            try:
                val = df.iloc[df.index.get_loc(label), col_idx]
                if pd.notna(val):
                    return float(val)
            except Exception:
                continue
    return None


@function_tool
def compute_ratios_from_ticker(ticker: str) -> str:
    """Compute a full ratio set for a public company by ticker.

    Pulls the latest annual statements via yfinance and produces:
      - Liquidity (current, quick, cash)
      - Profitability (gross/operating/net margin, ROA, ROE, ROIC)
      - Leverage (D/E, D/A, interest coverage, debt/EBITDA)
      - Efficiency (asset turnover, inventory turnover, receivables turnover)
      - Valuation (P/E, P/B, EV/EBITDA, EV/Sales, FCF yield)

    Returns JSON with each ratio, the raw inputs used, and short interpretive notes.
    """
    try:
        t = yf.Ticker(ticker.upper())
        info = t.info or {}
        inc = t.income_stmt
        bs = t.balance_sheet
        cf = t.cashflow

        revenue = _get(inc, ["Total Revenue", "Revenue", "Operating Revenue"])
        gross = _get(inc, ["Gross Profit"])
        op_inc = _get(inc, ["Operating Income", "Operating Income Loss"])
        net_inc = _get(inc, ["Net Income", "Net Income Common Stockholders", "Net Income From Continuing Operation Net Minority Interest"])
        interest_exp = _get(inc, ["Interest Expense", "Interest Expense Non Operating"])
        ebit = _get(inc, ["EBIT"]) or op_inc
        ebitda = _get(inc, ["EBITDA", "Normalized EBITDA"])
        tax = _get(inc, ["Tax Provision", "Income Tax Expense"])
        pretax = _get(inc, ["Pretax Income", "Income Before Tax"])

        cur_assets = _get(bs, ["Current Assets", "Total Current Assets"])
        cur_liab = _get(bs, ["Current Liabilities", "Total Current Liabilities"])
        cash = _get(bs, ["Cash And Cash Equivalents", "Cash Cash Equivalents And Short Term Investments"])
        inventory = _get(bs, ["Inventory"])
        receivables = _get(bs, ["Accounts Receivable", "Receivables"])
        total_assets = _get(bs, ["Total Assets"])
        total_liab = _get(bs, ["Total Liabilities Net Minority Interest", "Total Liab"])
        equity = _get(bs, ["Stockholders Equity", "Common Stock Equity", "Total Equity Gross Minority Interest"])
        total_debt = _get(bs, ["Total Debt"])
        long_debt = _get(bs, ["Long Term Debt"])
        short_debt = _get(bs, ["Current Debt", "Short Long Term Debt"])
        if total_debt is None:
            total_debt = sum(x for x in [long_debt, short_debt] if x) or None

        cfo = _get(cf, ["Operating Cash Flow", "Cash Flow From Continuing Operating Activities"])
        capex = _get(cf, ["Capital Expenditure"])
        fcf = None
        if cfo is not None and capex is not None:
            fcf = cfo + capex  # capex is negative in yfinance

        tax_rate = _safe_div(tax, pretax) or 0.21
        nopat = ebit * (1 - tax_rate) if ebit is not None else None
        invested_capital = None
        if equity is not None and total_debt is not None:
            invested_capital = equity + total_debt

        market_cap = info.get("marketCap")
        enterprise_value = info.get("enterpriseValue")
        book_value = equity
        shares = info.get("sharesOutstanding")

        ratios = {
            "liquidity": {
                "current_ratio": _safe_div(cur_assets, cur_liab),
                "quick_ratio": _safe_div(((cur_assets or 0) - (inventory or 0)), cur_liab) if cur_assets else None,
                "cash_ratio": _safe_div(cash, cur_liab),
            },
            "profitability": {
                "gross_margin": _safe_div(gross, revenue),
                "operating_margin": _safe_div(op_inc, revenue),
                "net_margin": _safe_div(net_inc, revenue),
                "return_on_assets": _safe_div(net_inc, total_assets),
                "return_on_equity": _safe_div(net_inc, equity),
                "return_on_invested_capital": _safe_div(nopat, invested_capital),
                "ebitda_margin": _safe_div(ebitda, revenue),
            },
            "leverage": {
                "debt_to_equity": _safe_div(total_debt, equity),
                "debt_to_assets": _safe_div(total_debt, total_assets),
                "equity_multiplier": _safe_div(total_assets, equity),
                "interest_coverage": _safe_div(ebit, abs(interest_exp) if interest_exp else None),
                "debt_to_ebitda": _safe_div(total_debt, ebitda),
            },
            "efficiency": {
                "asset_turnover": _safe_div(revenue, total_assets),
                "inventory_turnover": _safe_div(revenue, inventory),
                "receivables_turnover": _safe_div(revenue, receivables),
                "days_sales_outstanding": _safe_div(receivables, revenue) * 365 if receivables and revenue else None,
            },
            "valuation": {
                "price_to_earnings": info.get("trailingPE"),
                "forward_pe": info.get("forwardPE"),
                "price_to_book": info.get("priceToBook"),
                "price_to_sales": info.get("priceToSalesTrailing12Months"),
                "ev_to_ebitda": info.get("enterpriseToEbitda"),
                "ev_to_revenue": info.get("enterpriseToRevenue"),
                "fcf_yield": _safe_div(fcf, market_cap),
                "dividend_yield": info.get("dividendYield"),
            },
            "growth": {
                "revenue_growth_yoy": info.get("revenueGrowth"),
                "earnings_growth_yoy": info.get("earningsGrowth"),
            },
        }

        inputs_used = {
            "revenue": revenue,
            "gross_profit": gross,
            "operating_income": op_inc,
            "ebit": ebit,
            "ebitda": ebitda,
            "net_income": net_inc,
            "interest_expense": interest_exp,
            "tax_rate_used": tax_rate,
            "current_assets": cur_assets,
            "current_liabilities": cur_liab,
            "cash": cash,
            "inventory": inventory,
            "receivables": receivables,
            "total_assets": total_assets,
            "total_liabilities": total_liab,
            "equity": equity,
            "total_debt": total_debt,
            "operating_cash_flow": cfo,
            "capex": capex,
            "free_cash_flow": fcf,
            "market_cap": market_cap,
            "enterprise_value": enterprise_value,
            "shares_outstanding": shares,
        }

        return json.dumps(
            {
                "ticker": ticker.upper(),
                "company": info.get("longName"),
                "fiscal_period": "latest annual",
                "ratios": ratios,
                "inputs_used": inputs_used,
            },
            default=str,
            indent=2,
        )
    except Exception as e:
        return json.dumps({"error": str(e), "ticker": ticker})


@function_tool
def compute_ratios_from_inputs(
    revenue: float | None = None,
    gross_profit: float | None = None,
    operating_income: float | None = None,
    net_income: float | None = None,
    interest_expense: float | None = None,
    ebitda: float | None = None,
    current_assets: float | None = None,
    current_liabilities: float | None = None,
    cash: float | None = None,
    inventory: float | None = None,
    receivables: float | None = None,
    total_assets: float | None = None,
    total_debt: float | None = None,
    equity: float | None = None,
    operating_cash_flow: float | None = None,
    capex: float | None = None,
    market_cap: float | None = None,
    enterprise_value: float | None = None,
    shares_outstanding: float | None = None,
    tax_rate: float = 0.21,
) -> str:
    """Compute ratios from raw financial inputs supplied by the user.

    All inputs are optional; ratios that lack required data return null. Pass
    capex as a negative number (consistent with cash-flow-statement convention)
    or positive — the function uses abs(capex).
    """
    fcf = None
    if operating_cash_flow is not None and capex is not None:
        fcf = operating_cash_flow - abs(capex)

    nopat = operating_income * (1 - tax_rate) if operating_income is not None else None
    invested_capital = None
    if equity is not None and total_debt is not None:
        invested_capital = equity + total_debt

    out = {
        "liquidity": {
            "current_ratio": _safe_div(current_assets, current_liabilities),
            "quick_ratio": _safe_div((current_assets or 0) - (inventory or 0), current_liabilities) if current_assets else None,
            "cash_ratio": _safe_div(cash, current_liabilities),
        },
        "profitability": {
            "gross_margin": _safe_div(gross_profit, revenue),
            "operating_margin": _safe_div(operating_income, revenue),
            "net_margin": _safe_div(net_income, revenue),
            "ebitda_margin": _safe_div(ebitda, revenue),
            "return_on_assets": _safe_div(net_income, total_assets),
            "return_on_equity": _safe_div(net_income, equity),
            "return_on_invested_capital": _safe_div(nopat, invested_capital),
        },
        "leverage": {
            "debt_to_equity": _safe_div(total_debt, equity),
            "debt_to_assets": _safe_div(total_debt, total_assets),
            "interest_coverage": _safe_div(operating_income, abs(interest_expense) if interest_expense else None),
            "debt_to_ebitda": _safe_div(total_debt, ebitda),
        },
        "efficiency": {
            "asset_turnover": _safe_div(revenue, total_assets),
            "inventory_turnover": _safe_div(revenue, inventory),
            "receivables_turnover": _safe_div(revenue, receivables),
        },
        "valuation": {
            "price_to_earnings": _safe_div(market_cap, net_income),
            "price_to_book": _safe_div(market_cap, equity),
            "price_to_sales": _safe_div(market_cap, revenue),
            "ev_to_ebitda": _safe_div(enterprise_value, ebitda),
            "ev_to_revenue": _safe_div(enterprise_value, revenue),
            "fcf_yield": _safe_div(fcf, market_cap),
        },
        "derived": {
            "free_cash_flow": fcf,
            "nopat": nopat,
            "invested_capital": invested_capital,
            "tax_rate_used": tax_rate,
        },
    }
    return json.dumps(out, indent=2)
