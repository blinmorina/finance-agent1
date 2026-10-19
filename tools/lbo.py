"""Institutional LBO model builder for PE Associate workflows.

Integrated Operating → IS → BS → CF → Debt schedule with BoP-based interest
(no intentional circularity). Cash sweep: RCF repayment → TLA prepayment →
TLB prepayment (% of remainder). TLA mandatory amort can be straight % of BoP
or fully amortizing (level principal+interest / sculpted mortgage-style payment).

Mezz: cash + PIK accretion to principal.

Sensitivity grids are pre-computed in Python (each scenario full numerical solve).
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import yfinance as yf
from agents import function_tool
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

from .dcf import OUTPUTS_DIR, _formula as _formula_cell, _input_cell, _label as _label_cell
from .excel import (
    BOLD_FONT,
    FMT_MULT,
    FMT_PCT,
    FMT_USD_M,
    HEADER_FILL,
    HEADER_FONT,
    SUBHEADER_FILL,
    TOTAL_FILL,
    autosize_columns,
)

PROJ_COLS = ["B", "C", "D", "E", "F"]


def _pull_lbo_defaults(ticker: str) -> dict[str, Any]:
    try:
        t = yf.Ticker(ticker.upper())
        info = t.info or {}
        inc = t.income_stmt
        bs = t.balance_sheet

        def _row(df, names: list[str]):
            if df is None or df.empty:
                return None
            for n in names:
                if n in df.index:
                    try:
                        v = df.iloc[df.index.get_loc(n), 0]
                        if v == v:
                            return float(v)
                    except Exception:
                        pass
            return None

        revenue = _row(inc, ["Total Revenue", "Revenue", "Operating Revenue"])
        op_inc = _row(inc, ["Operating Income", "Operating Income Loss"])
        da = _row(inc, ["Reconciled Depreciation", "Depreciation And Amortization"])
        if op_inc and da:
            ebitda_est = op_inc + da
        elif op_inc:
            ebitda_est = op_inc * 1.15
        else:
            ebitda_est = None
        debt = _row(bs, ["Total Debt"]) or 0.0
        cash = _row(bs, ["Cash And Cash Equivalents", "Cash Cash Equivalents And Short Term Investments"]) or 0.0
        ppe = _row(bs, ["Net PPE", "Property Plant And Equipment Net"])
        if ppe is None:
            ppe = (revenue or 500_000_000) * 0.25

        if revenue and ebitda_est:
            margin = max(0.08, min(0.45, ebitda_est / revenue))
        else:
            margin = 0.22
            revenue = revenue or 500_000_000

        rev_growth = float(info.get("revenueGrowth") or 0.07)
        rev_growth = max(0.0, min(0.25, rev_growth))

        return {
            "company_name": info.get("longName") or ticker.upper(),
            "currency": info.get("currency") or "USD",
            "ltm_revenue": float(revenue),
            "ltm_ebitda_margin": margin,
            "existing_net_debt": float(debt - cash),
            "ltm_ppe": float(ppe),
            "rev_growth_default": rev_growth,
        }
    except Exception:
        return {
            "company_name": ticker.upper(),
            "currency": "USD",
            "ltm_revenue": 500_000_000,
            "ltm_ebitda_margin": 0.22,
            "existing_net_debt": 50_000_000,
            "ltm_ppe": 200_000_000,
            "rev_growth_default": 0.07,
        }


def _irr_bisection(cfs: list[float], low: float = -0.99, high: float = 5.0, iters: int = 200) -> float:
    if not cfs or cfs[0] >= 0:
        return float("nan")

    def npv(r: float) -> float:
        return sum(cf / (1 + r) ** i for i, cf in enumerate(cfs))

    if npv(high) > 0:
        while high < 100 and npv(high) > 0:
            high *= 2
    lo, hi = low, high
    for _ in range(iters):
        mid = (lo + hi) / 2
        v = npv(mid)
        if abs(v) < 1e-6:
            return mid
        if v > 0:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def _python_lbo_headline(
    hold_years: int,
    entry_mult: float,
    exit_mult: float,
    ltm_rev: float,
    ltm_ebitda: float,
    growth: list[float],
    margins: list[float],
    debt_to_ebitda: float,
    rcf_draw: float,
    pct_tla: float,
    pct_tlb: float,
    pct_mezz: float,
    tla_rate: float,
    tla_amort_mode: str,
    tla_amort_tenor: int,
    tla_amort: float,
    tlb_rate: float,
    tlb_min_amort: float,
    tlb_sweep_pct: float,
    mezz_cash: float,
    mezz_pik: float,
    rcf_rate: float,
    tax_rate: float,
    fees_pct_ev: float,
    min_cash_pct_rev: float,
    existing_nd: float,
    rollover: float,
    cogs_pct: float,
    da_pct: float,
    capex_maint_pct: float,
    capex_grow_pct: float,
    dso: float,
    dio: float,
    dpo: float,
    mgmt_pool_pct: float,
    mgmt_strike: float,
    perf_pool_pct: float,
    perf_hurdle: float,
) -> dict[str, float]:
    """Numerical path for JSON summary + sensitivity cells (approximation)."""
    hy = min(max(hold_years, 3), 7)
    growth = (growth + [growth[-1]] * hy)[:hy]
    margins = (margins + [margins[-1]] * hy)[:hy]

    entry_ev = entry_mult * ltm_ebitda
    fees = fees_pct_ev * entry_ev
    min_cash_close = min_cash_pct_rev * ltm_rev
    equity_to_sellers = entry_ev - existing_nd
    uses = equity_to_sellers + existing_nd + fees + min_cash_close

    funded_debt = debt_to_ebitda * ltm_ebitda - rcf_draw
    if funded_debt < 0:
        funded_debt = 0.0
    tla0 = funded_debt * pct_tla
    tlb0 = funded_debt * pct_tlb
    mezz0 = funded_debt * pct_mezz

    mode_norm = (tla_amort_mode or "percent_bop").strip().lower().replace("-", "_")
    sculpted = mode_norm in ("fully_amortizing", "sculpted", "mortgage", "level_pay")
    l_pay_fixed = None
    if sculpted:
        nper = max(1, min(int(tla_amort_tenor), hy))
        r_ann = tla_rate
        if tla0 > 0 and nper > 0:
            if r_ann <= 1e-12:
                l_pay_fixed = tla0 / float(nper)
            else:
                l_pay_fixed = tla0 * r_ann / (1 - (1 + r_ann) ** (-float(nper)))
    sources_debt = rcf_draw + tla0 + tlb0 + mezz0
    sponsor_equity = uses - sources_debt - rollover
    if sponsor_equity <= 0:
        sponsor_equity = max(uses * 0.25, 1.0)

    rev = ltm_rev * (1 + growth[0])
    ebitda = rev * margins[0]
    rcf_bal = rcf_draw
    tla_bal = tla0
    tlb_bal = tlb0
    mezz_bal = mezz0
    cash = min_cash_close
    ar_prev = ltm_rev * dso / 365.0
    inv_prev = (ltm_rev * cogs_pct) * dio / 365.0
    ap_prev = (ltm_rev * cogs_pct) * dpo / 365.0
    ppe = ltm_rev * 0.4

    for y in range(hy):
        rev_prev = rev / (1 + growth[y]) if y == 0 else rev / (1 + growth[y])
        if y == 0:
            rev_prev = ltm_rev
        rev = rev_prev * (1 + growth[y])
        ebitda = rev * margins[y]
        da = rev * da_pct
        ebit = ebitda - da
        int_rcf = rcf_bal * rcf_rate
        int_tla = tla_bal * tla_rate
        int_tlb = tlb_bal * tlb_rate
        int_mezz_cash = mezz_bal * mezz_cash
        pik_exp = mezz_bal * mezz_pik
        int_total = int_rcf + int_tla + int_tlb + int_mezz_cash + pik_exp
        ebt = ebit - int_total
        tax = max(ebt, 0) * tax_rate
        ni = ebt - tax

        ar = rev * dso / 365.0
        inv = (rev * cogs_pct) * dio / 365.0
        ap = (rev * cogs_pct) * dpo / 365.0
        delta_nwc = (ar - ar_prev) + (inv - inv_prev) - (ap - ap_prev)
        ar_prev, inv_prev, ap_prev = ar, inv, ap

        capex = rev * (capex_maint_pct + capex_grow_pct)
        cfo = ni + da + pik_exp - delta_nwc - capex

        int_tla_bop = tla_bal * tla_rate
        if sculpted and l_pay_fixed is not None:
            mand_tla = min(max(0.0, l_pay_fixed - int_tla_bop), tla_bal)
        else:
            mand_tla = tla_bal * tla_amort
        mand_tlb = tlb_bal * tlb_min_amort
        mand_total = mand_tla + mand_tlb

        min_floor = rev * min_cash_pct_rev
        avail = cash + cfo - mand_total - min_floor
        avail = max(avail, 0.0)

        rcf_pay = min(rcf_bal, avail)
        avail -= rcf_pay
        tla_pay = min(max(tla_bal - mand_tla, 0), avail)
        avail -= tla_pay
        tlb_pay = min(max(tlb_bal - mand_tlb, 0), avail * tlb_sweep_pct)

        rcf_bal += -rcf_pay
        tla_bal = tla_bal - mand_tla - tla_pay
        tlb_bal = tlb_bal - mand_tlb - tlb_pay
        mezz_bal = mezz_bal + pik_exp

        cash = cash + cfo - mand_total - rcf_pay - tla_pay - tlb_pay

    exit_ebitda = ebitda
    exit_ev = exit_mult * exit_ebitda
    net_debt_exit = rcf_bal + tla_bal + tlb_bal + mezz_bal - cash
    exit_equity = exit_ev - net_debt_exit

    time_vest = min(1.0, hy * 0.25)
    base_opts = max(0.0, exit_equity * mgmt_pool_pct * time_vest - mgmt_strike)
    irr_guess = _irr_bisection([-sponsor_equity] + [0.0] * (hy - 1) + [max(0.0, exit_equity - base_opts)])
    if irr_guess != irr_guess:
        irr_guess = 0.25
    perf_extra = exit_equity * perf_pool_pct if irr_guess >= perf_hurdle else 0.0
    mgmt_total = base_opts + perf_extra + rollover * min(1.5, exit_equity / max(equity_to_sellers, 1))
    sponsor_proceeds = max(exit_equity - mgmt_total, 0.0)

    cfs = [-sponsor_equity] + [0.0] * (hy - 1) + [sponsor_proceeds]
    irr = _irr_bisection(cfs)
    moic = sponsor_proceeds / sponsor_equity if sponsor_equity > 0 else float("nan")

    return {
        "sponsor_equity": sponsor_equity,
        "sponsor_proceeds": sponsor_proceeds,
        "exit_ev": exit_ev,
        "exit_equity": exit_equity,
        "moic": moic,
        "irr": irr,
        "mgmt_waterfall_approx": mgmt_total,
    }


@function_tool(strict_mode=False)
def build_lbo_model(
    ticker: str,
    output_file: str | None = None,
    entry_ebitda_multiple: float | None = None,
    exit_ebitda_multiple: float | None = None,
    hold_years: int = 5,
    ltm_revenue: float | None = None,
    ltm_ebitda_margin: float | None = None,
    ltm_ebitda: float | None = None,
    revenue_growth_rates: list[float] | None = None,
    ebitda_margins: list[float] | None = None,
    cogs_pct_revenue: float = 0.58,
    da_pct_revenue: float = 0.035,
    maint_capex_pct_revenue: float = 0.03,
    growth_capex_pct_revenue: float = 0.015,
    tax_rate: float = 0.21,
    transaction_fees_pct_ev: float = 0.015,
    min_cash_pct_revenue: float = 0.02,
    debt_to_ebitda_entry: float = 5.5,
    rcf_commitment: float = 50_000_000,
    rcf_drawn_at_close: float = 0.0,
    pct_tla_of_term: float = 0.35,
    pct_tlb_of_term: float = 0.45,
    pct_mezz_of_term: float = 0.20,
    tla_rate: float = 0.065,
    tla_amort_pct_per_year: float = 0.10,
    tla_amortization_mode: str = "percent_bop",
    tla_fully_amortizing_tenor_years: int = 5,
    tlb_rate: float = 0.085,
    tlb_min_amort_pct: float = 0.01,
    tlb_sweep_pct_of_excess: float = 0.50,
    mezz_cash_rate: float = 0.10,
    mezz_pik_rate: float = 0.03,
    rcf_rate: float = 0.055,
    rcf_commitment_fee: float = 0.005,
    dso_days: float = 52.0,
    dio_days: float = 42.0,
    dpo_days: float = 38.0,
    management_rollover_equity: float = 15_000_000,
    mgmt_option_pool_pct_exit_equity: float = 0.08,
    mgmt_strike_aggregate: float = 5_000_000,
    perf_pool_pct_if_hurdle_met: float = 0.02,
    sponsor_perf_irr_hurdle: float = 0.20,
    existing_net_debt: float | None = None,
    ltm_ppe: float | None = None,
) -> str:
    """Build a 5-year integrated LBO workbook (Excel formulas + Python sensitivities).

    Monetary inputs are raw dollars. Output: ./outputs/{TICKER}_LBO.xlsx
    """
    d = _pull_lbo_defaults(ticker)
    ltm_rev = float(ltm_revenue if ltm_revenue is not None else d["ltm_revenue"])
    ltm_margin = float(ltm_ebitda_margin if ltm_ebitda_margin is not None else d["ltm_ebitda_margin"])
    ltm_ebitda_calc = ltm_rev * ltm_margin
    ltm_ebitda_use = float(ltm_ebitda) if ltm_ebitda is not None else ltm_ebitda_calc
    existing_nd = float(existing_net_debt if existing_net_debt is not None else d["existing_net_debt"])
    ppe0 = float(ltm_ppe if ltm_ppe is not None else d["ltm_ppe"])

    entry_mult = float(entry_ebitda_multiple if entry_ebitda_multiple is not None else 11.0)
    exit_mult = float(exit_ebitda_multiple if exit_ebitda_multiple is not None else 13.0)

    hy = min(max(int(hold_years), 3), 7)
    hold_years = hy

    g_def = d["rev_growth_default"]
    if revenue_growth_rates is None:
        revenue_growth_rates = [g_def * (1 - i * 0.12) for i in range(5)]
    elif isinstance(revenue_growth_rates, (int, float)):
        revenue_growth_rates = [float(revenue_growth_rates)] * 5
    revenue_growth_rates = list(revenue_growth_rates)[:5]
    while len(revenue_growth_rates) < 5:
        revenue_growth_rates.append(revenue_growth_rates[-1] if revenue_growth_rates else g_def)

    if ebitda_margins is None:
        ebitda_margins = [ltm_margin + min(0.01 * i, 0.03) for i in range(5)]
    elif isinstance(ebitda_margins, (int, float)):
        ebitda_margins = [float(ebitda_margins)] * 5
    ebitda_margins = list(ebitda_margins)[:5]
    while len(ebitda_margins) < 5:
        ebitda_margins.append(ebitda_margins[-1])

    fname = output_file or f"{ticker.upper()}_LBO.xlsx"
    if not fname.endswith(".xlsx"):
        fname += ".xlsx"
    path = Path(fname)
    if not path.is_absolute():
        path = OUTPUTS_DIR / path

    hold_years_excel = min(hold_years, 5)

    _mode_raw = (tla_amortization_mode or "percent_bop").strip().lower().replace("-", "_")
    _tla_py_mode = (
        "fully_amortizing"
        if _mode_raw in ("fully_amortizing", "sculpted", "mortgage", "level_pay")
        else "percent_bop"
    )
    _tla_style_val = 1 if _tla_py_mode == "fully_amortizing" else 0
    _tla_tenor_val = max(1, min(int(tla_fully_amortizing_tenor_years), 10))

    py_headline = _python_lbo_headline(
        hold_years,
        entry_mult,
        exit_mult,
        ltm_rev,
        ltm_ebitda_use,
        revenue_growth_rates,
        ebitda_margins,
        debt_to_ebitda_entry,
        rcf_drawn_at_close,
        pct_tla_of_term,
        pct_tlb_of_term,
        pct_mezz_of_term,
        tla_rate,
        _tla_py_mode,
        _tla_tenor_val,
        tla_amort_pct_per_year,
        tlb_rate,
        tlb_min_amort_pct,
        tlb_sweep_pct_of_excess,
        mezz_cash_rate,
        mezz_pik_rate,
        rcf_rate,
        tax_rate,
        transaction_fees_pct_ev,
        min_cash_pct_revenue,
        existing_nd,
        management_rollover_equity,
        cogs_pct_revenue,
        da_pct_revenue,
        maint_capex_pct_revenue,
        growth_capex_pct_revenue,
        dso_days,
        dio_days,
        dpo_days,
        mgmt_option_pool_pct_exit_equity,
        mgmt_strike_aggregate,
        perf_pool_pct_if_hurdle_met,
        sponsor_perf_irr_hurdle,
    )

    wb = Workbook()
    wb.remove(wb.active)
    co = d["company_name"]
    cur = d["currency"]

    sh_asm = wb.create_sheet("LBO_Assump")
    _ = wb.create_sheet("Sources_Uses")
    _ = wb.create_sheet("Operating")
    _ = wb.create_sheet("IS")
    _ = wb.create_sheet("BS")
    _ = wb.create_sheet("CF")
    _ = wb.create_sheet("Debt_Sched")
    _ = wb.create_sheet("Mgmt_Equity")
    _ = wb.create_sheet("Returns")
    _ = wb.create_sheet("Bridge")
    _ = wb.create_sheet("Sensitivity")
    sh_cover = wb.create_sheet("Cover", 0)

    # --- ASSUMPTIONS (order: override before calculated EBITDA) ---
    a = sh_asm
    a["A1"] = f"{co} — LBO assumptions ({cur})"
    a["A1"].font = Font(size=14, bold=True, color="1F4E78")
    a.merge_cells("A1:G1")
    r = 3
    a[f"A{r}"] = "Transaction"
    a[f"A{r}"].fill = SUBHEADER_FILL
    a[f"A{r}"].font = BOLD_FONT
    r += 1
    _label_cell(a, f"A{r}", "Ticker", bold=True)
    a[f"B{r}"] = ticker.upper()
    r += 1
    _label_cell(a, f"A{r}", "Entry EBITDA multiple", bold=True)
    _input_cell(a, f"B{r}", entry_mult, FMT_MULT)
    row_entry_mult = r
    r += 1
    _label_cell(a, f"A{r}", "Exit EBITDA multiple", bold=True)
    _input_cell(a, f"B{r}", exit_mult, FMT_MULT)
    row_exit_mult = r
    r += 1
    _label_cell(a, f"A{r}", "Hold period (years)", bold=True)
    _input_cell(a, f"B{r}", hold_years, "0")
    row_hold = r
    r += 1
    _label_cell(a, f"A{r}", "Transaction fees (% of EV)", bold=True)
    _input_cell(a, f"B{r}", transaction_fees_pct_ev, FMT_PCT)
    row_fees_pct = r
    r += 1
    _label_cell(a, f"A{r}", "Existing net debt refinanced ($)", bold=True)
    _input_cell(a, f"B{r}", existing_nd, FMT_USD_M)
    row_exist_nd = r
    r += 1
    _label_cell(a, f"A{r}", "Min cash (% of LTM revenue)", bold=True)
    _input_cell(a, f"B{r}", min_cash_pct_revenue, FMT_PCT)
    row_min_cash_pct = r
    r += 2

    a[f"A{r}"] = "Operating — LTM"
    a[f"A{r}"].fill = SUBHEADER_FILL
    a[f"A{r}"].font = BOLD_FONT
    r += 1
    _label_cell(a, f"A{r}", "LTM revenue ($)", bold=True)
    _input_cell(a, f"B{r}", ltm_rev, FMT_USD_M)
    row_ltm_rev = r
    r += 1
    _label_cell(a, f"A{r}", "LTM EBITDA margin", bold=True)
    _input_cell(a, f"B{r}", ltm_margin, FMT_PCT)
    row_ltm_margin = r
    r += 1
    _label_cell(a, f"A{r}", "LTM EBITDA override ($, 0 = calc from rev × margin)", bold=True)
    _input_cell(a, f"B{r}", float(ltm_ebitda) if ltm_ebitda is not None else 0.0, FMT_USD_M)
    row_ltm_ebitda_override = r
    r += 1
    _label_cell(a, f"A{r}", "LTM EBITDA ($)", bold=True)
    _formula_cell(a, f"B{r}", f"=IF(B{row_ltm_ebitda_override}>0,B{row_ltm_ebitda_override},B{row_ltm_rev}*B{row_ltm_margin})", FMT_USD_M, bold=True)
    row_ltm_ebitda = r
    r += 2

    a[f"A{r}"] = "Forecast drivers (Y1–Y5)"
    a[f"A{r}"].fill = SUBHEADER_FILL
    a.merge_cells(f"A{r}:F{r}")
    a[f"A{r}"].font = BOLD_FONT
    r += 1
    _label_cell(a, f"A{r}", "Period")
    for i, col in enumerate(PROJ_COLS):
        a[f"{col}{r}"] = f"Year {i+1}"
        a[f"{col}{r}"].font = HEADER_FONT
        a[f"{col}{r}"].fill = HEADER_FILL
        a[f"{col}{r}"].alignment = Alignment(horizontal="center")
    r += 1
    _label_cell(a, f"A{r}", "Revenue growth", bold=True)
    row_growth = r
    for i, col in enumerate(PROJ_COLS):
        _input_cell(a, f"{col}{r}", float(revenue_growth_rates[i]), FMT_PCT)
    r += 1
    _label_cell(a, f"A{r}", "EBITDA margin", bold=True)
    row_margin_row = r
    for i, col in enumerate(PROJ_COLS):
        _input_cell(a, f"{col}{r}", float(ebitda_margins[i]), FMT_PCT)
    r += 2

    a[f"A{r}"] = "Margins / capex / tax"
    a[f"A{r}"].fill = SUBHEADER_FILL
    a[f"A{r}"].font = BOLD_FONT
    r += 1
    _label_cell(a, f"A{r}", "COGS % revenue", bold=True)
    _input_cell(a, f"B{r}", cogs_pct_revenue, FMT_PCT)
    row_cogs = r
    r += 1
    _label_cell(a, f"A{r}", "D&A % revenue", bold=True)
    _input_cell(a, f"B{r}", da_pct_revenue, FMT_PCT)
    row_da_pct = r
    r += 1
    _label_cell(a, f"A{r}", "Maintenance capex % revenue", bold=True)
    _input_cell(a, f"B{r}", maint_capex_pct_revenue, FMT_PCT)
    row_capex_maint = r
    r += 1
    _label_cell(a, f"A{r}", "Growth capex % revenue", bold=True)
    _input_cell(a, f"B{r}", growth_capex_pct_revenue, FMT_PCT)
    row_capex_growth = r
    r += 1
    _label_cell(a, f"A{r}", "Effective tax rate", bold=True)
    _input_cell(a, f"B{r}", tax_rate, FMT_PCT)
    row_tax = r
    r += 2

    a[f"A{r}"] = "Working capital (days) / PP&E"
    a[f"A{r}"].fill = SUBHEADER_FILL
    r += 1
    _label_cell(a, f"A{r}", "DSO", bold=True)
    _input_cell(a, f"B{r}", dso_days, "0.0")
    row_dso = r
    r += 1
    _label_cell(a, f"A{r}", "DIO", bold=True)
    _input_cell(a, f"B{r}", dio_days, "0.0")
    row_dio = r
    r += 1
    _label_cell(a, f"A{r}", "DPO", bold=True)
    _input_cell(a, f"B{r}", dpo_days, "0.0")
    row_dpo = r
    r += 1
    _label_cell(a, f"A{r}", "Opening PP&E ($)", bold=True)
    _input_cell(a, f"B{r}", ppe0, FMT_USD_M)
    row_ppe0 = r
    r += 2

    a[f"A{r}"] = "Debt structure"
    a[f"A{r}"].fill = SUBHEADER_FILL
    a[f"A{r}"].font = BOLD_FONT
    r += 1
    _label_cell(a, f"A{r}", "Leverage (funded debt / LTM EBITDA)", bold=True)
    _input_cell(a, f"B{r}", debt_to_ebitda_entry, FMT_MULT)
    row_lev = r
    r += 1
    _label_cell(a, f"A{r}", "RCF commitment cap ($) — disclosure only", bold=True)
    _input_cell(a, f"B{r}", rcf_commitment, FMT_USD_M)
    row_rcf_cap = r
    r += 1
    _label_cell(a, f"A{r}", "RCF drawn at close ($)", bold=True)
    _input_cell(a, f"B{r}", rcf_drawn_at_close, FMT_USD_M)
    row_rcf_draw0 = r
    r += 1
    for lbl, pct, rw in [("TLA % of term debt", pct_tla_of_term, "row_pct_tla"), ("TLB % of term debt", pct_tlb_of_term, "row_pct_tlb"), ("Mezz % of term debt", pct_mezz_of_term, "row_pct_mezz")]:
        _label_cell(a, f"A{r}", lbl, bold=True)
        _input_cell(a, f"B{r}", pct, FMT_PCT)
        if rw == "row_pct_tla":
            row_pct_tla = r
        elif rw == "row_pct_tlb":
            row_pct_tlb = r
        else:
            row_pct_mezz = r
        r += 1
    _label_cell(a, f"A{r}", "TLA coupon", bold=True)
    _input_cell(a, f"B{r}", tla_rate, FMT_PCT)
    row_tla_rate = r
    r += 1
    _label_cell(a, f"A{r}", "TLA amort % BoP / yr", bold=True)
    _input_cell(a, f"B{r}", tla_amort_pct_per_year, FMT_PCT)
    row_tla_amort = r
    r += 1
    _label_cell(a, f"A{r}", "TLA amort mode (0 = % BoP, 1 = fully amortizing)", bold=True)
    _input_cell(a, f"B{r}", float(_tla_style_val), "0")
    row_tla_mode = r
    r += 1
    _label_cell(a, f"A{r}", "TLA amort tenor (yrs, mode 1)", bold=True)
    _input_cell(a, f"B{r}", float(_tla_tenor_val), "0")
    row_tla_tenor = r
    r += 1
    _label_cell(a, f"A{r}", "TLA fixed debt service ($/yr if mode 1)", bold=True)
    _formula_cell(
        a,
        f"B{r}",
        f"=IF(B{row_tla_mode}=1,IF(B{row_tla_rate}=0,Sources_Uses!$D$5/B{row_tla_tenor},"
        f"-PMT(B{row_tla_rate},B{row_tla_tenor},Sources_Uses!$D$5)),0)",
        FMT_USD_M,
        bold=True,
    )
    row_tla_level_ds = r
    r += 1
    _label_cell(a, f"A{r}", "TLB coupon", bold=True)
    _input_cell(a, f"B{r}", tlb_rate, FMT_PCT)
    row_tlb_rate = r
    r += 1
    _label_cell(a, f"A{r}", "TLB min amort % BoP", bold=True)
    _input_cell(a, f"B{r}", tlb_min_amort_pct, FMT_PCT)
    row_tlb_min_amort = r
    r += 1
    _label_cell(a, f"A{r}", "TLB excess cash sweep %", bold=True)
    _input_cell(a, f"B{r}", tlb_sweep_pct_of_excess, FMT_PCT)
    row_tlb_sweep_pct = r
    r += 1
    _label_cell(a, f"A{r}", "Mezz cash coupon", bold=True)
    _input_cell(a, f"B{r}", mezz_cash_rate, FMT_PCT)
    row_mezz_cash = r
    r += 1
    _label_cell(a, f"A{r}", "Mezz PIK coupon", bold=True)
    _input_cell(a, f"B{r}", mezz_pik_rate, FMT_PCT)
    row_mezz_pik = r
    r += 1
    _label_cell(a, f"A{r}", "RCF drawn coupon", bold=True)
    _input_cell(a, f"B{r}", rcf_rate, FMT_PCT)
    row_rcf_rate = r
    r += 1
    _label_cell(a, f"A{r}", "RCF commitment fee (unused)", bold=True)
    _input_cell(a, f"B{r}", rcf_commitment_fee, FMT_PCT)
    row_rcf_fee = r
    r += 2

    a[f"A{r}"] = "Management equity"
    a[f"A{r}"].fill = SUBHEADER_FILL
    r += 1
    _label_cell(a, f"A{r}", "Management rollover equity ($)", bold=True)
    _input_cell(a, f"B{r}", management_rollover_equity, FMT_USD_M)
    row_m_roll = r
    r += 1
    _label_cell(a, f"A{r}", "Mgmt option pool (% exit equity)", bold=True)
    _input_cell(a, f"B{r}", mgmt_option_pool_pct_exit_equity, FMT_PCT)
    row_m_pool = r
    r += 1
    _label_cell(a, f"A{r}", "Mgmt aggregate strike ($)", bold=True)
    _input_cell(a, f"B{r}", mgmt_strike_aggregate, FMT_USD_M)
    row_m_strike = r
    r += 1
    _label_cell(a, f"A{r}", "Perf pool (% exit equity if IRR hurdle met)", bold=True)
    _input_cell(a, f"B{r}", perf_pool_pct_if_hurdle_met, FMT_PCT)
    row_perf_pool = r
    r += 1
    _label_cell(a, f"A{r}", "Sponsor IRR hurdle for perf pool", bold=True)
    _input_cell(a, f"B{r}", sponsor_perf_irr_hurdle, FMT_PCT)
    row_perf_hurdle = r

    autosize_columns(a)

    # --- Sources & Uses ---
    su = wb["Sources_Uses"]
    su["A1"] = "Sources & Uses (close)"
    su["A1"].font = Font(size=14, bold=True)
    su.merge_cells("A1:D1")
    _label_cell(su, "A3", "USES", bold=True)
    _label_cell(su, "C3", "SOURCES", bold=True)
    _label_cell(su, "A4", "Equity consideration (cash to sellers)")
    su["B4"] = f"=LBO_Assump!B{row_entry_mult}*LBO_Assump!B{row_ltm_ebitda}-LBO_Assump!B{row_exist_nd}"
    su["B4"].number_format = FMT_USD_M
    _label_cell(su, "A5", "Refinance existing net debt")
    _formula_cell(su, "B5", f"=LBO_Assump!B{row_exist_nd}", FMT_USD_M)
    _label_cell(su, "A6", "Transaction fees")
    _formula_cell(su, "B6", f"=LBO_Assump!B{row_entry_mult}*LBO_Assump!B{row_ltm_ebitda}*LBO_Assump!B{row_fees_pct}", FMT_USD_M)
    _label_cell(su, "A7", "Minimum cash at close")
    _formula_cell(su, "B7", f"=LBO_Assump!B{row_ltm_rev}*LBO_Assump!B{row_min_cash_pct}", FMT_USD_M)
    _label_cell(su, "A8", "Total Uses")
    _formula_cell(su, "B8", "=SUM(B4:B7)", FMT_USD_M, bold=True)
    su["B8"].fill = TOTAL_FILL

    _label_cell(su, "C4", "RCF drawn at close")
    _formula_cell(su, "D4", f"=LBO_Assump!B{row_rcf_draw0}", FMT_USD_M)
    _label_cell(su, "C5", "New TLA")
    _formula_cell(
        su,
        "D5",
        f"=(LBO_Assump!B{row_lev}*LBO_Assump!B{row_ltm_ebitda}-LBO_Assump!B{row_rcf_draw0})*LBO_Assump!B{row_pct_tla}",
        FMT_USD_M,
    )
    _label_cell(su, "C6", "New TLB")
    _formula_cell(
        su,
        "D6",
        f"=(LBO_Assump!B{row_lev}*LBO_Assump!B{row_ltm_ebitda}-LBO_Assump!B{row_rcf_draw0})*LBO_Assump!B{row_pct_tlb}",
        FMT_USD_M,
    )
    _label_cell(su, "C7", "New Mezzanine")
    _formula_cell(
        su,
        "D7",
        f"=(LBO_Assump!B{row_lev}*LBO_Assump!B{row_ltm_ebitda}-LBO_Assump!B{row_rcf_draw0})*LBO_Assump!B{row_pct_mezz}",
        FMT_USD_M,
    )
    _label_cell(su, "C8", "Management rollover")
    _formula_cell(su, "D8", f"=LBO_Assump!B{row_m_roll}", FMT_USD_M)
    _label_cell(su, "C9", "Sponsor equity (plug)")
    _formula_cell(su, "D9", "=B8-SUM(D4:D8)", FMT_USD_M, bold=True)
    su["D9"].fill = TOTAL_FILL
    _label_cell(su, "C10", "Total Sources")
    _formula_cell(su, "D10", "=SUM(D4:D9)", FMT_USD_M, bold=True)
    _label_cell(su, "A11", "Sources − Uses check")
    _formula_cell(su, "B11", "=D10-B8", FMT_USD_M)
    autosize_columns(su)

    # --- Operating ---
    op = wb["Operating"]
    op["A1"] = "Operating model ($)"
    op["A1"].font = Font(size=14, bold=True)
    _label_cell(op, "A3", "Line item")
    for i, col in enumerate(PROJ_COLS):
        op[f"{col}3"] = f"Y{i+1}"
        op[f"{col}3"].font = HEADER_FONT
        op[f"{col}3"].fill = HEADER_FILL
    _formula_cell(op, "B4", f"=LBO_Assump!B{row_ltm_rev}*(1+LBO_Assump!B{row_growth})", FMT_USD_M, bold=True)
    for i in range(1, 5):
        pc, pp = PROJ_COLS[i], PROJ_COLS[i - 1]
        _formula_cell(op, f"{pc}4", f"={pp}4*(1+LBO_Assump!{pc}{row_growth})", FMT_USD_M, bold=True)
    for col in PROJ_COLS:
        _formula_cell(op, f"{col}5", f"={col}4*LBO_Assump!{col}{row_margin_row}", FMT_USD_M)
    for col in PROJ_COLS:
        _formula_cell(op, f"{col}6", f"={col}4*LBO_Assump!B{row_cogs}", FMT_USD_M)
    for col in PROJ_COLS:
        _formula_cell(op, f"{col}7", f"={col}4-{col}5-{col}6", FMT_USD_M)
    for col in PROJ_COLS:
        _formula_cell(op, f"{col}8", f"={col}4*LBO_Assump!B{row_da_pct}", FMT_USD_M)
    for col in PROJ_COLS:
        _formula_cell(op, f"{col}9", f"={col}5-{col}8", FMT_USD_M, bold=True)
    autosize_columns(op)

    # --- Debt schedule (rows fixed for cross-sheet refs) ---
    ds = wb["Debt_Sched"]
    ds["A1"] = "Debt schedule ($) — BoP balances; sweep from CF"
    ds["A1"].font = Font(size=14, bold=True)
    hdr = 3
    _label_cell(ds, f"A{hdr}", "Line item")
    for i, col in enumerate(PROJ_COLS):
        ds[f"{col}{hdr}"] = f"Y{i+1}"
        ds[f"{col}{hdr}"].font = HEADER_FONT
        ds[f"{col}{hdr}"].fill = HEADER_FILL

    # RCF 5-8: BoP, net change (negative=repay), EoP, interest expense (cash)
    _label_cell(ds, "A5", "RCF — BoP balance", bold=True)
    ds["B5"] = "=Sources_Uses!D4"
    ds["B5"].number_format = FMT_USD_M
    for i in range(1, 5):
        ds[f"{PROJ_COLS[i]}5"] = f"={PROJ_COLS[i - 1]}8"
        ds[f"{PROJ_COLS[i]}5"].number_format = FMT_USD_M
    _label_cell(ds, "A6", "RCF — Repayment / (draw)")
    for col in PROJ_COLS:
        _formula_cell(ds, f"{col}6", f"=-CF!{col}24", FMT_USD_M)
    _label_cell(ds, "A7", "RCF — BoP × coupon (cash interest)")
    for col in PROJ_COLS:
        _formula_cell(ds, f"{col}7", f"={col}5*LBO_Assump!B{row_rcf_rate}", FMT_USD_M)
    _label_cell(ds, "A8", "RCF — EoP balance")
    for col in PROJ_COLS:
        _formula_cell(ds, f"{col}8", f"={col}5+{col}6", FMT_USD_M)

    # TLA 10-13
    _label_cell(ds, "A10", "TLA — BoP", bold=True)
    ds["B10"] = "=Sources_Uses!D5"
    ds["B10"].number_format = FMT_USD_M
    for i in range(1, 5):
        ds[f"{PROJ_COLS[i]}10"] = f"={PROJ_COLS[i - 1]}13"
        ds[f"{PROJ_COLS[i]}10"].number_format = FMT_USD_M
    _label_cell(ds, "A11", "TLA — Mandatory amort")
    for col in PROJ_COLS:
        _formula_cell(
            ds,
            f"{col}11",
            f"=IF(LBO_Assump!B{row_tla_mode}=1,"
            f"MIN(MAX(0,LBO_Assump!B{row_tla_level_ds}-{col}14),{col}10),"
            f"LBO_Assump!B{row_tla_amort}*{col}10)",
            FMT_USD_M,
        )
    _label_cell(ds, "A12", "TLA — Cash sweep prepay")
    for col in PROJ_COLS:
        _formula_cell(ds, f"{col}12", f"=CF!{col}25", FMT_USD_M)
    _label_cell(ds, "A13", "TLA — EoP balance")
    for col in PROJ_COLS:
        _formula_cell(ds, f"{col}13", f"={col}10-{col}11-{col}12", FMT_USD_M)

    _label_cell(ds, "A14", "TLA — BoP × coupon")
    for col in PROJ_COLS:
        _formula_cell(ds, f"{col}14", f"={col}10*LBO_Assump!B{row_tla_rate}", FMT_USD_M)

    # TLB 16-20
    _label_cell(ds, "A16", "TLB — BoP", bold=True)
    ds["B16"] = "=Sources_Uses!D6"
    ds["B16"].number_format = FMT_USD_M
    for i in range(1, 5):
        ds[f"{PROJ_COLS[i]}16"] = f"={PROJ_COLS[i - 1]}20"
        ds[f"{PROJ_COLS[i]}16"].number_format = FMT_USD_M
    _label_cell(ds, "A17", "TLB — Mandatory amort")
    for col in PROJ_COLS:
        _formula_cell(ds, f"{col}17", f"={col}16*LBO_Assump!B{row_tlb_min_amort}", FMT_USD_M)
    _label_cell(ds, "A18", "TLB — Cash sweep prepay")
    for col in PROJ_COLS:
        _formula_cell(ds, f"{col}18", f"=CF!{col}26", FMT_USD_M)
    _label_cell(ds, "A19", "TLB — EoP balance")
    for col in PROJ_COLS:
        _formula_cell(ds, f"{col}19", f"={col}16-{col}17-{col}18", FMT_USD_M)
    _label_cell(ds, "A20", "TLB — BoP × coupon")
    for col in PROJ_COLS:
        _formula_cell(ds, f"{col}20", f"={col}16*LBO_Assump!B{row_tlb_rate}", FMT_USD_M)

    # Mezz 22-26
    _label_cell(ds, "A22", "Mezz — BoP", bold=True)
    ds["B22"] = "=Sources_Uses!D7"
    ds["B22"].number_format = FMT_USD_M
    for i in range(1, 5):
        ds[f"{PROJ_COLS[i]}22"] = f"={PROJ_COLS[i - 1]}26"
        ds[f"{PROJ_COLS[i]}22"].number_format = FMT_USD_M
    _label_cell(ds, "A23", "Mezz — PIK accretion (adds to principal)")
    for col in PROJ_COLS:
        _formula_cell(ds, f"{col}23", f"={col}22*LBO_Assump!B{row_mezz_pik}", FMT_USD_M)
    _label_cell(ds, "A24", "Mezz — BoP × cash coupon (expense)")
    for col in PROJ_COLS:
        _formula_cell(ds, f"{col}24", f"={col}22*LBO_Assump!B{row_mezz_cash}", FMT_USD_M)
    _label_cell(ds, "A25", "Mezz — PIK expense (= accretion)")
    for col in PROJ_COLS:
        _formula_cell(ds, f"{col}25", f"={col}23", FMT_USD_M)
    _label_cell(ds, "A26", "Mezz — EoP balance")
    for col in PROJ_COLS:
        _formula_cell(ds, f"{col}26", f"={col}22+{col}23", FMT_USD_M)

    _label_cell(ds, "A28", "Total cash interest expense")
    for col in PROJ_COLS:
        _formula_cell(ds, f"{col}28", f"={col}7+{col}14+{col}20+{col}24", FMT_USD_M)
    _label_cell(ds, "A29", "Total PIK interest expense")
    for col in PROJ_COLS:
        _formula_cell(ds, f"{col}29", f"={col}25", FMT_USD_M)

    _label_cell(ds, "A31", "Mandatory principal (TLA+TLB)")
    for col in PROJ_COLS:
        _formula_cell(ds, f"{col}31", f"={col}11+{col}17", FMT_USD_M)

    autosize_columns(ds)

    # --- Income statement ---
    ist = wb["IS"]
    ist["A1"] = "Income Statement ($)"
    ist["A1"].font = Font(size=14, bold=True)
    _label_cell(ist, "A3", "Line item")
    for i, col in enumerate(PROJ_COLS):
        ist[f"{col}3"] = f"Y{i+1}"
        ist[f"{col}3"].font = HEADER_FONT
        ist[f"{col}3"].fill = HEADER_FILL
    for col in PROJ_COLS:
        _formula_cell(ist, f"{col}4", f"=Operating!{col}9", FMT_USD_M, bold=True)
    for col in PROJ_COLS:
        _formula_cell(ist, f"{col}5", f"=Debt_Sched!{col}28", FMT_USD_M)
    for col in PROJ_COLS:
        _formula_cell(ist, f"{col}6", f"=Debt_Sched!{col}29", FMT_USD_M)
    for col in PROJ_COLS:
        _formula_cell(ist, f"{col}7", f"={col}4-{col}5-{col}6", FMT_USD_M)
    for col in PROJ_COLS:
        _formula_cell(ist, f"{col}8", f"=-MAX({col}7,0)*LBO_Assump!B{row_tax}", FMT_USD_M)
    for col in PROJ_COLS:
        _formula_cell(ist, f"{col}9", f"={col}7+{col}8", FMT_USD_M, bold=True)
        ist[f"{col}9"].fill = TOTAL_FILL
    autosize_columns(ist)

    # --- Balance sheet ---
    bs = wb["BS"]
    bs["A1"] = "Balance Sheet ($)"
    bs["A1"].font = Font(size=14, bold=True)
    _label_cell(bs, "A3", "Line item")
    for i, col in enumerate(PROJ_COLS):
        bs[f"{col}3"] = f"Y{i+1} end"
        bs[f"{col}3"].font = HEADER_FONT
        bs[f"{col}3"].fill = HEADER_FILL

    _label_cell(bs, "A4", "Cash")
    _formula_cell(bs, "B4", "=CF!B36", FMT_USD_M)
    for i in range(1, 5):
        bs[f"{PROJ_COLS[i]}4"] = f"=CF!{PROJ_COLS[i]}36"
        bs[f"{PROJ_COLS[i]}4"].number_format = FMT_USD_M

    _label_cell(bs, "A5", "Accounts receivable")
    _formula_cell(bs, "B5", f"=Operating!B4*LBO_Assump!B{row_dso}/365", FMT_USD_M)
    for i, col in enumerate(PROJ_COLS):
        _formula_cell(bs, f"{col}5", f"=Operating!{col}4*LBO_Assump!B{row_dso}/365", FMT_USD_M)

    _label_cell(bs, "A6", "Inventory")
    _formula_cell(bs, "B6", f"=Operating!B6*LBO_Assump!B{row_dio}/365", FMT_USD_M)
    for col in PROJ_COLS:
        _formula_cell(bs, f"{col}6", f"=Operating!{col}6*LBO_Assump!B{row_dio}/365", FMT_USD_M)

    _label_cell(bs, "A7", "Total current assets")
    for col in PROJ_COLS:
        _formula_cell(bs, f"{col}7", f"=SUM({col}4:{col}6)", FMT_USD_M)

    _label_cell(bs, "A8", "PP&E net")
    _formula_cell(bs, "B8", f"=LBO_Assump!B{row_ppe0}+Operating!B4*(LBO_Assump!B{row_capex_maint}+LBO_Assump!B{row_capex_growth})-Operating!B8", FMT_USD_M)
    for i, col in enumerate(PROJ_COLS):
        prev = PROJ_COLS[i - 1] if i > 0 else None
        if prev is None:
            _formula_cell(bs, f"{col}8", f"=B8+Operating!{col}4*(LBO_Assump!B{row_capex_maint}+LBO_Assump!B{row_capex_growth})-Operating!{col}8", FMT_USD_M)
        else:
            _formula_cell(bs, f"{col}8", f"={prev}8+Operating!{col}4*(LBO_Assump!B{row_capex_maint}+LBO_Assump!B{row_capex_growth})-Operating!{col}8", FMT_USD_M)

    _label_cell(bs, "A9", "Total assets")
    for col in PROJ_COLS:
        _formula_cell(bs, f"{col}9", f"={col}7+{col}8", FMT_USD_M, bold=True)

    _label_cell(bs, "A11", "Accounts payable")
    _formula_cell(bs, "B11", f"=Operating!B6*LBO_Assump!B{row_dpo}/365", FMT_USD_M)
    for col in PROJ_COLS:
        _formula_cell(bs, f"{col}11", f"=Operating!{col}6*LBO_Assump!B{row_dpo}/365", FMT_USD_M)

    _label_cell(bs, "A12", "RCF")
    for col in PROJ_COLS:
        _formula_cell(bs, f"{col}12", f"=Debt_Sched!{col}8", FMT_USD_M)
    _label_cell(bs, "A13", "TLA")
    for col in PROJ_COLS:
        _formula_cell(bs, f"{col}13", f"=Debt_Sched!{col}13", FMT_USD_M)
    _label_cell(bs, "A14", "TLB")
    for col in PROJ_COLS:
        _formula_cell(bs, f"{col}14", f"=Debt_Sched!{col}19", FMT_USD_M)
    _label_cell(bs, "A15", "Mezzanine")
    for col in PROJ_COLS:
        _formula_cell(bs, f"{col}15", f"=Debt_Sched!{col}26", FMT_USD_M)

    _label_cell(bs, "A16", "Total debt")
    for col in PROJ_COLS:
        _formula_cell(bs, f"{col}16", f"=SUM({col}12:{col}15)", FMT_USD_M)

    _label_cell(bs, "A18", "Sponsor equity")
    for col in PROJ_COLS:
        _formula_cell(bs, f"{col}18", "=Sources_Uses!D9", FMT_USD_M)
    _label_cell(bs, "A19", "Management rollover")
    for col in PROJ_COLS:
        _formula_cell(bs, f"{col}19", "=Sources_Uses!D8", FMT_USD_M)
    _label_cell(bs, "A20", "Retained earnings")
    _formula_cell(bs, "B20", "=IS!B9", FMT_USD_M)
    for i in range(1, 5):
        pc, pp = PROJ_COLS[i], PROJ_COLS[i - 1]
        _formula_cell(bs, f"{pc}20", f"={pp}20+IS!{pc}9", FMT_USD_M)

    _label_cell(bs, "A21", "Total equity")
    for col in PROJ_COLS:
        _formula_cell(bs, f"{col}21", f"=SUM({col}18:{col}20)", FMT_USD_M)

    _label_cell(bs, "A22", "Total liabilities & equity")
    for col in PROJ_COLS:
        _formula_cell(bs, f"{col}22", f"={col}11+{col}16+{col}21", FMT_USD_M, bold=True)

    _label_cell(bs, "A23", "Assets − Liab & Eq (should be 0)")
    for col in PROJ_COLS:
        _formula_cell(bs, f"{col}23", f"={col}9-{col}22", FMT_USD_M)
    autosize_columns(bs)

    # --- Cash flow ---
    cf = wb["CF"]
    cf["A1"] = "Cash Flow ($)"
    cf["A1"].font = Font(size=14, bold=True)
    _label_cell(cf, "A3", "Line item")
    for i, col in enumerate(PROJ_COLS):
        cf[f"{col}3"] = f"Y{i+1}"
        cf[f"{col}3"].font = HEADER_FONT
        cf[f"{col}3"].fill = HEADER_FILL

    for col in PROJ_COLS:
        _formula_cell(cf, f"{col}4", f"=IS!{col}9", FMT_USD_M)
    for col in PROJ_COLS:
        _formula_cell(cf, f"{col}5", f"=Operating!{col}8", FMT_USD_M)
    for col in PROJ_COLS:
        _formula_cell(cf, f"{col}6", f"=Debt_Sched!{col}29", FMT_USD_M)

    _label_cell(cf, "A7", "Δ Accounts receivable")
    _formula_cell(cf, "B7", f"=-(BS!B5-LBO_Assump!B{row_ltm_rev}*LBO_Assump!B{row_dso}/365)", FMT_USD_M)
    for i in range(1, 5):
        pc, pp = PROJ_COLS[i], PROJ_COLS[i - 1]
        _formula_cell(cf, f"{pc}7", f"=-({pc}5-{pp}5)", FMT_USD_M)

    _label_cell(cf, "A8", "Δ Inventory")
    _formula_cell(cf, "B8", f"=-(BS!B6-LBO_Assump!B{row_ltm_rev}*LBO_Assump!B{row_cogs}*LBO_Assump!B{row_dio}/365)", FMT_USD_M)
    for i in range(1, 5):
        pc, pp = PROJ_COLS[i], PROJ_COLS[i - 1]
        _formula_cell(cf, f"{pc}8", f"=-({pc}6-{pp}6)", FMT_USD_M)

    _label_cell(cf, "A9", "Δ Accounts payable")
    _formula_cell(cf, "B9", f"=(BS!B11-LBO_Assump!B{row_ltm_rev}*LBO_Assump!B{row_cogs}*LBO_Assump!B{row_dpo}/365)", FMT_USD_M)
    for i in range(1, 5):
        pc, pp = PROJ_COLS[i], PROJ_COLS[i - 1]
        _formula_cell(cf, f"{pc}9", f"=({pc}11-{pp}11)", FMT_USD_M)

    _label_cell(cf, "A10", "Maintenance + growth capex (cash)")
    for col in PROJ_COLS:
        _formula_cell(cf, f"{col}10", f"=-Operating!{col}4*(LBO_Assump!B{row_capex_maint}+LBO_Assump!B{row_capex_growth})", FMT_USD_M)

    _label_cell(cf, "A11", "Cash from operations (before debt)")
    for col in PROJ_COLS:
        _formula_cell(cf, f"{col}11", f"={col}4+{col}5+{col}6+{col}7+{col}8+{col}9+{col}10", FMT_USD_M)

    _label_cell(cf, "A12", "Beginning cash")
    _formula_cell(cf, "B12", "=Sources_Uses!B7", FMT_USD_M)
    for i in range(1, 5):
        pc, pp = PROJ_COLS[i], PROJ_COLS[i - 1]
        _formula_cell(cf, f"{pc}12", f"={pp}36", FMT_USD_M)

    _label_cell(cf, "A13", "Minimum cash floor (% of revenue)")
    for col in PROJ_COLS:
        _formula_cell(cf, f"{col}13", f"=Operating!{col}4*LBO_Assump!B{row_min_cash_pct}", FMT_USD_M)

    _label_cell(cf, "A14", "Mandatory principal")
    for col in PROJ_COLS:
        _formula_cell(cf, f"{col}14", f"=Debt_Sched!{col}31", FMT_USD_M)

    _label_cell(cf, "A15", "Excess cash before sweep")
    for col in PROJ_COLS:
        _formula_cell(cf, f"{col}15", f"=MAX(0,{col}12+{col}11-{col}14-{col}13)", FMT_USD_M)

    _label_cell(cf, "A16", "After RCF repayment capacity")
    for col in PROJ_COLS:
        _formula_cell(cf, f"{col}16", f"={col}15", FMT_USD_M)

    _label_cell(cf, "A17", "RCF repayment")
    for col in PROJ_COLS:
        _formula_cell(cf, f"{col}17", f"=MIN(MAX(0,Debt_Sched!{col}5),{col}16)", FMT_USD_M)

    _label_cell(cf, "A18", "Remainder after RCF")
    for col in PROJ_COLS:
        _formula_cell(cf, f"{col}18", f"=MAX(0,{col}16-{col}17)", FMT_USD_M)

    _label_cell(cf, "A19", "TLA sweep prepay")
    for col in PROJ_COLS:
        _formula_cell(cf, f"{col}19", f"=MIN(MAX(0,Debt_Sched!{col}10-Debt_Sched!{col}11),{col}18)", FMT_USD_M)

    _label_cell(cf, "A20", "Remainder after TLA sweep")
    for col in PROJ_COLS:
        _formula_cell(cf, f"{col}20", f"=MAX(0,{col}18-{col}19)", FMT_USD_M)

    _label_cell(cf, "A21", "TLB sweep prepay (% of remainder)")
    for col in PROJ_COLS:
        _formula_cell(
            cf,
            f"{col}21",
            f"=MIN(MAX(0,Debt_Sched!{col}16-Debt_Sched!{col}17),{col}20*LBO_Assump!B{row_tlb_sweep_pct})",
            FMT_USD_M,
        )

    _label_cell(cf, "A23", "Total voluntary debt paydown (RCF)")
    for col in PROJ_COLS:
        _formula_cell(cf, f"{col}23", f"={col}17", FMT_USD_M)

    _label_cell(cf, "A24", "RCF voluntary line (= row 17)")
    for col in PROJ_COLS:
        _formula_cell(cf, f"{col}24", f"={col}17", FMT_USD_M)
    _label_cell(cf, "A25", "TLA voluntary line")
    for col in PROJ_COLS:
        _formula_cell(cf, f"{col}25", f"={col}19", FMT_USD_M)
    _label_cell(cf, "A26", "TLB voluntary line")
    for col in PROJ_COLS:
        _formula_cell(cf, f"{col}26", f"={col}21", FMT_USD_M)

    _label_cell(cf, "A28", "Net change in cash")
    for col in PROJ_COLS:
        _formula_cell(cf, f"{col}28", f"={col}11-{col}14-{col}17-{col}19-{col}21", FMT_USD_M)

    _label_cell(cf, "A36", "Ending cash")
    _formula_cell(cf, "B36", "=B12+B28", FMT_USD_M)
    for i in range(1, 5):
        pc, pp = PROJ_COLS[i], PROJ_COLS[i - 1]
        _formula_cell(cf, f"{pc}36", f"={pc}12+{pc}28", FMT_USD_M)

    autosize_columns(cf)

    # --- Management equity ---
    mg = wb["Mgmt_Equity"]
    mg["A1"] = "Management equity — simplified vesting @ exit"
    mg["A1"].font = Font(size=14, bold=True)
    _label_cell(mg, "A3", "Time-vest factor (25%/yr × hold)", bold=True)
    _formula_cell(mg, "B3", f"=MIN(1,LBO_Assump!B{row_hold}*0.25)", FMT_PCT)
    _label_cell(mg, "A4", "Perf pool (% if hurdle met — use Python headline; avoids IRR circularity)", bold=True)
    _formula_cell(mg, "B4", f"=LBO_Assump!B{row_perf_pool}", FMT_PCT)
    _label_cell(mg, "A6", "Exit equity (before mgmt carve)", bold=True)
    _formula_cell(mg, "B6", "=Returns!B7", FMT_USD_M)
    _label_cell(mg, "A7", "Option-like payoff (time-vest pool × vest − strike)", bold=True)
    _formula_cell(
        mg,
        "B7",
        f"=MAX(0,B6*LBO_Assump!B{row_m_pool}*B3-LBO_Assump!B{row_m_strike})",
        FMT_USD_M,
    )
    _label_cell(mg, "A8", "Plus: rollover participation (simplified)", bold=True)
    _formula_cell(mg, "B8", f"=LBO_Assump!B{row_m_roll}*MIN(1.5,Returns!B7/(LBO_Assump!B{row_entry_mult}*LBO_Assump!B{row_ltm_ebitda}-LBO_Assump!B{row_exist_nd}))", FMT_USD_M)
    autosize_columns(mg)

    # --- Returns (Excel forecast capped at 5 cols; exit year = min(hold,5)) ---
    ex_col = PROJ_COLS[hold_years_excel - 1]
    ret = wb["Returns"]
    ret["A1"] = "Returns — exit & sponsor economics"
    ret["A1"].font = Font(size=14, bold=True)
    _label_cell(ret, "A3", "Exit EBITDA (exit year)", bold=True)
    _formula_cell(ret, "B3", f"=Operating!{ex_col}5", FMT_USD_M)
    _label_cell(ret, "A4", "Exit Enterprise Value", bold=True)
    _formula_cell(ret, "B4", f"=B3*LBO_Assump!B{row_exit_mult}", FMT_USD_M)
    _label_cell(ret, "A5", "Exit net debt (Debt − Cash)", bold=True)
    _formula_cell(
        ret,
        "B5",
        f"=Debt_Sched!{ex_col}8+Debt_Sched!{ex_col}13+Debt_Sched!{ex_col}19+Debt_Sched!{ex_col}26-BS!{ex_col}4",
        FMT_USD_M,
    )
    _label_cell(ret, "A7", "Exit equity value", bold=True)
    _formula_cell(ret, "B7", "=B4-B5", FMT_USD_M, bold=True)
    ret["B7"].fill = TOTAL_FILL

    _label_cell(ret, "A9", "Less: mgmt waterfall (from Mgmt_Equity)", bold=True)
    _formula_cell(ret, "B9", "=Mgmt_Equity!B7+Mgmt_Equity!B8", FMT_USD_M)

    _label_cell(ret, "A10", "Sponsor proceeds @ exit", bold=True)
    _formula_cell(ret, "B10", "=MAX(0,B7-B9)", FMT_USD_M, bold=True)
    ret["B10"].fill = TOTAL_FILL

    _label_cell(ret, "A11", "Sponsor invested (Sources & Uses)", bold=True)
    _formula_cell(ret, "B11", "=Sources_Uses!D9", FMT_USD_M)

    _label_cell(ret, "A12", "Sponsor MOIC")
    _formula_cell(ret, "B12", "=IF(B11<>0,B10/B11,0)", FMT_MULT, bold=True)
    _label_cell(ret, "A13", "Sponsor IRR (annual)")
    first_cf_row = 40
    irr_cf_last = first_cf_row + hold_years_excel
    ret["B13"] = f"=IRR(B{first_cf_row}:B{irr_cf_last})"
    ret["B13"].number_format = FMT_PCT
    for i in range(hold_years_excel + 1):
        addr = f"B{first_cf_row + i}"
        if i == 0:
            ret[addr] = "=-Sources_Uses!D9"
            ret[addr].number_format = FMT_USD_M
        elif i < hold_years_excel:
            ret[addr] = 0
            ret[addr].number_format = FMT_USD_M
        else:
            ret[addr] = "=B10"
            ret[addr].number_format = FMT_USD_M

    _label_cell(ret, "A14", "Total mgmt economics (approx)", bold=True)
    _formula_cell(ret, "B14", "=B9", FMT_USD_M)

    autosize_columns(ret)

    # --- Bridge ---
    br = wb["Bridge"]
    br["A1"] = "Return bridge (sponsor equity value)"
    br["A1"].font = Font(size=14, bold=True)
    _label_cell(br, "A3", "Entry sponsor equity", bold=True)
    _formula_cell(br, "B3", "=Sources_Uses!D9", FMT_USD_M)
    _label_cell(br, "A4", "Exit sponsor proceeds", bold=True)
    _formula_cell(br, "B4", "=Returns!B10", FMT_USD_M)
    _label_cell(br, "A5", "Δ from EBITDA × entry multiple (growth)", bold=True)
    _formula_cell(br, "B5", f"=(Operating!{ex_col}5-LBO_Assump!B{row_ltm_ebitda})*LBO_Assump!B{row_entry_mult}", FMT_USD_M)
    _label_cell(br, "A6", "Δ from multiple expansion", bold=True)
    _formula_cell(br, "B6", f"=Operating!{ex_col}5*(LBO_Assump!B{row_exit_mult}-LBO_Assump!B{row_entry_mult})", FMT_USD_M)
    _label_cell(br, "A7", "Δ from deleveraging (−Δ net debt to sponsor)", bold=True)
    _formula_cell(br, "B7", f"=-(Returns!B5-(LBO_Assump!B{row_lev}*LBO_Assump!B{row_ltm_ebitda}-Sources_Uses!B7))", FMT_USD_M)
    _label_cell(br, "A8", "Δ from mgmt dilution", bold=True)
    _formula_cell(br, "B8", "=-Returns!B9", FMT_USD_M)
    _label_cell(br, "A9", "Residual / rounding", bold=True)
    _formula_cell(br, "B9", "=B4-B3-B5-B6-B7-B8", FMT_USD_M)
    autosize_columns(br)

    # --- Sensitivity (static, Python) ---
    sen = wb["Sensitivity"]
    sen["A1"] = "Sensitivity tables (static snapshot — rebuild model to refresh)"
    sen["A1"].font = Font(size=12, bold=True)

    def grid_entry_exit():
        em_vals = [entry_mult - 1.5, entry_mult - 0.5, entry_mult, entry_mult + 0.5, entry_mult + 1.5]
        xm_vals = [exit_mult - 1.5, exit_mult - 0.5, exit_mult, exit_mult + 0.5, exit_mult + 1.5]
        sen["A3"] = "Sponsor IRR — Entry mult × Exit mult"
        sen["A4"] = "Exit \\ Entry"
        for j, xm in enumerate(xm_vals):
            sen.cell(row=4, column=2 + j, value=xm).number_format = FMT_MULT
        for i, em in enumerate(em_vals):
            sen.cell(row=5 + i, column=1, value=em).number_format = FMT_MULT
            for j, xm in enumerate(xm_vals):
                out = _python_lbo_headline(
                    hold_years,
                    max(em, 1.0),
                    max(xm, 1.0),
                    ltm_rev,
                    ltm_ebitda_use,
                    revenue_growth_rates,
                    ebitda_margins,
                    debt_to_ebitda_entry,
                    rcf_drawn_at_close,
                    pct_tla_of_term,
                    pct_tlb_of_term,
                    pct_mezz_of_term,
                    tla_rate,
                    _tla_py_mode,
                    _tla_tenor_val,
                    tla_amort_pct_per_year,
                    tlb_rate,
                    tlb_min_amort_pct,
                    tlb_sweep_pct_of_excess,
                    mezz_cash_rate,
                    mezz_pik_rate,
                    rcf_rate,
                    tax_rate,
                    transaction_fees_pct_ev,
                    min_cash_pct_revenue,
                    existing_nd,
                    management_rollover_equity,
                    cogs_pct_revenue,
                    da_pct_revenue,
                    maint_capex_pct_revenue,
                    growth_capex_pct_revenue,
                    dso_days,
                    dio_days,
                    dpo_days,
                    mgmt_option_pool_pct_exit_equity,
                    mgmt_strike_aggregate,
                    perf_pool_pct_if_hurdle_met,
                    sponsor_perf_irr_hurdle,
                )
                cell = sen.cell(row=5 + i, column=2 + j)
                cell.value = out["irr"]
                cell.number_format = FMT_PCT
                if abs(em - entry_mult) < 1e-6 and abs(xm - exit_mult) < 1e-6:
                    cell.fill = TOTAL_FILL

    grid_entry_exit()

    # Grid 2: Revenue growth (applied all yrs) × steady EBITDA margin
    base_r = 12
    sen.cell(row=base_r, column=1, value="Sponsor IRR — Revenue growth × EBITDA margin (static)")
    g_mid = sum(revenue_growth_rates) / len(revenue_growth_rates)
    m_mid = sum(ebitda_margins) / len(ebitda_margins)
    g_vals = [max(-0.1, g_mid + d) for d in [-0.04, -0.02, 0, 0.02, 0.04]]
    m_vals = [max(0.05, min(0.5, m_mid + d)) for d in [-0.04, -0.02, 0, 0.02, 0.04]]
    sen.cell(row=base_r + 1, column=1, value="Margin \\ Growth")
    for j, gv in enumerate(g_vals):
        sen.cell(row=base_r + 1, column=2 + j, value=gv).number_format = FMT_PCT
    for i, mv in enumerate(m_vals):
        sen.cell(row=base_r + 2 + i, column=1, value=mv).number_format = FMT_PCT
        for j, gv in enumerate(g_vals):
            g_list = [gv] * 5
            m_list = [mv] * 5
            out = _python_lbo_headline(
                hold_years,
                entry_mult,
                exit_mult,
                ltm_rev,
                ltm_ebitda_use,
                g_list,
                m_list,
                debt_to_ebitda_entry,
                rcf_drawn_at_close,
                pct_tla_of_term,
                pct_tlb_of_term,
                pct_mezz_of_term,
                tla_rate,
                _tla_py_mode,
                _tla_tenor_val,
                tla_amort_pct_per_year,
                tlb_rate,
                tlb_min_amort_pct,
                tlb_sweep_pct_of_excess,
                mezz_cash_rate,
                mezz_pik_rate,
                rcf_rate,
                tax_rate,
                transaction_fees_pct_ev,
                min_cash_pct_revenue,
                existing_nd,
                management_rollover_equity,
                cogs_pct_revenue,
                da_pct_revenue,
                maint_capex_pct_revenue,
                growth_capex_pct_revenue,
                dso_days,
                dio_days,
                dpo_days,
                mgmt_option_pool_pct_exit_equity,
                mgmt_strike_aggregate,
                perf_pool_pct_if_hurdle_met,
                sponsor_perf_irr_hurdle,
            )
            c = sen.cell(row=base_r + 2 + i, column=2 + j)
            c.value = out["irr"]
            c.number_format = FMT_PCT

    # Grid 3: Leverage × Exit multiple → Sponsor MOIC
    base_r3 = base_r + 8
    sen.cell(row=base_r3, column=1, value="Sponsor MOIC — Leverage (× EBITDA) × Exit multiple (static)")
    lev_mid = debt_to_ebitda_entry
    xm_mid = exit_mult
    lev_vals = [max(2.0, lev_mid + d) for d in [-1.5, -0.5, 0, 0.5, 1.5]]
    xm_vals3 = [max(6.0, xm_mid + d) for d in [-2.0, -1.0, 0, 1.0, 2.0]]
    sen.cell(row=base_r3 + 1, column=1, value="Exit \\ Leverage")
    for j, xv in enumerate(xm_vals3):
        sen.cell(row=base_r3 + 1, column=2 + j, value=xv).number_format = FMT_MULT
    for i, lv in enumerate(lev_vals):
        sen.cell(row=base_r3 + 2 + i, column=1, value=lv).number_format = FMT_MULT
        for j, xv in enumerate(xm_vals3):
            out = _python_lbo_headline(
                hold_years,
                entry_mult,
                xv,
                ltm_rev,
                ltm_ebitda_use,
                revenue_growth_rates,
                ebitda_margins,
                lv,
                rcf_drawn_at_close,
                pct_tla_of_term,
                pct_tlb_of_term,
                pct_mezz_of_term,
                tla_rate,
                _tla_py_mode,
                _tla_tenor_val,
                tla_amort_pct_per_year,
                tlb_rate,
                tlb_min_amort_pct,
                tlb_sweep_pct_of_excess,
                mezz_cash_rate,
                mezz_pik_rate,
                rcf_rate,
                tax_rate,
                transaction_fees_pct_ev,
                min_cash_pct_revenue,
                existing_nd,
                management_rollover_equity,
                cogs_pct_revenue,
                da_pct_revenue,
                maint_capex_pct_revenue,
                growth_capex_pct_revenue,
                dso_days,
                dio_days,
                dpo_days,
                mgmt_option_pool_pct_exit_equity,
                mgmt_strike_aggregate,
                perf_pool_pct_if_hurdle_met,
                sponsor_perf_irr_hurdle,
            )
            c = sen.cell(row=base_r3 + 2 + i, column=2 + j)
            c.value = out["moic"]
            c.number_format = FMT_MULT

    # --- Cover headline ---
    sh_cover["A1"] = f"{co} ({ticker.upper()}) — Leveraged Buyout Model"
    sh_cover["A1"].font = Font(size=16, bold=True, color="1F4E78")
    sh_cover.merge_cells("A1:F1")
    sh_cover["A3"] = "Headline outputs (linked)"
    sh_cover["A3"].font = BOLD_FONT
    sh_cover["A3"].fill = SUBHEADER_FILL
    sh_cover["A4"] = "Sponsor MOIC"
    _formula_cell(sh_cover, "B4", "=Returns!B12", FMT_MULT, bold=True)
    sh_cover["A5"] = "Sponsor IRR"
    _formula_cell(sh_cover, "B5", "=Returns!B13", FMT_PCT, bold=True)
    sh_cover["A6"] = "Mgmt waterfall @ exit"
    _formula_cell(sh_cover, "B6", "=Returns!B14", FMT_USD_M)
    sh_cover["A7"] = "Exit Enterprise Value"
    _formula_cell(sh_cover, "B7", "=Returns!B4", FMT_USD_M)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)

    return json.dumps(
        {
            "status": "ok",
            "path": str(path),
            "ticker": ticker.upper(),
            "python_headline_approx": {
                "sponsor_moic": py_headline["moic"],
                "sponsor_irr": py_headline["irr"],
                "exit_ev_approx": py_headline["exit_ev"],
                "mgmt_waterfall_approx": py_headline["mgmt_waterfall_approx"],
            },
            "notes": (
                "Excel uses BoP interest and non-iterative sweep; enable iterative calc "
                "for average-balance interest extensions. Sponsor IRR uses Excel IRR(). "
                "Perf-vest portion of mgmt equity is in python_headline_approx only (Excel mgmt tab "
                "avoids circular references). Forecast columns cap at 5 years — set hold > 5 for exit timing "
                "still uses column Year min(hold,5). "
                "TLA amort: mode 0 = %% of BoP / yr; mode 1 = fully amortizing fixed debt service via Excel PMT "
                "over the tenor from Sources & Uses TLA."
            ),
        },
        default=str,
        indent=2,
    )
