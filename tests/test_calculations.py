"""Unit tests for core valuation math (no network required)."""

from __future__ import annotations

import math

from tools.dcf import _compute_share_price, _computed_wacc
from tools.lbo import _irr_bisection, _python_lbo_headline


def test_wacc_hamada_reference_value():
    """Regression guard for CAPM + Hamada relevering implementation."""
    w = _computed_wacc(0.04, 0.055, 1.1, 0.25, 0.05, 0.065, 0.21)
    assert abs(w - 0.097873) < 1e-4


def test_dcf_share_price_positive_for_sane_inputs():
    price = _compute_share_price(
        base_rev=1_000_000_000,
        growths=[0.08, 0.07, 0.06, 0.05, 0.04],
        margin=0.20,
        tax=0.21,
        da=0.04,
        capex=0.05,
        nwc=0.02,
        wacc=0.10,
        g_term=0.025,
        net_debt=200_000_000,
        shares=100_000_000,
        terminal_exit_multiple=10.0,
        terminal_method=1,
    )
    assert price > 0


def test_dcf_gordon_requires_wacc_above_terminal_growth():
    price = _compute_share_price(
        base_rev=500_000_000,
        growths=[0.05] * 5,
        margin=0.15,
        tax=0.21,
        da=0.03,
        capex=0.04,
        nwc=0.02,
        wacc=0.02,
        g_term=0.03,
        net_debt=0,
        shares=50_000_000,
        terminal_method=1,
    )
    assert math.isnan(price)


def test_irr_bisection_known_cashflows():
    # Invest 100, receive 121 in one year → 21% IRR
    irr = _irr_bisection([-100.0, 121.0])
    assert abs(irr - 0.21) < 1e-4


def test_lbo_headline_returns_finite_moic():
    out = _python_lbo_headline(
        hold_years=5,
        entry_mult=10.0,
        exit_mult=11.0,
        ltm_rev=1_000_000_000,
        ltm_ebitda=200_000_000,
        growth=[0.05] * 5,
        margins=[0.20] * 5,
        debt_to_ebitda=5.0,
        rcf_draw=50_000_000,
        pct_tla=0.35,
        pct_tlb=0.50,
        pct_mezz=0.15,
        tla_rate=0.07,
        tla_amort_mode="percent_bop",
        tla_amort_tenor=5,
        tla_amort=0.10,
        tlb_rate=0.09,
        tlb_min_amort=0.01,
        tlb_sweep_pct=0.5,
        mezz_cash=0.10,
        mezz_pik=0.04,
        rcf_rate=0.06,
        tax_rate=0.21,
        fees_pct_ev=0.02,
        min_cash_pct_rev=0.02,
        existing_nd=100_000_000,
        rollover=50_000_000,
        cogs_pct=0.60,
        da_pct=0.03,
        capex_maint_pct=0.02,
        capex_grow_pct=0.01,
        dso=45.0,
        dio=30.0,
        dpo=40.0,
        mgmt_pool_pct=0.10,
        mgmt_strike=0.0,
        perf_pool_pct=0.05,
        perf_hurdle=0.20,
    )
    assert out["moic"] > 0
    assert out["irr"] == out["irr"]  # not NaN
