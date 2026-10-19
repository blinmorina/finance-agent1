"""Structured analyst report output.

Produces a machine-readable JSON record alongside the narrative the agent
writes in chat. Every numeric field carries a `source_tag` so reviewers (or
downstream evals) can verify where each number came from:

  - "reported"   — number lifted directly from a source (e.g. 10-K, market data)
  - "calculated" — derived deterministically from reported numbers
  - "estimated"  — model output (DCF, ratio extrapolation)
  - "assumption" — user-provided or default input

This makes the agent's output verifiable instead of just persuasive.
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any, Literal

from agents import function_tool

from .excel import OUTPUTS_DIR

SourceTag = Literal["reported", "calculated", "estimated", "assumption"]


@function_tool(strict_mode=False)
def save_analyst_report(
    ticker: str,
    company_name: str,
    recommendation: str,
    thesis: str,
    target_price: float | None = None,
    current_price: float | None = None,
    valuation_method: str = "DCF",
    key_assumptions: dict[str, Any] | None = None,
    key_metrics: dict[str, Any] | None = None,
    risks: list[str] | None = None,
    catalysts: list[str] | None = None,
    sources: list[str] | None = None,
    confidence: str = "medium",
    output_file: str | None = None,
) -> str:
    """Persist a structured analyst report as JSON. Useful for evals and audit.

    Args:
        ticker: Stock symbol.
        company_name: Full company name.
        recommendation: One of "buy", "hold", "sell", "watch", "no-rating".
        thesis: 2-4 sentence investment thesis.
        target_price: Implied / target share price.
        current_price: Current market price (for upside calc).
        valuation_method: e.g. "DCF", "Comps", "DCF + Comps", "Sum-of-the-parts".
        key_assumptions: Dict of {assumption_name: value}, e.g. {"wacc": 0.09}.
        key_metrics: Dict of {metric_name: {value, source_tag, period}}, where
            source_tag is one of reported/calculated/estimated/assumption.
            Example:
              {"revenue_ttm": {"value": 391035000000, "source_tag": "reported", "period": "FY2024", "source": "10-K"}}
        risks: List of key downside risks.
        catalysts: List of key upside catalysts.
        sources: List of source citations (URLs or "10-K FY2024", "yfinance",
            "SEC EDGAR XBRL", etc.).
        confidence: Analyst confidence: "low" / "medium" / "high".
        output_file: Output filename. Defaults to "{TICKER}_report_{date}.json".

    Returns JSON with the saved path and a brief echo of the report.
    """
    if recommendation.lower() not in {"buy", "hold", "sell", "watch", "no-rating"}:
        return json.dumps({"error": f"recommendation must be buy/hold/sell/watch/no-rating, got '{recommendation}'"})
    if confidence.lower() not in {"low", "medium", "high"}:
        return json.dumps({"error": f"confidence must be low/medium/high, got '{confidence}'"})

    upside_pct = None
    if target_price and current_price:
        upside_pct = (target_price / current_price - 1) * 100

    report = {
        "schema_version": "1.0",
        "generated_at": datetime.utcnow().isoformat() + "Z",
        "ticker": ticker.upper(),
        "company_name": company_name,
        "recommendation": recommendation.lower(),
        "confidence": confidence.lower(),
        "thesis": thesis,
        "valuation": {
            "method": valuation_method,
            "target_price": target_price,
            "current_price": current_price,
            "upside_pct": upside_pct,
        },
        "key_assumptions": key_assumptions or {},
        "key_metrics": key_metrics or {},
        "risks": risks or [],
        "catalysts": catalysts or [],
        "sources": sources or [],
    }

    fname = output_file or f"{ticker.upper()}_report_{datetime.utcnow().strftime('%Y%m%d')}.json"
    if not fname.endswith(".json"):
        fname += ".json"
    path = Path(fname)
    if not path.is_absolute():
        path = OUTPUTS_DIR / path
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w") as f:
        json.dump(report, f, indent=2, default=str)

    quality_warnings = []
    if not sources:
        quality_warnings.append("no sources cited — every reported number should have a source")
    if key_metrics:
        untagged = [k for k, v in key_metrics.items() if not isinstance(v, dict) or "source_tag" not in v]
        if untagged:
            quality_warnings.append(f"these metrics are missing source_tag: {untagged}")
    if not risks:
        quality_warnings.append("no risks listed — every recommendation should disclose downside")
    if recommendation.lower() in {"buy", "sell"} and target_price is None:
        quality_warnings.append("buy/sell rating without an explicit target price")

    return json.dumps(
        {
            "status": "ok",
            "path": str(path),
            "report_summary": {
                "ticker": report["ticker"],
                "recommendation": report["recommendation"],
                "target_price": target_price,
                "upside_pct": upside_pct,
                "confidence": confidence,
                "num_metrics": len(report["key_metrics"]),
                "num_risks": len(report["risks"]),
                "num_sources": len(report["sources"]),
            },
            "quality_warnings": quality_warnings,
        },
        default=str,
        indent=2,
    )


@function_tool
def read_analyst_report(file_path: str) -> str:
    """Read back a previously saved analyst report JSON."""
    path = Path(file_path).expanduser()
    if not path.is_absolute():
        path = OUTPUTS_DIR / path
    if not path.exists():
        return json.dumps({"error": f"file not found: {path}"})
    with open(path) as f:
        return f.read()
