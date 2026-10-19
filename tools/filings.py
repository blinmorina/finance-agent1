"""SEC filing text extraction.

Pulls the raw HTML of a 10-K or 10-Q, strips formatting, and extracts the most
analytically valuable sections (MD&A, Risk Factors, Business, Financial
Statements). The text returned is plain text suitable for the agent to read
and summarize.
"""

from __future__ import annotations

import json
import re
import warnings
from typing import Any

import requests
from agents import function_tool
from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

from .market_data import _sec_headers, _ticker_to_cik

SEC_BASE = "https://data.sec.gov"
SEC_WWW = "https://www.sec.gov"

ITEM_PATTERNS: dict[str, list[str]] = {
    # 10-K item map. Different filers use slightly different headings, so each
    # section has multiple candidate regexes.
    "business": [
        r"item\s*1\b\.?\s*business",
    ],
    "risk_factors": [
        r"item\s*1a\b\.?\s*risk\s*factors",
    ],
    "properties": [
        r"item\s*2\b\.?\s*properties",
    ],
    "legal_proceedings": [
        r"item\s*3\b\.?\s*legal\s*proceedings",
    ],
    "mda": [
        r"item\s*7\b\.?\s*management",
        r"item\s*7\b\.?\s*md&a",
    ],
    "market_risk": [
        r"item\s*7a\b\.?\s*quantitative",
    ],
    "financial_statements": [
        r"item\s*8\b\.?\s*financial\s*statements",
    ],
}

# Order matters: when slicing the document, the next "item N." heading marks
# the end of the current section.
ITEM_ORDER = [
    "business",
    "risk_factors",
    "properties",
    "legal_proceedings",
    "mda",
    "market_risk",
    "financial_statements",
]


def _fetch_filing_html(ticker: str, form_type: str = "10-K", year: int | None = None) -> dict[str, Any]:
    """Resolve a ticker to its most recent matching filing and download HTML."""
    cik = _ticker_to_cik(ticker)
    if not cik:
        raise ValueError(f"CIK not found for {ticker}")

    sub_url = f"{SEC_BASE}/submissions/CIK{cik}.json"
    r = requests.get(sub_url, headers=_sec_headers(), timeout=15)
    r.raise_for_status()
    data = r.json()
    recent = data.get("filings", {}).get("recent", {})

    forms = recent.get("form", [])
    dates = recent.get("filingDate", [])
    accessions = recent.get("accessionNumber", [])
    primaries = recent.get("primaryDocument", [])

    target = None
    for form, date, acc, prim in zip(forms, dates, accessions, primaries):
        if form.upper() != form_type.upper():
            continue
        if year is not None and not date.startswith(str(year)):
            continue
        target = {"form": form, "filing_date": date, "accession": acc, "primary": prim}
        break

    if not target:
        raise ValueError(f"No {form_type} filing found for {ticker}" + (f" in {year}" if year else ""))

    acc_nodash = target["accession"].replace("-", "")
    url = f"{SEC_WWW}/Archives/edgar/data/{int(cik)}/{acc_nodash}/{target['primary']}"
    h = requests.get(url, headers={**_sec_headers(), "Accept": "text/html"}, timeout=30)
    h.raise_for_status()
    target["html"] = h.text
    target["filing_url"] = url
    target["company"] = data.get("name")
    return target


def _html_to_text(html: str) -> str:
    """Convert filing HTML to clean plain text."""
    soup = BeautifulSoup(html, "lxml")
    for tag in soup(["script", "style", "table"]):
        tag.decompose()
    text = soup.get_text("\n")
    text = re.sub(r"\xa0", " ", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n\n", text)
    return text.strip()


def _find_section(text: str, section: str) -> tuple[int, int] | None:
    """Locate the start/end byte offsets of a given item section in plain text."""
    if section not in ITEM_PATTERNS:
        return None

    # Find all "Item N." headings to know where THIS section starts and the next begins
    all_items = []  # list of (name, position)
    for name, patterns in ITEM_PATTERNS.items():
        for pat in patterns:
            for m in re.finditer(pat, text, re.IGNORECASE):
                all_items.append((name, m.start()))
    all_items.sort(key=lambda x: x[1])

    # Discard the table-of-contents block: a heading is "real" only on its
    # second or later occurrence (TOC mentions every Item before they appear).
    seen: dict[str, int] = {}
    real_positions: list[tuple[str, int]] = []
    for name, pos in all_items:
        seen[name] = seen.get(name, 0) + 1
        if seen[name] >= 2:
            real_positions.append((name, pos))

    start_pos = None
    for name, pos in real_positions:
        if name == section:
            start_pos = pos
            break
    if start_pos is None:
        return None

    # Find the first "real" item AFTER start_pos to mark the end
    end_pos = len(text)
    for name, pos in real_positions:
        if pos > start_pos:
            end_pos = pos
            break
    return (start_pos, end_pos)


@function_tool
def get_filing_section(
    ticker: str,
    section: str = "mda",
    form_type: str = "10-K",
    year: int | None = None,
    max_chars: int = 20000,
) -> str:
    """Extract a named section from a company's SEC filing as plain text.

    Args:
        ticker: Stock symbol.
        section: One of "business", "risk_factors", "properties",
            "legal_proceedings", "mda", "market_risk", "financial_statements".
        form_type: "10-K" or "10-Q".
        year: Optional 4-digit year to target a specific filing.
        max_chars: Cap on returned text length (the section is truncated with
            a notice if it exceeds this).

    Returns JSON with the section text, filing metadata, and source URL.
    Use this for qualitative analysis: management commentary, segment trends,
    risk disclosures, accounting policies.
    """
    try:
        filing = _fetch_filing_html(ticker, form_type=form_type, year=year)
        text = _html_to_text(filing["html"])

        span = _find_section(text, section)
        if span is None:
            return json.dumps(
                {
                    "error": f"section '{section}' not found in {form_type}",
                    "available_sections": list(ITEM_PATTERNS.keys()),
                    "filing_url": filing["filing_url"],
                }
            )

        start, end = span
        body = text[start:end].strip()
        truncated = False
        if len(body) > max_chars:
            body = body[:max_chars] + f"\n\n... [truncated, {len(text[start:end]) - max_chars:,} more chars in full section]"
            truncated = True

        return json.dumps(
            {
                "ticker": ticker.upper(),
                "company": filing["company"],
                "form": filing["form"],
                "filing_date": filing["filing_date"],
                "section": section,
                "char_count": len(body),
                "truncated": truncated,
                "filing_url": filing["filing_url"],
                "text": body,
            },
            indent=2,
        )
    except Exception as e:
        return json.dumps({"error": str(e), "ticker": ticker, "section": section})


@function_tool
def list_filing_sections(ticker: str, form_type: str = "10-K", year: int | None = None) -> str:
    """List which sections were successfully detected in a filing, with sizes.

    Use this before calling get_filing_section to see what's available and
    roughly how large each section is.
    """
    try:
        filing = _fetch_filing_html(ticker, form_type=form_type, year=year)
        text = _html_to_text(filing["html"])
        result = {}
        for section in ITEM_ORDER:
            span = _find_section(text, section)
            if span:
                start, end = span
                result[section] = {"char_count": end - start, "found": True}
            else:
                result[section] = {"found": False}
        return json.dumps(
            {
                "ticker": ticker.upper(),
                "company": filing["company"],
                "form": filing["form"],
                "filing_date": filing["filing_date"],
                "filing_url": filing["filing_url"],
                "total_chars": len(text),
                "sections": result,
            },
            indent=2,
        )
    except Exception as e:
        return json.dumps({"error": str(e), "ticker": ticker})
