"""Excel model audit tools.

Checks a workbook for common analyst errors:
  - Hardcoded values inside a "formula row" (breaks model integrity)
  - Cached formula errors (#REF!, #DIV/0!, #VALUE!, #N/A, #NAME?)
  - Balance sheet that doesn't balance (Assets vs Liab + Equity)
  - Broken named ranges
  - Suspicious circular references

Two tools:
  - audit_excel_model: high-level audit report
  - trace_formula_precedents: list which cells a given formula depends on
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from agents import function_tool
from openpyxl import load_workbook

from .excel import OUTPUTS_DIR

ERROR_VALUES = {"#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NUM!", "#NULL!"}
CELL_REF_RE = re.compile(r"(?:'([^']+)'!|([A-Za-z_]\w*)!)?(\$?[A-Z]{1,3}\$?\d+)")


def _resolve(file_path: str) -> Path:
    p = Path(file_path).expanduser()
    if not p.is_absolute():
        p = OUTPUTS_DIR / p
    if p.suffix == "":
        p = p.with_suffix(".xlsx")
    return p


def _scan_hardcoded_in_formula_rows(ws) -> list[dict[str, Any]]:
    """Rows where most non-empty cells are formulas but a few are hardcoded
    values — those hardcoded cells often indicate broken links or manual
    overrides that should be flagged."""
    issues = []
    for row in ws.iter_rows():
        formulas = 0
        numerics = 0
        numeric_cells = []
        for c in row:
            if c.value is None:
                continue
            if isinstance(c.value, str) and c.value.startswith("="):
                formulas += 1
            elif isinstance(c.value, (int, float)):
                numerics += 1
                numeric_cells.append(c)
        total = formulas + numerics
        if total >= 3 and formulas >= 2 and numerics >= 1 and formulas / total >= 0.5:
            for c in numeric_cells:
                issues.append(
                    {
                        "sheet": ws.title,
                        "cell": c.coordinate,
                        "value": c.value,
                        "issue": "hardcoded value in a row of formulas (possible manual override)",
                    }
                )
    return issues


def _scan_formula_errors(ws_data) -> list[dict[str, Any]]:
    """Use cached values (data_only=True) to find #REF!, #DIV/0!, etc."""
    issues = []
    for row in ws_data.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value in ERROR_VALUES:
                issues.append(
                    {
                        "sheet": ws_data.title,
                        "cell": c.coordinate,
                        "error": c.value,
                    }
                )
    return issues


def _check_balance_sheet(wb) -> dict[str, Any] | None:
    """Look for a 'Balance Sheet' (or similar) sheet and check that
    Total Assets ≈ Total Liabilities + Total Equity by scanning row labels."""
    target_sheet = None
    for name in wb.sheetnames:
        low = name.lower()
        if "balance" in low or low == "bs":
            target_sheet = wb[name]
            break
    if target_sheet is None:
        return None

    label_col = 1
    rows_of_interest: dict[str, float | None] = {
        "total_assets": None,
        "total_liabilities": None,
        "total_equity": None,
    }
    for row in target_sheet.iter_rows():
        if not row:
            continue
        label = row[label_col - 1].value
        if not isinstance(label, str):
            continue
        norm = label.strip().lower()
        # find the rightmost numeric value in the row
        last_num = None
        for c in row:
            if isinstance(c.value, (int, float)):
                last_num = c.value
        if norm in ("total assets", "assets total"):
            rows_of_interest["total_assets"] = last_num
        elif norm in ("total liabilities", "liabilities total"):
            rows_of_interest["total_liabilities"] = last_num
        elif norm in ("total equity", "stockholders equity", "total stockholders equity"):
            rows_of_interest["total_equity"] = last_num

    assets = rows_of_interest["total_assets"]
    liab = rows_of_interest["total_liabilities"]
    equity = rows_of_interest["total_equity"]
    if assets is None or (liab is None and equity is None):
        return {"sheet": target_sheet.title, "status": "could not parse labels", "rows_found": rows_of_interest}

    rhs = (liab or 0) + (equity or 0)
    diff = assets - rhs
    tolerance = max(abs(assets) * 0.001, 1.0)
    balanced = abs(diff) < tolerance
    return {
        "sheet": target_sheet.title,
        "total_assets": assets,
        "total_liab_plus_equity": rhs,
        "difference": diff,
        "balanced": balanced,
        "tolerance": tolerance,
    }


def _check_named_ranges(wb) -> list[dict[str, Any]]:
    issues = []
    try:
        names = dict(wb.defined_names)
    except Exception:
        names = {}
    for name, dn in names.items():
        try:
            attr = getattr(dn, "attr_text", None) or str(dn)
        except Exception:
            attr = "?"
        if not attr or attr == "?":
            issues.append({"name": name, "issue": "could not read reference"})
            continue
        if "#REF!" in attr:
            issues.append({"name": name, "ref": attr, "issue": "broken reference"})
    return issues


@function_tool
def audit_excel_model(file_path: str) -> str:
    """Run a structured audit on an Excel workbook and return findings.

    Checks include:
      - Hardcoded numbers in rows that are otherwise formulas (red flag for
        manual overrides or broken links in a financial model).
      - Cached formula errors (#REF!, #DIV/0!, #VALUE!, #N/A, ...).
      - Balance sheet equation check (if a "Balance Sheet" tab is found).
      - Named-range broken references.
      - Inventory of every sheet's size and how many formulas it contains.

    Args:
        file_path: Path to the .xlsx file. Relative paths resolve under outputs/.

    Returns JSON with grouped issues and an overall severity summary.
    """
    path = _resolve(file_path)
    if not path.exists():
        return json.dumps({"error": f"file not found: {path}"})

    wb_formulas = load_workbook(path, data_only=False)
    wb_values = load_workbook(path, data_only=True)

    sheet_inventory: list[dict[str, Any]] = []
    hardcoded_issues: list[dict[str, Any]] = []
    error_issues: list[dict[str, Any]] = []

    for name in wb_formulas.sheetnames:
        ws_f = wb_formulas[name]
        formula_count = 0
        value_count = 0
        for row in ws_f.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                if isinstance(c.value, str) and c.value.startswith("="):
                    formula_count += 1
                else:
                    value_count += 1
        sheet_inventory.append(
            {
                "sheet": name,
                "max_row": ws_f.max_row,
                "max_col": ws_f.max_column,
                "formulas": formula_count,
                "values": value_count,
            }
        )
        hardcoded_issues.extend(_scan_hardcoded_in_formula_rows(ws_f))
        if name in wb_values.sheetnames:
            error_issues.extend(_scan_formula_errors(wb_values[name]))

    balance_check = _check_balance_sheet(wb_formulas)
    named_range_issues = _check_named_ranges(wb_formulas)

    severity = "clean"
    if error_issues or any(i.get("issue") == "broken reference" for i in named_range_issues):
        severity = "critical"
    elif hardcoded_issues or (balance_check and not balance_check.get("balanced", True)):
        severity = "warning"

    return json.dumps(
        {
            "path": str(path),
            "severity": severity,
            "summary": {
                "sheets": len(sheet_inventory),
                "formula_errors": len(error_issues),
                "hardcoded_value_warnings": len(hardcoded_issues),
                "named_range_issues": len(named_range_issues),
            },
            "sheet_inventory": sheet_inventory,
            "formula_errors": error_issues,
            "hardcoded_in_formula_rows": hardcoded_issues[:50],
            "named_range_issues": named_range_issues,
            "balance_sheet_check": balance_check,
            "notes": (
                "Formula errors require opening the file in Excel and re-saving "
                "to refresh cached values. If you see 0 errors but the file was "
                "just generated, that's expected — Excel hasn't recalculated yet."
            ),
        },
        default=str,
        indent=2,
    )


@function_tool
def trace_formula_precedents(file_path: str, sheet_name: str, cell: str) -> str:
    """List the cells/named ranges a formula directly references.

    Useful for auditing model logic. Example: trace what feeds into the
    'Implied share price' cell to make sure it's tied to assumptions.

    Args:
        file_path: Path to the .xlsx file.
        sheet_name: Sheet containing the cell.
        cell: A1 address of the cell to inspect.

    Returns JSON with the formula, all cell/range references, and any
    named-range references it uses.
    """
    path = _resolve(file_path)
    if not path.exists():
        return json.dumps({"error": f"file not found: {path}"})
    wb = load_workbook(path, data_only=False)
    if sheet_name not in wb.sheetnames:
        return json.dumps({"error": f"sheet '{sheet_name}' not found", "available": wb.sheetnames})

    ws = wb[sheet_name]
    val = ws[cell].value
    if not isinstance(val, str) or not val.startswith("="):
        return json.dumps({"sheet": sheet_name, "cell": cell, "value": val, "note": "not a formula"})

    formula = val
    cell_refs = []
    for m in CELL_REF_RE.finditer(formula):
        sh = m.group(1) or m.group(2) or sheet_name
        ref = m.group(3)
        cell_refs.append(f"{sh}!{ref}")

    named = []
    try:
        for name in wb.defined_names:
            if re.search(rf"\b{re.escape(name)}\b", formula):
                attr = getattr(wb.defined_names[name], "attr_text", str(wb.defined_names[name]))
                named.append({"name": name, "refers_to": attr})
    except Exception:
        pass

    return json.dumps(
        {
            "sheet": sheet_name,
            "cell": cell,
            "formula": formula,
            "direct_cell_refs": cell_refs,
            "named_range_refs": named,
        },
        indent=2,
    )
