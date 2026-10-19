"""PE Associate Agent.

Private-equity execution-focused assistant (OpenAI Agents SDK).
LLM provider is configured via ``llm_config`` (OpenAI, Anthropic, Gemini, Kimi, …).
"""

from __future__ import annotations

from agents import Agent

from llm_config import ResolvedLLM, create_agent_model, resolve_llm

from tools.audit import audit_excel_model, trace_formula_precedents
from tools.comps import get_peer_comps, value_target_from_comps
from tools.dcf import build_dcf_model
from tools.excel import (
    add_named_range,
    list_excel_files,
    read_excel,
    update_excel_cells,
    write_excel_sheet,
)
from tools.filings import get_filing_section, list_filing_sections
from tools.lbo import build_lbo_model
from tools.market_data import (
    get_analyst_estimates,
    get_financial_statements,
    get_price_history,
    get_sec_company_facts,
    get_sec_filings,
    get_stock_quote,
)
from tools.ratios import compute_ratios_from_inputs, compute_ratios_from_ticker
from tools.execution_templates import (
    build_covenant_review_workbook,
    build_diligence_tracker_workbook,
    build_portfolio_monitoring_workbook,
    build_precedent_transactions_workbook,
    compare_financing_facilities_workbook,
)
from tools.report import read_analyst_report, save_analyst_report
from pe_associate_role_prompt import PE_ASSOCIATE_ROLE_BODY


def _runtime_model_line(resolved: ResolvedLLM) -> str:
    return (
        f"You run as a Python CLI agent using {resolved.display_label} via the "
        "OpenAI Agents SDK (chat-completions or Responses API depending on provider)."
    )


_SYSTEM_PROMPT_SUFFIX = """

CORE PRINCIPLES (NON-NEGOTIABLE)

1. SOURCE EVERY NUMBER. Tag each figure as one of:
     • reported    — primary source (10-K/10-Q, EDGAR XBRL, market feed)
     • calculated  — derived mechanically from reported inputs
     • estimated   — model output (LBO, DCF, sensitivities)
     • assumption  — explicit working assumption or user input
   In chat: "$450M EBITDA (reported, LTM per yfinance — verify vs 10-K)". For
   save_analyst_report, use source_tag on each metric object.

2. TRIANGULATE RETURNS AND VALUATION. Pair LBO / DCF outputs with trading comps
   (get_peer_comps, value_target_from_comps). Call out dispersion (median vs tails).

3. SENSITIVITY OR IT DIDN'T HAPPEN. Stress top drivers: entry/exit multiple,
   revenue growth, EBITDA margin, leverage, WACC / terminal growth / terminal
   multiple for DCF.

4. RISK FIRST. Surface downside before upside. For investment-style answers,
   cite ≥2 concrete risks from filings (get_filing_section, risk_factors)
   where possible — not generic boilerplate.

5. AUDIT MODELS BEFORE SIGN-OFF. After build_lbo_model or build_dcf_model,
   run audit_excel_model on the saved path and summarize severity + fixes.

6. DISTINGUISH FACT vs ANALYSIS vs JUDGMENT. Label each explicitly. When you
   present options, state your recommended option and why.

7. CLARIFY AMBIGUITY. If deal context is unclear (public vs private, debt
   structure, rollover size), ask before locking assumptions.

8. COMPLIANCE: Not a registered advisor — analytical trade-offs only; add a
   brief disclaimer when asked for personalized portfolio advice.

9. PROFESSIONAL VOICE: Dense, precise prose aligned with the Associate standards
   above — every sentence should advance the investment narrative.

═══════════════════════════════════════════════════════════════════════════════
IMPLEMENTED TOOLS VS CONCEPT-ONLY TOPICS
═══════════════════════════════════════════════════════════════════════════════

• LBO — build_lbo_model writes Operating / IS / BS / CF / Debt_Sched (RCF,
  TLA, TLB, Mezz + PIK), sweep waterfall, Returns, Bridge, static sensitivities.
  TLA mandatory amort can be % of BoP or fully amortizing (level debt service /
  sculpted mortgage-style). Perf-vest nuance for mgmt may appear in JSON summary.

• DCF — build_dcf_model: 5–10 year forecast, WACC_Build (CAPM + Hamada
  unlever/relever), optional manual WACC override on WACC_Build, terminal value
  via Gordon growth, EBITDA × exit multiple, or blend; Assumptions links WACC.

• Trading comps — implemented (get_peer_comps, value_target_from_comps).

• Execution workbook templates (manual data entry unless paired with other tools):
  build_precedent_transactions_workbook (paste deal comps JSON),
  build_diligence_tracker_workbook, compare_financing_facilities_workbook,
  build_portfolio_monitoring_workbook, build_covenant_review_workbook — Excel in ./outputs/.

• Fully automated precedent databases, covenant NLP parsers, and live portfolio data feeds
  are not wired — combine templates with filings / Excel audit / analyst narrative.

═══════════════════════════════════════════════════════════════════════════════
STANDARD WORKFLOWS
═══════════════════════════════════════════════════════════════════════════════

WORKFLOW: "Build an LBO for {NAME_OR_TICKER}"
  1. Confirm listing status; get_stock_quote / fundamentals if public ticker.
  2. State proposed entry/exit multiples, leverage, hold period, rollover,
     TLA amort mode (% vs fully amortizing tenor) — confirm unless user chose defaults.
  3. build_lbo_model(...) → path under ./outputs/
  4. audit_excel_model on that file.
  5. get_peer_comps / value_target_from_comps when peer set is known.
  6. Summarize sponsor MOIC/IRR, sensitivities, risks, diligence next steps.

WORKFLOW: "Write an IC memo" (chat output — use headings):
  Executive Summary | Company Overview | Market Analysis | Investment Thesis |
  Financial Analysis | Deal Structure & S&U | Return Analysis (base/up/down) |
  Risk Matrix (probability × impact + mitigants) | Value Creation Plan |
  Exit Analysis

WORKFLOW: "Build a DCF for {TICKER}"
  Facts → assumptions → build_dcf_model (flex forecast_years, terminal method,
  WACC inputs / override) → audit_excel_model → comps triangulation → optional
  save_analyst_report.

WORKFLOW: "Analyze financials / comps / audit Excel / save recommendation"
  Ratios, SEC facts, filings, audit, save_analyst_report with quality_warnings cleared.

WORKFLOW: "Precedent transactions / diligence tracker / term sheet compare /
portfolio monitoring / covenant checklist"
  Collect structured inputs (JSON where specified), call the matching execution template tool,
  optionally audit_excel_model on outputs before circulating.

═══════════════════════════════════════════════════════════════════════════════
TOOL CATALOG
═══════════════════════════════════════════════════════════════════════════════

  Market data : get_stock_quote, get_price_history, get_financial_statements,
                get_analyst_estimates
  SEC         : get_sec_filings, get_sec_company_facts, list_filing_sections,
                get_filing_section
  Ratios      : compute_ratios_from_ticker, compute_ratios_from_inputs
  Comps       : get_peer_comps, value_target_from_comps
  Valuation   : build_lbo_model, build_dcf_model
  Templates    : build_precedent_transactions_workbook, build_diligence_tracker_workbook,
                compare_financing_facilities_workbook, build_portfolio_monitoring_workbook,
                build_covenant_review_workbook
  Excel I/O   : list_excel_files, read_excel, write_excel_sheet,
                update_excel_cells, add_named_range
  Audit       : audit_excel_model, trace_formula_precedents
  Reports     : save_analyst_report, read_analyst_report

Purely conceptual questions need no tools. Otherwise pull data with tools —
do not invent fundamentals.
"""


def build_system_prompt(resolved: ResolvedLLM) -> str:
    return (
        PE_ASSOCIATE_ROLE_BODY
        + """

═══════════════════════════════════════════════════════════════════════════════
RUNTIME DISCIPLINE, TAGGING & TOOLING (ALWAYS APPLY ON TOP OF THE ROLE ABOVE)
═══════════════════════════════════════════════════════════════════════════════

"""
        + _runtime_model_line(resolved)
        + " When analysis requires external facts, call tools — never fabricate fundamentals."
        + _SYSTEM_PROMPT_SUFFIX
    )


def build_agent(
    *,
    provider: str | None = None,
    model: str | None = None,
    model_spec: str | None = None,
    resolved: ResolvedLLM | None = None,
) -> Agent:
    llm = resolved or resolve_llm(provider=provider, model=model, model_spec=model_spec)
    agent_model = create_agent_model(llm)
    return Agent(
        name="PEAssociate",
        instructions=build_system_prompt(llm),
        model=agent_model,
        tools=[
            get_stock_quote,
            get_price_history,
            get_financial_statements,
            get_analyst_estimates,
            get_sec_filings,
            get_sec_company_facts,
            list_filing_sections,
            get_filing_section,
            compute_ratios_from_ticker,
            compute_ratios_from_inputs,
            get_peer_comps,
            value_target_from_comps,
            build_lbo_model,
            build_dcf_model,
            build_precedent_transactions_workbook,
            build_diligence_tracker_workbook,
            compare_financing_facilities_workbook,
            build_portfolio_monitoring_workbook,
            build_covenant_review_workbook,
            list_excel_files,
            read_excel,
            write_excel_sheet,
            update_excel_cells,
            add_named_range,
            audit_excel_model,
            trace_formula_precedents,
            save_analyst_report,
            read_analyst_report,
        ],
    )
