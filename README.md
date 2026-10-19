# PE Associate — Finance Agent

An open-source **private equity associate** agent: LBO and DCF Excel models, trading comps, SEC filings, ratio analysis, model audit, and structured analyst reports. It uses the [OpenAI Agents SDK](https://github.com/openai/openai-agents-python) with **pluggable LLMs** — run it on **OpenAI (ChatGPT), Anthropic (Claude), Google (Gemini), Moonshot (Kimi), OpenRouter**, or any **OpenAI-compatible** endpoint.

Inspired by the tool-first discipline of [vals-ai/finance-agent-v2](https://github.com/vals-ai/finance-agent-v2), but focused on **modeling and diligence workflows** rather than the Vals benchmark harness.

## What the agent can do

The agent calls Python tools (no hallucinated fundamentals). Main capabilities:

| Area | Tools |
| --- | --- |
| **Market data** | `get_stock_quote`, `get_price_history`, `get_financial_statements`, `get_analyst_estimates` (yfinance) |
| **SEC / EDGAR** | `get_sec_filings`, `get_sec_company_facts`, `list_filing_sections`, `get_filing_section` (free EDGAR APIs) |
| **Ratios & comps** | `compute_ratios_from_ticker`, `compute_ratios_from_inputs`, `get_peer_comps`, `value_target_from_comps` |
| **Valuation models** | `build_lbo_model`, `build_dcf_model` → `.xlsx` under `./outputs/` |
| **Deal templates** | Precedent TX, diligence tracker, facility compare, portfolio monitoring, covenant checklist workbooks |
| **Excel I/O & audit** | `read_excel`, `write_excel_sheet`, `audit_excel_model`, `trace_formula_precedents` |
| **Reports** | `save_analyst_report`, `read_analyst_report` (JSON with `source_tag` on metrics) |

Generated workbooks use **Excel formulas**; open them in Excel or LibreOffice to recalculate. The tools also return **Python-computed headlines** in JSON for quick review without opening Excel.

---

## Set up

### Dependencies

**Option A — Make (recommended)**

```bash
git clone https://github.com/<your-username>/finance-agent.git
cd finance-agent
make install
source .venv/bin/activate
```

**Option B — pip**

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Environment variables

Copy the example env file and add **only the API keys you use**:

```bash
cp .env.example .env
```

```bash
# ── Provider selection (optional) ──
FINANCE_AGENT_PROVIDER=openai          # openai | anthropic | google | moonshot | openrouter | custom
FINANCE_AGENT_MODEL=gpt-4o             # model id for that provider

# ── LLM keys (set one or more; first match can auto-select provider) ──
OPENAI_API_KEY=<openai_api_key>
ANTHROPIC_API_KEY=<anthropic_api_key>
GOOGLE_API_KEY=<google_api_key>
MOONSHOT_API_KEY=<moonshot_api_key>
OPENROUTER_API_KEY=<openrouter_api_key>

# Custom OpenAI-compatible server (LM Studio, vLLM, etc.)
# CUSTOM_BASE_URL=http://localhost:1234/v1
# CUSTOM_API_KEY=<optional>

# ── SEC (required for EDGAR; use your real email) ──
SEC_USER_AGENT=Finance Agent your-email@example.com
```

`.env` takes precedence over variables already exported in your shell.

| Provider | API key | Default model | Notes |
| --- | --- | --- | --- |
| **OpenAI** | `OPENAI_API_KEY` | `gpt-4o` | Uses Responses API by default; set `OPENAI_API_MODE=chat_completions` to force chat API |
| **Anthropic** | `ANTHROPIC_API_KEY` | `claude-opus-5-5` | [OpenAI-compatible endpoint](https://docs.anthropic.com/en/api/openai-sdk) |
| **Google** | `GOOGLE_API_KEY` | `gemini-2.0-flash` | [Gemini OpenAI compatibility](https://ai.google.dev/gemini-api/docs/openai) |
| **Moonshot (Kimi)** | `MOONSHOT_API_KEY` | `kimi-k2.6` | [Moonshot platform](https://platform.moonshot.ai/) |
| **OpenRouter** | `OPENROUTER_API_KEY` | `openai/gpt-4o` | One key for many upstream models — [openrouter.ai](https://openrouter.ai/) |
| **Custom** | `CUSTOM_API_KEY` | (you set) | Requires `CUSTOM_BASE_URL` |

Override base URLs if needed: `MOONSHOT_BASE_URL`, `ANTHROPIC_BASE_URL`, `GOOGLE_BASE_URL`, etc. (see `.env.example`).

---

## Running the agent

### Interactive chat

```bash
python main.py
```

### One-shot prompt

```bash
python main.py "Build an LBO for MSFT using defaults, audit the workbook, and summarize sponsor returns"
```

### Choose model on the command line

Use `provider/model`, similar to model ids in other agent harnesses:

```bash
python main.py --model openai/gpt-4o
python main.py --model anthropic/claude-opus-5-5 "Compute ratios for MSFT vs AAPL and GOOGL"
python main.py --model google/gemini-2.0-flash
python main.py --model moonshot/kimi-k2.6
python main.py --model openrouter/anthropic/claude-3.5-sonnet
```

Or set provider explicitly:

```bash
python main.py --provider openai --model gpt-4o-mini "Build a DCF for NVDA"
```

CLI flags override `FINANCE_AGENT_PROVIDER` / `FINANCE_AGENT_MODEL` in `.env`.

### Example prompts

```
Build an LBO for COST using defaults, then audit outputs/COST_LBO.xlsx
Build a DCF for NVDA and triangulate with trading comps
Pull Item 1A risk factors from AAPL's latest 10-K and tie to a short investment view
Draft an IC memo outline for a take-private of TICKER with bull/base/bear returns
```

Chat commands: `/exit`, `/reset`, `/files` (list `./outputs/*.xlsx`).

### CLI help

```bash
python main.py --help
```

---

## Tests

Unit tests cover valuation math and provider parsing (no network):

```bash
make install-dev
make test
# or: pytest tests/ -q
```

---

## Project layout

```
finance-agent/
├── main.py                 # CLI (--provider, --model)
├── agent.py                # Agent + tools + system prompt
├── llm_config.py           # Multi-provider LLM wiring
├── pe_associate_role_prompt.py
├── tools/                  # Market data, SEC, LBO, DCF, comps, audit, …
├── outputs/                # Generated .xlsx / reports
├── tests/
├── Makefile
└── .env.example
```

---

## Analyst guardrails (built into the system prompt)

- Tag numbers: `reported` / `calculated` / `estimated` / `assumption`
- Prefer **SEC XBRL and filing text** over secondary summaries
- **Triangulate** DCF/LBO with trading comps; run **sensitivity** on key drivers
- **Audit** models after `build_lbo_model` / `build_dcf_model`
- Cite **specific risks** from filings when giving investment-style views

---

## Comparison with vals-ai/finance-agent-v2

| | **vals-ai/finance-agent-v2** | **This repo** |
| --- | --- | --- |
| Purpose | FAB v2 benchmark (Q&A evals) | PE modeling & diligence assistant |
| LLM | Vals model library | OpenAI, Anthropic, Gemini, Kimi, OpenRouter, custom |
| Tools | Tavily, SEC search API, Tiingo prices, HTML store/retrieve | yfinance, free EDGAR, Excel LBO/DCF, filing sections |
| Output | Graded answers + logs | Chat + Excel workbooks + optional JSON reports |

---

## Notes & limitations

- Not investment advice. Defaults from yfinance are **starting points** — verify against 10-K/XBRL.
- LBO Excel vs Python headline can differ slightly (interest timing, mgmt perf-vest circularity); see tool JSON `notes`.
- Precedent M&A databases and covenant NLP are **not** automated — use execution templates + your data.

## License

MIT — see [LICENSE](LICENSE).
