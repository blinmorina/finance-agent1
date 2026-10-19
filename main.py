"""Interactive CLI for the PE Associate agent.

Run:
    python main.py                              # interactive chat
    python main.py --model openai/gpt-4o        # pick provider + model
    python main.py "Build an LBO for MSFT"      # one-shot (positional prompt)

Configure API keys in `.env` (see `.env.example`). Only set the provider(s) you use.
"""

from __future__ import annotations

import argparse
import asyncio
import sys
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / ".env")

from agents import Runner  # noqa: E402
from rich.console import Console  # noqa: E402
from rich.markdown import Markdown  # noqa: E402
from rich.panel import Panel  # noqa: E402

from agent import build_agent  # noqa: E402
from llm_config import get_api_key, missing_provider_help, resolve_llm  # noqa: E402

console = Console()


def _parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="PE Associate — LBO, DCF, comps, SEC diligence (multi-provider LLM).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Model examples:\n"
            "  --model openai/gpt-4o\n"
            "  --model anthropic/claude-opus-5-5\n"
            "  --model google/gemini-2.0-flash\n"
            "  --model moonshot/kimi-k2.6\n"
            "  --model openrouter/anthropic/claude-3.5-sonnet\n"
        ),
    )
    parser.add_argument(
        "--provider",
        choices=["openai", "anthropic", "google", "moonshot", "kimi", "openrouter", "custom"],
        help="LLM provider (overrides FINANCE_AGENT_PROVIDER in .env).",
    )
    parser.add_argument(
        "--model",
        dest="model_spec",
        metavar="PROVIDER/MODEL",
        help="Provider and model id, e.g. openai/gpt-4o or moonshot/kimi-k2.6.",
    )
    parser.add_argument(
        "prompt",
        nargs="*",
        help="Optional one-shot prompt (if omitted, starts interactive chat).",
    )
    return parser.parse_args(argv)


def _check_env(resolved) -> None:
    try:
        get_api_key(resolved)
    except RuntimeError as e:
        console.print(
            Panel.fit(
                f"[bold red]{e}[/bold red]\n\n{missing_provider_help()}",
                title="Setup",
                border_style="red",
            )
        )
        sys.exit(1)


def _banner(resolved) -> None:
    console.print(
        Panel.fit(
            f"[bold cyan]PE Associate[/bold cyan] — LBO modeling, valuation & diligence agent\n"
            f"Model: [yellow]{resolved.display_label}[/yellow]   Outputs: [green]./outputs/[/green]\n\n"
            "Try:\n"
            "  • [italic]Build an LBO for COST using defaults and audit the workbook[/italic]\n"
            "  • [italic]Build a DCF for NVDA and triangulate with trading comps[/italic]\n"
            "  • [italic]Pull risk factors from AAPL's latest 10-K[/italic]\n\n"
            "Commands: [bold]/exit[/bold] quit · [bold]/reset[/bold] clear history · [bold]/files[/bold] list outputs",
            title="Welcome",
            border_style="cyan",
        )
    )


async def _run_once(agent, history: list, user_msg: str) -> list:
    history.append({"role": "user", "content": user_msg})
    with console.status("[bold cyan]Thinking…[/bold cyan]", spinner="dots"):
        result = await Runner.run(agent, history)
    output = result.final_output or ""
    console.print()
    console.print(Panel(Markdown(output), title="PE Associate", border_style="cyan"))
    console.print()
    return result.to_input_list()


async def interactive(agent) -> None:
    history: list = []
    while True:
        try:
            user = console.input("[bold green]you ›[/bold green] ").strip()
        except (EOFError, KeyboardInterrupt):
            console.print("\n[dim]bye[/dim]")
            return
        if not user:
            continue
        if user.lower() in {"/exit", "/quit", ":q"}:
            console.print("[dim]bye[/dim]")
            return
        if user.lower() == "/reset":
            history = []
            console.print("[yellow]conversation history cleared[/yellow]\n")
            continue
        if user.lower() == "/files":
            outputs = Path(__file__).parent / "outputs"
            files = sorted(outputs.glob("*.xlsx"))
            if not files:
                console.print("[dim]no excel files yet in ./outputs[/dim]\n")
            else:
                for f in files:
                    kb = f.stat().st_size / 1024
                    console.print(f"  [green]•[/green] {f.name} [dim]({kb:.1f} KB)[/dim]")
                console.print()
            continue
        try:
            history = await _run_once(agent, history, user)
        except Exception as e:
            console.print(f"[red]error:[/red] {e}\n")


async def one_shot(agent, prompt: str) -> None:
    result = await Runner.run(agent, prompt)
    console.print(Panel(Markdown(result.final_output or ""), title="PE Associate", border_style="cyan"))


def main(argv: list[str] | None = None) -> None:
    argv = argv if argv is not None else sys.argv[1:]
    args = _parse_args(argv)

    provider = args.provider
    if provider == "kimi":
        provider = "moonshot"

    try:
        resolved = resolve_llm(
            provider=provider,
            model_spec=args.model_spec,
        )
    except ValueError as e:
        console.print(f"[red]error:[/red] {e}")
        sys.exit(1)

    _check_env(resolved)
    agent = build_agent(resolved=resolved)

    prompt_text = " ".join(args.prompt).strip()
    if prompt_text:
        asyncio.run(one_shot(agent, prompt_text))
    else:
        _banner(resolved)
        asyncio.run(interactive(agent))


if __name__ == "__main__":
    main()
