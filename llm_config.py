"""Multi-provider LLM configuration for the OpenAI Agents SDK.

The SDK talks to providers through an OpenAI-compatible HTTP API. Native
OpenAI can use the Responses API; other providers use chat-completions.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Literal

from agents import (
    OpenAIChatCompletionsModel,
    OpenAIResponsesModel,
    set_default_openai_api,
    set_default_openai_client,
    set_tracing_disabled,
)
from openai import AsyncOpenAI

ProviderId = Literal[
    "openai",
    "anthropic",
    "google",
    "moonshot",
    "openrouter",
    "custom",
]


@dataclass(frozen=True)
class ResolvedLLM:
    provider: ProviderId
    model: str
    api_key_env: str
    base_url: str | None
    use_chat_completions: bool
    display_label: str


_PROVIDER_DEFAULTS: dict[ProviderId, dict[str, str | bool | None]] = {
    "openai": {
        "api_key_env": "OPENAI_API_KEY",
        "base_url": None,
        "default_model": "gpt-4o",
        "use_chat_completions": False,
        "label": "OpenAI",
    },
    "anthropic": {
        "api_key_env": "ANTHROPIC_API_KEY",
        "base_url": "https://api.anthropic.com/v1/",
        "default_model": "claude-opus-5-5",
        "use_chat_completions": True,
        "label": "Anthropic",
    },
    "google": {
        "api_key_env": "GOOGLE_API_KEY",
        "base_url": "https://generativelanguage.googleapis.com/v1beta/openai/",
        "default_model": "gemini-2.0-flash",
        "use_chat_completions": True,
        "label": "Google Gemini",
    },
    "moonshot": {
        "api_key_env": "MOONSHOT_API_KEY",
        "base_url": "https://api.moonshot.ai/v1",
        "default_model": "kimi-k2.6",
        "use_chat_completions": True,
        "label": "Moonshot (Kimi)",
    },
    "openrouter": {
        "api_key_env": "OPENROUTER_API_KEY",
        "base_url": "https://openrouter.ai/api/v1",
        "default_model": "openai/gpt-4o",
        "use_chat_completions": True,
        "label": "OpenRouter",
    },
    "custom": {
        "api_key_env": "CUSTOM_API_KEY",
        "base_url": None,
        "default_model": "default",
        "use_chat_completions": True,
        "label": "Custom endpoint",
    },
}


def list_providers() -> list[ProviderId]:
    return list(_PROVIDER_DEFAULTS.keys())


def parse_model_spec(spec: str | None) -> tuple[ProviderId | None, str | None]:
    """Parse ``provider/model`` or return (None, model) if no slash."""
    if not spec or not spec.strip():
        return None, None
    spec = spec.strip()
    if "/" in spec:
        provider_part, model_part = spec.split("/", 1)
        provider = provider_part.strip().lower()
        if provider == "kimi":
            provider = "moonshot"
        if provider in _PROVIDER_DEFAULTS:
            return provider, model_part.strip() or None  # type: ignore[return-value]
    return None, spec


def resolve_llm(
    *,
    provider: str | None = None,
    model: str | None = None,
    model_spec: str | None = None,
) -> ResolvedLLM:
    """Resolve provider + model from CLI flags, env, and defaults."""
    spec_provider, spec_model = parse_model_spec(model_spec)
    if model_spec and spec_provider is None and spec_model and "/" in model_spec:
        raise ValueError(
            f"Unknown provider in '{model_spec}'. "
            f"Use one of: {', '.join(list_providers())}."
        )

    env_provider = (os.environ.get("FINANCE_AGENT_PROVIDER") or "").strip().lower()
    if env_provider == "kimi":
        env_provider = "moonshot"

    chosen_provider: ProviderId
    if spec_provider:
        chosen_provider = spec_provider
    elif provider:
        p = provider.strip().lower()
        if p == "kimi":
            p = "moonshot"
        if p not in _PROVIDER_DEFAULTS:
            raise ValueError(f"Unknown provider '{provider}'. Use: {', '.join(list_providers())}.")
        chosen_provider = p  # type: ignore[assignment]
    elif env_provider in _PROVIDER_DEFAULTS:
        chosen_provider = env_provider  # type: ignore[assignment]
    else:
        # Auto-detect from first available API key (vals-style: set only what you use).
        for pid in ("openai", "anthropic", "google", "moonshot", "openrouter", "custom"):
            env_name = str(_PROVIDER_DEFAULTS[pid]["api_key_env"])
            if os.environ.get(env_name):
                chosen_provider = pid  # type: ignore[assignment]
                break
        else:
            chosen_provider = "moonshot"

    meta = _PROVIDER_DEFAULTS[chosen_provider]
    api_key_env = str(meta["api_key_env"])

    chosen_model = (
        spec_model
        or model
        or os.environ.get("FINANCE_AGENT_MODEL")
        or str(meta["default_model"])
    )

    base_url = os.environ.get(f"{chosen_provider.upper()}_BASE_URL") or meta["base_url"]
    if chosen_provider == "custom":
        base_url = os.environ.get("CUSTOM_BASE_URL") or base_url
        if not base_url:
            raise ValueError("CUSTOM_BASE_URL is required when FINANCE_AGENT_PROVIDER=custom.")

    if chosen_provider == "moonshot":
        base_url = os.environ.get("MOONSHOT_BASE_URL") or base_url

    use_cc = bool(meta["use_chat_completions"])
    if chosen_provider == "openai":
        mode = (os.environ.get("OPENAI_API_MODE") or "").strip().lower()
        if mode == "chat_completions":
            use_cc = True
        elif mode == "responses":
            use_cc = False

    label = f"{meta['label']} / {chosen_model}"
    return ResolvedLLM(
        provider=chosen_provider,
        model=chosen_model,
        api_key_env=api_key_env,
        base_url=base_url,
        use_chat_completions=use_cc,
        display_label=label,
    )


def get_api_key(resolved: ResolvedLLM) -> str:
    key = os.environ.get(resolved.api_key_env)
    if key:
        return key
    # Moonshot historically accepted OPENAI_API_KEY against Moonshot URL — avoid that
    # when user explicitly chose another provider.
    if resolved.provider == "moonshot":
        key = os.environ.get("OPENAI_API_KEY")
        if key:
            return key
    raise RuntimeError(
        f"{resolved.api_key_env} is not set. Add it to .env (see .env.example). "
        f"Provider: {resolved.provider}, model: {resolved.model}."
    )


def missing_provider_help() -> str:
    lines = [
        "Set FINANCE_AGENT_PROVIDER and the matching API key in .env, or export a key for auto-detect.",
        "",
        "Providers (set only the keys you use):",
    ]
    for pid in list_providers():
        meta = _PROVIDER_DEFAULTS[pid]
        lines.append(
            f"  • {pid}: {meta['api_key_env']}  (default model: {meta['default_model']})"
        )
    lines.extend(
        [
            "",
            "Examples:",
            "  FINANCE_AGENT_PROVIDER=openai",
            "  FINANCE_AGENT_MODEL=gpt-4o",
            "  python main.py --model anthropic/claude-opus-5-5",
            "  python main.py --model moonshot/kimi-k2.6",
        ]
    )
    return "\n".join(lines)


def configure_openai_client(resolved: ResolvedLLM) -> AsyncOpenAI:
    api_key = get_api_key(resolved)
    kwargs: dict = {"api_key": api_key}
    if resolved.base_url:
        kwargs["base_url"] = resolved.base_url
    client = AsyncOpenAI(**kwargs)

    set_default_openai_client(client, use_for_tracing=False)
    if resolved.use_chat_completions:
        set_default_openai_api("chat_completions")
    else:
        set_default_openai_api("responses")
    # Avoid sending traces to OpenAI when using third-party endpoints / keys.
    if resolved.provider != "openai":
        set_tracing_disabled(True)
    elif os.environ.get("OPENAI_AGENTS_DISABLE_TRACING", "1") == "1":
        set_tracing_disabled(True)
    return client


def create_agent_model(resolved: ResolvedLLM):
    client = configure_openai_client(resolved)
    if resolved.use_chat_completions:
        return OpenAIChatCompletionsModel(model=resolved.model, openai_client=client)
    return OpenAIResponsesModel(model=resolved.model, openai_client=client)
