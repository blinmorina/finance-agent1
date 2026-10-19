from __future__ import annotations

import os

import pytest

from llm_config import parse_model_spec, resolve_llm


def test_parse_model_spec_provider_slash_model():
    p, m = parse_model_spec("openai/gpt-4o")
    assert p == "openai"
    assert m == "gpt-4o"


def test_parse_model_spec_kimi_alias():
    p, m = parse_model_spec("kimi/kimi-k2.6")
    assert p == "moonshot"
    assert m == "kimi-k2.6"


def test_resolve_llm_from_model_spec(monkeypatch):
    monkeypatch.delenv("FINANCE_AGENT_PROVIDER", raising=False)
    monkeypatch.delenv("FINANCE_AGENT_MODEL", raising=False)
    r = resolve_llm(model_spec="anthropic/claude-opus-5-5")
    assert r.provider == "anthropic"
    assert r.model == "claude-opus-5-5"
    assert r.use_chat_completions is True


def test_resolve_llm_env_provider(monkeypatch):
    monkeypatch.setenv("FINANCE_AGENT_PROVIDER", "openai")
    monkeypatch.setenv("FINANCE_AGENT_MODEL", "gpt-4o-mini")
    r = resolve_llm()
    assert r.provider == "openai"
    assert r.model == "gpt-4o-mini"


def test_unknown_provider_in_spec_raises():
    with pytest.raises(ValueError, match="Unknown provider"):
        resolve_llm(model_spec="notaprovider/foo")
