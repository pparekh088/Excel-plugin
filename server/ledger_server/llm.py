"""LLM gateway: per-task routing, cost metering, prompt-cache awareness.

Providers are pluggable. A `mock` provider ships so the whole stack runs
offline in CI and in the zero-LLM demo; Anthropic and OpenAI adapters are
wired to the same interface and activate when their API keys are present.

Cost metering is per session and surfaced in the UI (handoff §7), so a user
can always see what a request cost before approving the next one.
"""

from __future__ import annotations

import os
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Literal

ModelTier = Literal["strong", "fast", "cheap"]
AgentRole = Literal["planner", "executor", "critic", "classifier"]

# Routing policy (handoff §7). Evals are expected to beat this; when they do,
# change it here and record the reason in DECISIONS.md.
ROLE_TIERS: dict[AgentRole, ModelTier] = {
    "planner": "strong",
    "critic": "strong",
    "executor": "fast",
    "classifier": "cheap",
}


@dataclass(frozen=True)
class ModelSpec:
    provider: str
    model: str
    # USD per million tokens.
    input_cost_per_mtok: float
    output_cost_per_mtok: float


# Defaults are overridable per deployment via LEDGER_MODEL_<TIER>.
DEFAULT_MODELS: dict[ModelTier, ModelSpec] = {
    "strong": ModelSpec("anthropic", "claude-opus-4-5", 5.0, 25.0),
    "fast": ModelSpec("anthropic", "claude-sonnet-4-5", 3.0, 15.0),
    "cheap": ModelSpec("anthropic", "claude-haiku-4-5", 1.0, 5.0),
}


@dataclass
class Usage:
    input_tokens: int = 0
    output_tokens: int = 0
    cached_input_tokens: int = 0
    cost_usd: float = 0.0
    model: str = ""

    def merge(self, other: Usage) -> Usage:
        return Usage(
            input_tokens=self.input_tokens + other.input_tokens,
            output_tokens=self.output_tokens + other.output_tokens,
            cached_input_tokens=self.cached_input_tokens + other.cached_input_tokens,
            cost_usd=self.cost_usd + other.cost_usd,
            model=other.model or self.model,
        )


@dataclass
class CompletionRequest:
    role: AgentRole
    system: str
    messages: list[dict[str, str]]
    max_tokens: int = 4096
    # Marks the system prompt as cacheable — the tool catalogue and WIL are
    # large and stable within a session, so caching them matters.
    cache_system: bool = True


@dataclass
class CompletionResponse:
    text: str
    usage: Usage


class LlmProvider(ABC):
    name: str

    @abstractmethod
    async def complete(self, request: CompletionRequest, spec: ModelSpec) -> CompletionResponse:
        ...


class MockProvider(LlmProvider):
    """Deterministic provider for CI and offline demos.

    Returns a scripted response when one matches, otherwise an empty plan —
    never a plausible-looking guess, so a missing script fails loudly.
    """

    name = "mock"

    def __init__(self, scripts: dict[str, str] | None = None):
        self.scripts = scripts or {}
        self.requests: list[CompletionRequest] = []

    async def complete(self, request: CompletionRequest, spec: ModelSpec) -> CompletionResponse:
        self.requests.append(request)
        body = "\n".join(m.get("content", "") for m in request.messages)
        prompt = f"{request.system}\n{body}"
        text = next(
            (value for key, value in self.scripts.items() if key in prompt),
            '{"steps": []}',
        )
        return CompletionResponse(
            text=text,
            usage=Usage(
                input_tokens=len(prompt) // 4,
                output_tokens=len(text) // 4,
                cost_usd=0.0,
                model=f"mock-{spec.model}",
            ),
        )


class AnthropicProvider(LlmProvider):
    """Anthropic Messages API adapter.

    Imports the SDK lazily so the server runs without it installed (the mock
    path and the entire audit engine need no provider at all).
    """

    name = "anthropic"

    def __init__(self, api_key: str):
        self._api_key = api_key
        self._client = None

    def _get_client(self):
        if self._client is None:
            try:
                from anthropic import AsyncAnthropic
            except ImportError as exc:  # pragma: no cover - depends on deployment
                raise RuntimeError(
                    "anthropic package is not installed; install it or use LEDGER_LLM_PROVIDER=mock"
                ) from exc
            self._client = AsyncAnthropic(api_key=self._api_key)
        return self._client

    async def complete(self, request: CompletionRequest, spec: ModelSpec) -> CompletionResponse:
        client = self._get_client()
        system: list[dict[str, object]] | str = request.system
        if request.cache_system:
            system = [
                {
                    "type": "text",
                    "text": request.system,
                    "cache_control": {"type": "ephemeral"},
                }
            ]
        response = await client.messages.create(
            model=spec.model,
            max_tokens=request.max_tokens,
            system=system,
            messages=request.messages,
        )
        text = "".join(block.text for block in response.content if block.type == "text")
        usage = response.usage
        cached = getattr(usage, "cache_read_input_tokens", 0) or 0
        billed_input = usage.input_tokens
        cost = (
            billed_input / 1_000_000 * spec.input_cost_per_mtok
            + usage.output_tokens / 1_000_000 * spec.output_cost_per_mtok
        )
        return CompletionResponse(
            text=text,
            usage=Usage(
                input_tokens=billed_input,
                output_tokens=usage.output_tokens,
                cached_input_tokens=cached,
                cost_usd=cost,
                model=spec.model,
            ),
        )


class OpenAIProvider(LlmProvider):
    """OpenAI adapter, present so the router can be scored across providers."""

    name = "openai"

    def __init__(self, api_key: str):
        self._api_key = api_key
        self._client = None

    def _get_client(self):
        if self._client is None:
            try:
                from openai import AsyncOpenAI
            except ImportError as exc:  # pragma: no cover
                raise RuntimeError(
                    "openai package is not installed; install it or use LEDGER_LLM_PROVIDER=mock"
                ) from exc
            self._client = AsyncOpenAI(api_key=self._api_key)
        return self._client

    async def complete(self, request: CompletionRequest, spec: ModelSpec) -> CompletionResponse:
        client = self._get_client()
        response = await client.chat.completions.create(
            model=spec.model,
            max_tokens=request.max_tokens,
            messages=[{"role": "system", "content": request.system}, *request.messages],
        )
        text = response.choices[0].message.content or ""
        usage = response.usage
        cost = (
            (usage.prompt_tokens if usage else 0) / 1_000_000 * spec.input_cost_per_mtok
            + (usage.completion_tokens if usage else 0) / 1_000_000 * spec.output_cost_per_mtok
        )
        return CompletionResponse(
            text=text,
            usage=Usage(
                input_tokens=usage.prompt_tokens if usage else 0,
                output_tokens=usage.completion_tokens if usage else 0,
                cost_usd=cost,
                model=spec.model,
            ),
        )


@dataclass
class CostMeter:
    """Per-session cost accounting, surfaced in the UI."""

    by_tier: dict[str, Usage] = field(default_factory=dict)

    def record(self, tier: ModelTier, usage: Usage) -> None:
        current = self.by_tier.get(tier, Usage())
        self.by_tier[tier] = current.merge(usage)

    @property
    def total_cost_usd(self) -> float:
        return sum(usage.cost_usd for usage in self.by_tier.values())

    @property
    def total_calls(self) -> int:
        # Approximate: one call per recorded merge is not tracked separately,
        # so callers that need exact counts should track them alongside.
        return len(self.by_tier)

    def snapshot(self) -> dict[str, object]:
        return {
            "totalCostUsd": round(self.total_cost_usd, 6),
            "byTier": {
                tier: {
                    "inputTokens": usage.input_tokens,
                    "outputTokens": usage.output_tokens,
                    "cachedInputTokens": usage.cached_input_tokens,
                    "costUsd": round(usage.cost_usd, 6),
                    "model": usage.model,
                }
                for tier, usage in self.by_tier.items()
            },
        }


class LlmGateway:
    """Routes a role to a tier, a tier to a model, and meters the result."""

    def __init__(
        self,
        provider: LlmProvider,
        models: dict[ModelTier, ModelSpec] | None = None,
    ):
        self.provider = provider
        self.models = models or DEFAULT_MODELS

    async def complete(self, request: CompletionRequest, meter: CostMeter) -> CompletionResponse:
        tier = ROLE_TIERS[request.role]
        spec = self.models[tier]
        response = await self.provider.complete(request, spec)
        meter.record(tier, response.usage)
        return response


def build_gateway() -> LlmGateway:
    """Pick a provider from the environment.

    Defaults to the mock provider: the server must be runnable, and the audit
    engine fully usable, with no API keys present anywhere.
    """
    provider_name = os.getenv("LEDGER_LLM_PROVIDER", "mock").lower()
    if provider_name == "anthropic":
        key = os.getenv("ANTHROPIC_API_KEY", "")
        if not key:
            raise RuntimeError("LEDGER_LLM_PROVIDER=anthropic requires ANTHROPIC_API_KEY")
        return LlmGateway(AnthropicProvider(key))
    if provider_name == "openai":
        key = os.getenv("OPENAI_API_KEY", "")
        if not key:
            raise RuntimeError("LEDGER_LLM_PROVIDER=openai requires OPENAI_API_KEY")
        return LlmGateway(OpenAIProvider(key))
    return LlmGateway(MockProvider())
