"""LLM gateway: per-role routing across providers, with cost metering.

Routing is per ROLE, and a route names both the provider and the model. The
earlier design had one configured provider plus a tier->model table whose
entries all named Anthropic models, so selecting the OpenAI provider sent it
`model="claude-opus-4-5"`. A route is now the unit of configuration, the
gateway holds every configured provider simultaneously, and a provider is only
ever handed a model belonging to it — asserted at construction, not hoped for.

Providers are optional: the mock provider is the default so the server runs
with no API keys and no SDKs installed, which is what keeps the zero-LLM audit
demo honest.
"""

from __future__ import annotations

import os
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Literal

ModelTier = Literal["strong", "fast", "cheap"]
AgentRole = Literal["planner", "executor", "critic", "classifier"]
ProviderName = Literal["anthropic", "openai", "mock"]


class RoutingError(RuntimeError):
    """A route names a provider that is not configured, or a mismatched model."""


@dataclass(frozen=True)
class ModelRoute:
    """Where one role's traffic goes, and what it costs."""

    provider: ProviderName
    model: str
    tier: ModelTier
    input_cost_per_mtok: float
    output_cost_per_mtok: float
    max_tokens: int = 4096
    # Reasoning-effort profile, for providers that expose one.
    reasoning_effort: str | None = None


# Default policy (handoff §7): planner and critic get the strongest model,
# executor a fast one, classification the cheapest. Evals are expected to beat
# this; when they do, change it here and record why in DECISIONS.md.
#
# Every model here belongs to the provider named beside it — a mismatch is a
# configuration bug and is rejected at startup.
DEFAULT_ROUTES: dict[AgentRole, ModelRoute] = {
    "planner": ModelRoute("anthropic", "claude-opus-4-5", "strong", 5.0, 25.0),
    "critic": ModelRoute("anthropic", "claude-opus-4-5", "strong", 5.0, 25.0),
    "executor": ModelRoute("anthropic", "claude-sonnet-4-5", "fast", 3.0, 15.0),
    "classifier": ModelRoute("anthropic", "claude-haiku-4-5", "cheap", 1.0, 5.0),
}

# Model-name prefixes that identify a provider, used to catch routes that pair
# a provider with another provider's model.
_MODEL_PREFIXES: dict[ProviderName, tuple[str, ...]] = {
    "anthropic": ("claude-",),
    "openai": ("gpt-", "o1", "o3", "o4"),
    "mock": (),
}


def validate_route(role: AgentRole, route: ModelRoute) -> None:
    prefixes = _MODEL_PREFIXES.get(route.provider, ())
    if prefixes and not route.model.startswith(prefixes):
        raise RoutingError(
            f"Route for '{role}' sends model '{route.model}' to provider "
            f"'{route.provider}', which does not serve it. Fix the routing policy."
        )


@dataclass
class Usage:
    input_tokens: int = 0
    output_tokens: int = 0
    cached_input_tokens: int = 0
    cost_usd: float = 0.0
    model: str = ""
    provider: str = ""
    # Real counter, not inferred from how many tiers were touched.
    calls: int = 0

    def merge(self, other: Usage) -> Usage:
        return Usage(
            input_tokens=self.input_tokens + other.input_tokens,
            output_tokens=self.output_tokens + other.output_tokens,
            cached_input_tokens=self.cached_input_tokens + other.cached_input_tokens,
            cost_usd=self.cost_usd + other.cost_usd,
            model=other.model or self.model,
            provider=other.provider or self.provider,
            calls=self.calls + other.calls,
        )


@dataclass
class CompletionRequest:
    role: AgentRole
    system: str
    messages: list[dict[str, str]]
    max_tokens: int | None = None
    # The system block is large and stable within a session, so it is the part
    # worth caching.
    cache_system: bool = True


@dataclass
class CompletionResponse:
    text: str
    usage: Usage


class LlmProvider(ABC):
    name: ProviderName

    @abstractmethod
    async def complete(self, request: CompletionRequest, route: ModelRoute) -> CompletionResponse:
        ...


class MockProvider(LlmProvider):
    """Deterministic provider for CI and offline demos.

    Returns a scripted response when one matches, otherwise an empty plan —
    never a plausible-looking guess, so a missing script fails loudly.
    """

    name: ProviderName = "mock"

    def __init__(self, scripts: dict[str, str] | None = None):
        self.scripts = scripts or {}
        self.requests: list[CompletionRequest] = []

    async def complete(self, request: CompletionRequest, route: ModelRoute) -> CompletionResponse:
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
                model=f"mock-{route.model}",
                provider="mock",
                calls=1,
            ),
        )


class AnthropicProvider(LlmProvider):
    """Anthropic Messages API adapter. SDK imported lazily."""

    name: ProviderName = "anthropic"

    def __init__(self, api_key: str):
        self._api_key = api_key
        self._client = None

    def _get_client(self):
        if self._client is None:
            try:
                from anthropic import AsyncAnthropic
            except ImportError as exc:  # pragma: no cover - depends on deployment
                raise RuntimeError(
                    "anthropic package is not installed; install it or route to the mock provider"
                ) from exc
            self._client = AsyncAnthropic(api_key=self._api_key)
        return self._client

    async def complete(self, request: CompletionRequest, route: ModelRoute) -> CompletionResponse:
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
            model=route.model,
            max_tokens=request.max_tokens or route.max_tokens,
            system=system,
            messages=request.messages,
        )
        text = "".join(block.text for block in response.content if block.type == "text")
        usage = response.usage
        cached = getattr(usage, "cache_read_input_tokens", 0) or 0
        cost = (
            usage.input_tokens / 1_000_000 * route.input_cost_per_mtok
            + usage.output_tokens / 1_000_000 * route.output_cost_per_mtok
        )
        return CompletionResponse(
            text=text,
            usage=Usage(
                input_tokens=usage.input_tokens,
                output_tokens=usage.output_tokens,
                cached_input_tokens=cached,
                cost_usd=cost,
                model=route.model,
                provider="anthropic",
                calls=1,
            ),
        )


class OpenAIProvider(LlmProvider):
    """OpenAI adapter, so the router can be scored across providers."""

    name: ProviderName = "openai"

    def __init__(self, api_key: str):
        self._api_key = api_key
        self._client = None

    def _get_client(self):
        if self._client is None:
            try:
                from openai import AsyncOpenAI
            except ImportError as exc:  # pragma: no cover
                raise RuntimeError(
                    "openai package is not installed; install it or route to the mock provider"
                ) from exc
            self._client = AsyncOpenAI(api_key=self._api_key)
        return self._client

    async def complete(self, request: CompletionRequest, route: ModelRoute) -> CompletionResponse:
        client = self._get_client()
        extra: dict[str, object] = {}
        if route.reasoning_effort:
            extra["reasoning_effort"] = route.reasoning_effort
        response = await client.chat.completions.create(
            model=route.model,
            max_completion_tokens=request.max_tokens or route.max_tokens,
            messages=[{"role": "system", "content": request.system}, *request.messages],
            **extra,
        )
        text = response.choices[0].message.content or ""
        usage = response.usage
        prompt_tokens = usage.prompt_tokens if usage else 0
        completion_tokens = usage.completion_tokens if usage else 0
        cost = (
            prompt_tokens / 1_000_000 * route.input_cost_per_mtok
            + completion_tokens / 1_000_000 * route.output_cost_per_mtok
        )
        return CompletionResponse(
            text=text,
            usage=Usage(
                input_tokens=prompt_tokens,
                output_tokens=completion_tokens,
                cost_usd=cost,
                model=route.model,
                provider="openai",
                calls=1,
            ),
        )


@dataclass
class CostMeter:
    """Cost accounting. One meter per scope — never one per process.

    Keyed by (provider, model) as well as tier, because "which model did this
    cost come from" is the question routing decisions are made on.
    """

    by_key: dict[str, Usage] = field(default_factory=dict)

    def record(self, route: ModelRoute, usage: Usage) -> None:
        key = f"{route.provider}/{route.model}"
        self.by_key[key] = self.by_key.get(key, Usage()).merge(usage)

    @property
    def total_cost_usd(self) -> float:
        return sum(usage.cost_usd for usage in self.by_key.values())

    @property
    def total_calls(self) -> int:
        """Actual number of completions, not the number of distinct models."""
        return sum(usage.calls for usage in self.by_key.values())

    @property
    def total_input_tokens(self) -> int:
        return sum(usage.input_tokens for usage in self.by_key.values())

    @property
    def total_output_tokens(self) -> int:
        return sum(usage.output_tokens for usage in self.by_key.values())

    def snapshot(self) -> dict[str, object]:
        return {
            "totalCostUsd": round(self.total_cost_usd, 6),
            "totalCalls": self.total_calls,
            "byModel": {
                key: {
                    "provider": usage.provider,
                    "model": usage.model,
                    "calls": usage.calls,
                    "inputTokens": usage.input_tokens,
                    "outputTokens": usage.output_tokens,
                    "cachedInputTokens": usage.cached_input_tokens,
                    "costUsd": round(usage.cost_usd, 6),
                }
                for key, usage in self.by_key.items()
            },
        }


class LlmGateway:
    """Routes a role to (provider, model), calls it, and meters the result."""

    def __init__(
        self,
        providers: dict[ProviderName, LlmProvider],
        routes: dict[AgentRole, ModelRoute] | None = None,
    ):
        self.providers = providers
        self.routes = routes or DEFAULT_ROUTES
        for role, route in self.routes.items():
            validate_route(role, route)
            if route.provider not in self.providers:
                raise RoutingError(
                    f"Route for '{role}' names provider '{route.provider}', which is not "
                    f"configured. Configured: {sorted(self.providers)}."
                )

    def route(self, role: AgentRole) -> ModelRoute:
        return self.routes[role]

    async def complete(self, request: CompletionRequest, meter: CostMeter) -> CompletionResponse:
        route = self.route(request.role)
        provider = self.providers[route.provider]
        response = await provider.complete(request, route)
        meter.record(route, response.usage)
        return response


def _routes_from_env() -> dict[AgentRole, ModelRoute]:
    """Per-role overrides: LEDGER_ROUTE_PLANNER="anthropic:claude-opus-4-5"."""
    routes = dict(DEFAULT_ROUTES)
    for role in ("planner", "executor", "critic", "classifier"):
        raw = os.getenv(f"LEDGER_ROUTE_{role.upper()}")
        if not raw:
            continue
        if ":" not in raw:
            raise RoutingError(
                f"LEDGER_ROUTE_{role.upper()}='{raw}' must be 'provider:model'."
            )
        provider, model = raw.split(":", 1)
        base = routes[role]  # type: ignore[index]
        routes[role] = ModelRoute(  # type: ignore[index]
            provider=provider,  # type: ignore[arg-type]
            model=model,
            tier=base.tier,
            input_cost_per_mtok=base.input_cost_per_mtok,
            output_cost_per_mtok=base.output_cost_per_mtok,
            max_tokens=base.max_tokens,
            reasoning_effort=base.reasoning_effort,
        )
    return routes


def build_gateway() -> LlmGateway:
    """Configure every provider whose key is present; default to mock only.

    The server must be runnable, and the audit engine fully usable, with no API
    keys anywhere — so an absent key is not an error, it just means that
    provider is not available for routing.
    """
    providers: dict[ProviderName, LlmProvider] = {"mock": MockProvider()}
    if os.getenv("ANTHROPIC_API_KEY"):
        providers["anthropic"] = AnthropicProvider(os.environ["ANTHROPIC_API_KEY"])
    if os.getenv("OPENAI_API_KEY"):
        providers["openai"] = OpenAIProvider(os.environ["OPENAI_API_KEY"])

    forced = os.getenv("LEDGER_LLM_PROVIDER", "").lower()
    if forced == "mock" or (not forced and len(providers) == 1):
        # No real provider configured (or explicitly forced): route everything
        # to the mock so the stack runs offline.
        mock_routes = {
            role: ModelRoute("mock", f"mock-{route.tier}", route.tier, 0.0, 0.0)
            for role, route in DEFAULT_ROUTES.items()
        }
        return LlmGateway(providers, mock_routes)  # type: ignore[arg-type]

    return LlmGateway(providers, _routes_from_env())
