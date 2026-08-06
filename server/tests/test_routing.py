"""Model-routing tests.

The bug these exist to prevent: a tier->model table whose entries all named
Anthropic models, combined with a separately-chosen single provider, meant
selecting OpenAI sent it `model="claude-opus-4-5"`. A route now names both,
and a mismatch is rejected at construction.
"""

import pytest

from ledger_server.llm import (
    CompletionRequest,
    CostMeter,
    LlmGateway,
    MockProvider,
    ModelRoute,
    RoutingError,
    Usage,
    validate_route,
)


class RecordingProvider(MockProvider):
    """Mock that remembers which model it was asked for."""

    def __init__(self, name):
        super().__init__()
        self.name = name
        self.models: list[str] = []

    async def complete(self, request, route):
        self.models.append(route.model)
        return await super().complete(request, route)


def test_a_provider_never_receives_another_providers_model():
    with pytest.raises(RoutingError) as excinfo:
        validate_route("planner", ModelRoute("openai", "claude-opus-4-5", "strong", 1, 1))
    assert "does not serve it" in str(excinfo.value)


def test_valid_pairings_are_accepted():
    validate_route("planner", ModelRoute("anthropic", "claude-opus-4-5", "strong", 1, 1))
    validate_route("executor", ModelRoute("openai", "gpt-5.6", "fast", 1, 1))
    validate_route("classifier", ModelRoute("openai", "o4-mini", "cheap", 1, 1))


def test_gateway_rejects_a_route_to_an_unconfigured_provider():
    with pytest.raises(RoutingError) as excinfo:
        LlmGateway(
            {"mock": MockProvider()},
            {"planner": ModelRoute("anthropic", "claude-opus-4-5", "strong", 1, 1)},
        )
    assert "not configured" in str(excinfo.value)


@pytest.mark.anyio
async def test_mixed_provider_routing_sends_each_role_to_its_own_provider():
    anthropic = RecordingProvider("anthropic")
    openai = RecordingProvider("openai")
    gateway = LlmGateway(
        {"anthropic": anthropic, "openai": openai},
        {
            "planner": ModelRoute("anthropic", "claude-opus-4-5", "strong", 5, 25),
            "executor": ModelRoute("openai", "gpt-5.6", "fast", 3, 15),
            "critic": ModelRoute("anthropic", "claude-opus-4-5", "strong", 5, 25),
            "classifier": ModelRoute("openai", "gpt-5.6-mini", "cheap", 1, 5),
        },
    )
    meter = CostMeter()
    for role in ("planner", "executor", "critic", "classifier"):
        await gateway.complete(
            CompletionRequest(role=role, system="s", messages=[{"role": "user", "content": "x"}]),
            meter,
        )

    # Each provider saw only its own models.
    assert all(model.startswith("claude-") for model in anthropic.models)
    assert all(model.startswith("gpt-") for model in openai.models)
    assert len(anthropic.models) == 2
    assert len(openai.models) == 2


@pytest.fixture
def anyio_backend():
    return "asyncio"


def test_route_lookup_exposes_the_tier():
    gateway = LlmGateway(
        {"mock": MockProvider()},
        {
            "planner": ModelRoute("mock", "mock-strong", "strong", 0, 0),
            "executor": ModelRoute("mock", "mock-fast", "fast", 0, 0),
            "critic": ModelRoute("mock", "mock-strong", "strong", 0, 0),
            "classifier": ModelRoute("mock", "mock-cheap", "cheap", 0, 0),
        },
    )
    assert gateway.route("planner").tier == "strong"
    assert gateway.route("classifier").tier == "cheap"


class TestCostMeter:
    def test_total_calls_counts_calls_not_distinct_models(self):
        """The old implementation returned len(by_tier): 30 calls read as 2."""
        meter = CostMeter()
        strong = ModelRoute("anthropic", "claude-opus-4-5", "strong", 5, 25)
        cheap = ModelRoute("anthropic", "claude-haiku-4-5", "cheap", 1, 5)
        for _ in range(20):
            meter.record(strong, Usage(input_tokens=10, calls=1, model=strong.model))
        for _ in range(10):
            meter.record(cheap, Usage(input_tokens=5, calls=1, model=cheap.model))

        assert meter.total_calls == 30
        assert meter.total_input_tokens == 250

    def test_costs_are_keyed_by_provider_and_model(self):
        meter = CostMeter()
        meter.record(
            ModelRoute("anthropic", "claude-opus-4-5", "strong", 5, 25),
            Usage(cost_usd=0.01, calls=1, provider="anthropic", model="claude-opus-4-5"),
        )
        meter.record(
            ModelRoute("openai", "gpt-5.6", "fast", 3, 15),
            Usage(cost_usd=0.002, calls=1, provider="openai", model="gpt-5.6"),
        )
        snapshot = meter.snapshot()
        assert "anthropic/claude-opus-4-5" in snapshot["byModel"]
        assert "openai/gpt-5.6" in snapshot["byModel"]
        assert snapshot["totalCalls"] == 2
        assert snapshot["totalCostUsd"] == pytest.approx(0.012)
