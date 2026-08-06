"""AI custom-function endpoints (handoff §8).

Two jobs:
  * /ai/batch    — batched AI.* calls, routed to the cheapest tier
  * /ai/forecast — exponential smoothing; NO language model is involved

The forecast split is deliberate. A number a model invented looks exactly like
a number it computed, and in a financial model that difference matters more
than almost anything else. The statistics are computed here; a model may only
ever narrate them.
"""

from __future__ import annotations

from typing import Any, Literal

from fastapi import APIRouter, Request
from pydantic import BaseModel, ConfigDict, Field

from ..auth import CurrentPrincipal
from ..llm import CompletionRequest, CostMeter

router = APIRouter(prefix="/api/v1/ai", tags=["ai"])

# Bounded so one request cannot become an unbounded fan-out.
MAX_BATCH = 100
MAX_SERIES = 10_000
MAX_HORIZON = 120


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid")


class AiRequestItem(StrictModel):
    fn: Literal["AI.ASK", "AI.EXTRACT", "AI.CLASSIFY", "AI.MATCH"]
    inputs: list[Any] = Field(max_length=500)
    prompt: str = Field(max_length=8000)
    model: str = Field(max_length=64)


class BatchRequest(StrictModel):
    requests: list[AiRequestItem] = Field(min_length=1, max_length=MAX_BATCH)


class BatchResponse(BaseModel):
    results: list[Any]
    cost_usd: float


class ForecastRequest(StrictModel):
    series: list[float] = Field(max_length=MAX_SERIES)
    horizon: int = Field(ge=1, le=MAX_HORIZON)
    seasonLength: int = Field(default=0, ge=0, le=52)


class ForecastResponse(BaseModel):
    values: list[float]
    method: str
    alpha: float
    mae: float
    caveat: str | None = None


def _prompt_for(item: AiRequestItem) -> str:
    inputs = "\n".join(str(value) for value in item.inputs)
    if item.fn == "AI.CLASSIFY":
        return (
            f"Classify the following into EXACTLY ONE of these categories: {item.prompt}.\n"
            f"Reply with the category only, no explanation.\n\n{inputs}"
        )
    if item.fn == "AI.EXTRACT":
        return (
            f"Extract the field '{item.prompt}' from the text below. "
            f"Reply with the value only. Reply #N/A if it is absent.\n\n{inputs}"
        )
    if item.fn == "AI.MATCH":
        return (
            f"Do these refer to the same entity? Reply TRUE or FALSE only.\n\n{inputs}"
        )
    return f"{item.prompt}\n\n{inputs}"


@router.post("/batch", response_model=BatchResponse)
async def batch(
    request: Request, body: BatchRequest, principal: CurrentPrincipal
) -> BatchResponse:
    gateway = request.app.state.llm_gateway
    meter: CostMeter = request.app.state.ai_cost_meter

    results: list[Any] = []
    for item in body.requests:
        completion = await gateway.complete(
            CompletionRequest(
                # Custom functions are the cheapest tier by definition — they
                # run thousands of times and the questions are small.
                role="classifier",
                system=(
                    "You answer spreadsheet cell functions. Reply with the value only: "
                    "no preamble, no explanation, no punctuation beyond the value itself."
                ),
                messages=[{"role": "user", "content": _prompt_for(item)}],
                max_tokens=256,
            ),
            meter,
        )
        results.append(completion.text.strip())

    void_principal = principal  # scoping is enforced by the dependency itself
    del void_principal
    return BatchResponse(results=results, cost_usd=round(meter.total_cost_usd, 6))


@router.post("/forecast", response_model=ForecastResponse)
async def forecast_endpoint(
    body: ForecastRequest, principal: CurrentPrincipal
) -> ForecastResponse:
    del principal
    result = _forecast(body.series, body.horizon, body.seasonLength)
    return ForecastResponse(**result)


def _forecast(series: list[float], horizon: int, season_length: int) -> dict[str, Any]:
    """Exponential smoothing: simple, Holt's linear, or additive Holt-Winters.

    Mirrors engine/src/aifn/forecast.ts so the add-in's offline path and the
    server produce the same numbers.
    """
    clean = [value for value in series if value == value and abs(value) != float("inf")]

    if not clean:
        return {
            "values": [0.0] * horizon,
            "method": "simple",
            "alpha": 0.0,
            "mae": 0.0,
            "caveat": "No numeric history was supplied; the forecast is all zeros.",
        }
    if len(clean) < 3:
        last = clean[-1]
        return {
            "values": [last] * horizon,
            "method": "simple",
            "alpha": 1.0,
            "mae": 0.0,
            "caveat": (
                f"Only {len(clean)} data point(s): the forecast simply repeats the last value."
            ),
        }

    grid = [0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9]
    best: dict[str, Any] | None = None

    def consider(candidate: dict[str, Any]) -> None:
        nonlocal best
        if best is None or candidate["mae"] < best["mae"]:
            best = candidate

    def mae_of(actual: list[float], fitted: list[float]) -> float:
        pairs = list(zip(actual, fitted, strict=False))
        if not pairs:
            return 0.0
        return sum(abs(a - f) for a, f in pairs) / len(pairs)

    seasonal_ok = season_length >= 2 and len(clean) >= season_length * 2
    if seasonal_ok:
        for alpha in grid:
            for beta in grid:
                for gamma in (0.1, 0.3, 0.5, 0.7):
                    seasonal = []
                    first_cycle = clean[:season_length]
                    first_mean = sum(first_cycle) / season_length
                    for index in range(season_length):
                        seasonal.append(clean[index] - first_mean)
                    level = first_mean
                    second_cycle = clean[season_length : season_length * 2]
                    trend = (sum(second_cycle) / season_length - first_mean) / season_length
                    fitted = []
                    for index, value in enumerate(clean):
                        season_index = index % season_length
                        previous_level = level
                        level = alpha * (value - seasonal[season_index]) + (1 - alpha) * (
                            level + trend
                        )
                        trend = beta * (level - previous_level) + (1 - beta) * trend
                        seasonal[season_index] = gamma * (value - level) + (1 - gamma) * seasonal[
                            season_index
                        ]
                        fitted.append(level + trend + seasonal[season_index])
                    values = [
                        level + step * trend + seasonal[(len(clean) + step - 1) % season_length]
                        for step in range(1, horizon + 1)
                    ]
                    consider(
                        {
                            "values": values,
                            "method": "holt-winters",
                            "alpha": alpha,
                            "mae": mae_of(clean, fitted),
                        }
                    )

    for alpha in grid:
        for beta in grid:
            level = clean[0]
            trend = (clean[1] if len(clean) > 1 else level) - level
            fitted = [level]
            for value in clean[1:]:
                previous_level = level
                level = alpha * value + (1 - alpha) * (level + trend)
                trend = beta * (level - previous_level) + (1 - beta) * trend
                fitted.append(level + trend)
            values = [level + step * trend for step in range(1, horizon + 1)]
            consider(
                {
                    "values": values,
                    "method": "holt",
                    "alpha": alpha,
                    "mae": mae_of(clean, fitted),
                }
            )

    result = best or {
        "values": [clean[-1]] * horizon,
        "method": "simple",
        "alpha": 1.0,
        "mae": 0.0,
    }
    if season_length >= 2 and result["method"] != "holt-winters":
        result["caveat"] = (
            f"Seasonality of {season_length} was requested but the series has only "
            f"{len(clean)} points — at least {season_length * 2} are needed, so a "
            f"non-seasonal model was used instead."
        )
    result.setdefault("caveat", None)
    return result
