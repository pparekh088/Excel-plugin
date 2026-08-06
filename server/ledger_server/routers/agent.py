"""Agent endpoints: planning, change-set records, and cost.

Division of labour (D-004, D-011): the deterministic engine — parser, graph,
WIL, audit, change-set computation — runs in TypeScript, in the add-in. The
server owns sessions, auth, the LLM gateway, and the durable record of what
was proposed and applied.

That record is the point: a change set that reached a workbook must be
recoverable from the server even if the add-in is closed, because "what did
the agent do to my model" is an auditor's question, not a UI concern.
"""

from datetime import UTC, datetime
from typing import Any, Literal

from fastapi import APIRouter, HTTPException, Request
from pydantic import BaseModel, ConfigDict, Field

from ..auth import CurrentPrincipal
from ..llm import CompletionRequest, CostMeter
from ..models import new_id
from .sessions import load_session_or_404

router = APIRouter(prefix="/api/v1/sessions/{session_id}/agent", tags=["agent"])


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid")


class PlanRequest(StrictModel):
    intent: str = Field(min_length=1, max_length=4000)
    # The add-in builds the WIL and sends it; the server never sees raw grids
    # (INV-6) and never needs the workbook itself.
    wil: str = Field(min_length=1, max_length=200_000)
    tool_catalogue: str = Field(min_length=1, max_length=200_000)
    role: Literal["planner", "executor", "critic", "classifier"] = "planner"


class PlanResponse(BaseModel):
    text: str
    model: str
    tier: str
    cost_usd: float
    input_tokens: int
    output_tokens: int
    cached_input_tokens: int


class ChangeSetRecord(StrictModel):
    """The durable record of a change set, as computed by the engine."""

    change_set_id: str = Field(min_length=1, max_length=100)
    intent: str = Field(max_length=4000)
    summary: str = Field(max_length=4000)
    risk: Literal["low", "medium", "high"]
    status: Literal["proposed", "previewed", "applied", "rolled-back", "aborted", "failed"]
    edit_count: int = Field(ge=0)
    affected_cells: int = Field(ge=0)
    explanation: str = Field(max_length=100_000)
    diff: list[dict[str, Any]] = Field(default_factory=list, max_length=5000)
    failure_reason: str | None = None


class ChangeSetAck(BaseModel):
    change_set_id: str
    status: str
    recorded_at: datetime


class CostResponse(BaseModel):
    total_cost_usd: float
    by_tier: dict[str, Any]


def _meter_for(session_id: str, request: Request) -> CostMeter:
    meters: dict[str, CostMeter] = request.app.state.cost_meters
    if session_id not in meters:
        meters[session_id] = CostMeter()
    return meters[session_id]


@router.post("/plan", response_model=PlanResponse)
async def plan(
    request: Request,
    session_id: str,
    body: PlanRequest,
    principal: CurrentPrincipal,
) -> PlanResponse:
    await load_session_or_404(request, session_id, principal)

    gateway = request.app.state.llm_gateway
    meter = _meter_for(session_id, request)
    completion = await gateway.complete(
        CompletionRequest(
            role=body.role,
            # Catalogue and WIL are stable within a session, so they go in the
            # cacheable system block rather than the per-turn messages.
            system=f"{body.tool_catalogue}\n\nWORKBOOK SUMMARY\n{body.wil}",
            messages=[{"role": "user", "content": f"INTENT\n{body.intent}"}],
        ),
        meter,
    )
    from ..llm import ROLE_TIERS

    tier = ROLE_TIERS[body.role]
    return PlanResponse(
        text=completion.text,
        model=completion.usage.model,
        tier=tier,
        cost_usd=round(completion.usage.cost_usd, 6),
        input_tokens=completion.usage.input_tokens,
        output_tokens=completion.usage.output_tokens,
        cached_input_tokens=completion.usage.cached_input_tokens,
    )


@router.post("/change-sets", response_model=ChangeSetAck, status_code=201)
async def record_change_set(
    request: Request,
    session_id: str,
    body: ChangeSetRecord,
    principal: CurrentPrincipal,
) -> ChangeSetAck:
    session = await load_session_or_404(request, session_id, principal)

    records: dict[str, list[dict[str, Any]]] = request.app.state.change_sets
    entries = records.setdefault(session.id, [])
    entry = body.model_dump()
    entry["recorded_at"] = datetime.now(UTC).isoformat()
    # Supersede an earlier record of the same change set: status advances
    # (proposed -> applied -> rolled-back) and the latest state is the truth.
    entries[:] = [item for item in entries if item["change_set_id"] != body.change_set_id]
    entries.append(entry)

    return ChangeSetAck(
        change_set_id=body.change_set_id,
        status=body.status,
        recorded_at=datetime.now(UTC),
    )


@router.get("/change-sets")
async def list_change_sets(
    request: Request, session_id: str, principal: CurrentPrincipal
) -> dict[str, Any]:
    session = await load_session_or_404(request, session_id, principal)
    records: dict[str, list[dict[str, Any]]] = request.app.state.change_sets
    return {"change_sets": records.get(session.id, [])}


@router.get("/change-sets/{change_set_id}")
async def get_change_set(
    request: Request, session_id: str, change_set_id: str, principal: CurrentPrincipal
) -> dict[str, Any]:
    session = await load_session_or_404(request, session_id, principal)
    records: dict[str, list[dict[str, Any]]] = request.app.state.change_sets
    for entry in records.get(session.id, []):
        if entry["change_set_id"] == change_set_id:
            return entry
    raise HTTPException(status_code=404, detail={"code": "unknown_change_set"})


@router.get("/cost", response_model=CostResponse)
async def cost(request: Request, session_id: str, principal: CurrentPrincipal) -> CostResponse:
    await load_session_or_404(request, session_id, principal)
    snapshot = _meter_for(session_id, request).snapshot()
    return CostResponse(
        total_cost_usd=snapshot["totalCostUsd"],  # type: ignore[arg-type]
        by_tier=snapshot["byTier"],  # type: ignore[arg-type]
    )


def new_change_set_id() -> str:
    return new_id("cs")
