"""Tool-call envelope endpoints.

Phase 0 flow (client-initiated):
    1. POST /tool-calls           -> server validates tool + params (INV-1), mints id
    2. add-in executes the validated call against the workbook via Office.js
    3. POST /tool-calls/{id}/result -> server validates result shape, logs it

The same envelope carries agent-initiated calls from Phase 3 (server pushes
validated calls, add-in executes and posts results).
"""

from datetime import UTC, datetime

from fastapi import APIRouter, HTTPException, Request

from ..auth import CurrentPrincipal
from ..config import get_settings
from ..models import (
    ToolCall,
    ToolCallAccepted,
    ToolCallRequest,
    ToolResultAck,
    ToolResultRequest,
    new_id,
)
from ..schema_registry import SchemaValidationError, ToolRegistry
from .sessions import load_session_or_404, store_of

router = APIRouter(prefix="/api/v1/sessions/{session_id}/tool-calls", tags=["tools"])


def registry_of(request: Request) -> ToolRegistry:
    return request.app.state.tool_registry


def _validation_422(exc: SchemaValidationError) -> HTTPException:
    return HTTPException(
        status_code=422,
        detail={
            "code": f"invalid_{exc.kind}",
            "tool": exc.tool,
            "errors": exc.errors,
        },
    )


@router.post("", response_model=ToolCallAccepted, status_code=201)
async def create_tool_call(
    request: Request,
    session_id: str,
    body: ToolCallRequest,
    principal: CurrentPrincipal,
) -> ToolCallAccepted:
    session = await load_session_or_404(request, session_id, principal)

    spec = registry_of(request).get(body.tool)
    if spec is None:
        raise HTTPException(
            status_code=404,
            detail={
                "code": "unknown_tool",
                "tool": body.tool,
                "known_tools": registry_of(request).tool_names,
            },
        )
    try:
        registry_of(request).validate_params(body.tool, body.params)
    except SchemaValidationError as exc:
        raise _validation_422(exc) from exc

    call = ToolCall(
        id=new_id("tc"),
        tool=body.tool,
        params=body.params,
        access=spec.access,
        status="validated",
    )
    session.tool_calls.append(call)
    # Bound the log (oldest first out).
    excess = len(session.tool_calls) - get_settings().max_tool_calls_per_session
    if excess > 0:
        del session.tool_calls[:excess]
    await store_of(request).save(session)
    return ToolCallAccepted(
        tool_call_id=call.id,
        tool=call.tool,
        params=call.params,
        access=call.access,
        status=call.status,
    )


@router.post("/{tool_call_id}/result", response_model=ToolResultAck)
async def post_tool_result(
    request: Request,
    session_id: str,
    tool_call_id: str,
    body: ToolResultRequest,
    principal: CurrentPrincipal,
) -> ToolResultAck:
    session = await load_session_or_404(request, session_id, principal)
    call = session.find_call(tool_call_id)
    if call is None:
        raise HTTPException(status_code=404, detail={"code": "unknown_tool_call"})
    if call.status != "validated":
        raise HTTPException(
            status_code=409,
            detail={"code": "tool_call_already_resolved", "status": call.status},
        )

    if body.ok:
        if body.result is None:
            raise HTTPException(
                status_code=422, detail={"code": "missing_result", "tool": call.tool}
            )
        try:
            registry_of(request).validate_result(call.tool, body.result)
        except SchemaValidationError as exc:
            raise _validation_422(exc) from exc
        call.status = "completed"
        call.result = body.result
        summary = f"{call.tool} completed"
    else:
        if body.error is None:
            raise HTTPException(
                status_code=422, detail={"code": "missing_error", "tool": call.tool}
            )
        call.status = "failed"
        call.error = body.error
        summary = f"{call.tool} failed: {body.error.code}"

    call.completed_at = datetime.now(UTC)
    await store_of(request).save(session)
    return ToolResultAck(tool_call_id=call.id, status=call.status, summary=summary)
