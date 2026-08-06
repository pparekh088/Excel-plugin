"""Pydantic models for the session + tool-call envelope.

The envelope is the persistent shape of INV-1: the only thing that ever reaches
a workbook is a tool call that was schema-validated server-side first. In Phase 0
the add-in initiates calls (client asks the server to validate, executes locally,
posts the result back); from Phase 3 the agent runtime mints calls server-side
and pushes them to the add-in over the same envelope.

Tool params/results are validated against the generated JSON Schemas in
shared/schemas (see schema_registry.py), not per-tool pydantic models — the
schemas are data, exported from the Zod source of truth in the add-in.
"""

import uuid
from datetime import UTC, datetime
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field


def _now() -> datetime:
    return datetime.now(UTC)


def new_id(prefix: str) -> str:
    return f"{prefix}_{uuid.uuid4().hex[:20]}"


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid")


# ---------------------------------------------------------------- tool calls


class ToolCallRequest(StrictModel):
    tool: str = Field(min_length=1, max_length=100)
    params: dict[str, Any] = Field(default_factory=dict)


class ToolError(StrictModel):
    code: str = Field(min_length=1, max_length=100)
    message: str = Field(max_length=4000)
    detail: dict[str, Any] | None = None


class ToolResultRequest(StrictModel):
    ok: bool
    result: dict[str, Any] | None = None
    error: ToolError | None = None


ToolCallStatus = Literal["validated", "completed", "failed"]


class ToolCall(BaseModel):
    id: str
    tool: str
    params: dict[str, Any]
    access: Literal["read", "write", "control"]
    status: ToolCallStatus
    created_at: datetime = Field(default_factory=_now)
    completed_at: datetime | None = None
    result: dict[str, Any] | None = None
    error: ToolError | None = None


class ToolCallAccepted(BaseModel):
    tool_call_id: str
    tool: str
    params: dict[str, Any]
    access: Literal["read", "write", "control"]
    status: ToolCallStatus


class ToolResultAck(BaseModel):
    tool_call_id: str
    status: ToolCallStatus
    summary: str


# ------------------------------------------------------------------ sessions


class Session(BaseModel):
    id: str
    created_at: datetime = Field(default_factory=_now)
    principal_subject: str
    tool_calls: list[ToolCall] = Field(default_factory=list)

    def find_call(self, tool_call_id: str) -> ToolCall | None:
        for call in self.tool_calls:
            if call.id == tool_call_id:
                return call
        return None


class SessionCreated(BaseModel):
    session_id: str
    created_at: datetime
    auth_mode: str


class SessionInfo(BaseModel):
    session_id: str
    created_at: datetime
    tool_call_count: int
    tool_calls: list[ToolCall]


# -------------------------------------------------------------------- health


class Health(BaseModel):
    status: Literal["ok"]
    version: str
    auth_mode: str
    session_store: str
    tools: list[str]
