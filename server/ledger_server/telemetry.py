"""Telemetry (handoff §10 Phase 5).

What we record and why: enough to answer "is the product working and is it
safe" without collecting anything that could reconstruct a customer's model.

NEVER recorded: cell values, formula text, sheet names, defined names, file
names, or anything else derived from workbook content. A financial model is
the most confidential thing most of our users own, and an audit tool that
leaks it is worthless regardless of how good the audit is.

Recorded: counts, durations, rule ids, risk tiers, outcomes, and error CLASSES.
All of which are answerable from structure alone.
"""

from __future__ import annotations

import json
import logging
import os
from dataclasses import asdict, dataclass, field
from datetime import UTC, datetime
from typing import Any, Literal

logger = logging.getLogger("ledger.telemetry")

EventKind = Literal[
    "session.started",
    "wil.built",
    "audit.run",
    "plan.proposed",
    "changeset.previewed",
    "changeset.applied",
    "changeset.rolled_back",
    "changeset.aborted_drift",
    "changeset.blocked_hazard",
    "verify.failed",
    "repair.attempted",
    "ai.batch",
    "ai.budget_exceeded",
    "error",
]

# Field names that must never appear in a telemetry payload. Enforced rather
# than trusted, because the easy mistake is passing a whole object through.
FORBIDDEN_FIELDS = frozenset(
    {
        "formula",
        "formulas",
        "value",
        "values",
        "text",
        "sheet",
        "sheets",
        "sheet_name",
        "address",
        "addresses",
        "explanation",
        "intent",
        "summary",
        "prompt",
        "diff",
        "workbook",
        "name",
        "names",
        "filename",
        "file_name",
        "content",
    }
)


class TelemetryRedactionError(ValueError):
    """A payload carried a field that could contain workbook content."""


@dataclass
class TelemetryEvent:
    kind: EventKind
    # Hashed session id: correlates events without identifying a session.
    session_hash: str
    properties: dict[str, Any] = field(default_factory=dict)
    timestamp: str = field(default_factory=lambda: datetime.now(UTC).isoformat())

    def to_json(self) -> str:
        return json.dumps(asdict(self), separators=(",", ":"))


def _assert_safe(properties: dict[str, Any], path: str = "") -> None:
    for key, value in properties.items():
        location = f"{path}.{key}" if path else key
        if key.lower() in FORBIDDEN_FIELDS:
            raise TelemetryRedactionError(
                f"Telemetry property '{location}' could contain workbook content and is refused. "
                f"Record a count, a duration or a category instead."
            )
        if isinstance(value, dict):
            _assert_safe(value, location)
        elif isinstance(value, str) and len(value) > 200:
            raise TelemetryRedactionError(
                f"Telemetry property '{location}' is {len(value)} characters; long strings are "
                f"refused because they are how content leaks."
            )


def session_hash(session_id: str) -> str:
    """Stable, non-reversible correlator for one session."""
    import hashlib

    return hashlib.sha256(session_id.encode()).hexdigest()[:16]


class Telemetry:
    """Emits events to a logger. A real sink (SQL, App Insights) plugs in here.

    Disabled by default: LEDGER_TELEMETRY=on turns it on. Off means no events
    are constructed at all, not merely not sent.
    """

    def __init__(self, enabled: bool | None = None):
        self.enabled = (
            enabled
            if enabled is not None
            else os.getenv("LEDGER_TELEMETRY", "off").lower() in {"on", "1", "true"}
        )
        self.events: list[TelemetryEvent] = []
        # Bounded so a long session cannot grow without limit in memory.
        self.max_retained = 1000

    def record(self, kind: EventKind, session_id: str, **properties: Any) -> TelemetryEvent | None:
        if not self.enabled:
            return None
        _assert_safe(properties)
        event = TelemetryEvent(
            kind=kind, session_hash=session_hash(session_id), properties=properties
        )
        self.events.append(event)
        if len(self.events) > self.max_retained:
            del self.events[: len(self.events) - self.max_retained]
        logger.info("ledger.telemetry %s", event.to_json())
        return event

    def counts(self) -> dict[str, int]:
        out: dict[str, int] = {}
        for event in self.events:
            out[event.kind] = out.get(event.kind, 0) + 1
        return out

    def clear(self) -> None:
        self.events.clear()


# Convenience recorders that encode the safe shape of each event, so callers
# do not have to remember what is allowed.


def record_audit(
    telemetry: Telemetry,
    session_id: str,
    *,
    duration_ms: int,
    findings_by_rule: dict[str, int],
    health_score: int,
    formula_cells: int,
    coverage_complete: bool,
) -> None:
    telemetry.record(
        "audit.run",
        session_id,
        duration_ms=duration_ms,
        findings_by_rule=findings_by_rule,
        finding_count=sum(findings_by_rule.values()),
        health_score=health_score,
        formula_cells=formula_cells,
        coverage_complete=coverage_complete,
    )


def record_changeset(
    telemetry: Telemetry,
    session_id: str,
    kind: EventKind,
    *,
    risk: str,
    edit_count: int,
    affected_cells: int,
    duration_ms: int | None = None,
    repair_attempts: int = 0,
) -> None:
    telemetry.record(
        kind,
        session_id,
        risk=risk,
        edit_count=edit_count,
        affected_cells=affected_cells,
        **({"duration_ms": duration_ms} if duration_ms is not None else {}),
        repair_attempts=repair_attempts,
    )
