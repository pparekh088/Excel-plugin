"""Telemetry tests.

The important property is negative: it must be IMPOSSIBLE to record workbook
content, even by accident. A financial model is the most confidential thing
most users own.
"""

import pytest

from ledger_server.telemetry import (
    Telemetry,
    TelemetryRedactionError,
    record_audit,
    record_changeset,
    session_hash,
)


def test_disabled_by_default():
    assert Telemetry().enabled is False


def test_disabled_telemetry_records_nothing():
    telemetry = Telemetry(enabled=False)
    assert telemetry.record("audit.run", "ses_1", duration_ms=5) is None
    assert telemetry.events == []


@pytest.mark.parametrize(
    "forbidden",
    [
        {"formula": "=A1*2"},
        {"value": 42},
        {"sheet": "Assumptions"},
        {"address": "Sheet1!A1"},
        {"intent": "change growth to 5.2%"},
        {"summary": "Set the growth driver"},
        {"explanation": "what changed and why"},
        {"prompt": "classify this"},
        {"diff": [{"before": 1}]},
        {"names": ["WACC"]},
        {"filename": "Project Atlas.xlsx"},
    ],
)
def test_content_bearing_fields_are_refused(forbidden):
    telemetry = Telemetry(enabled=True)
    with pytest.raises(TelemetryRedactionError):
        telemetry.record("audit.run", "ses_1", **forbidden)


def test_nested_content_fields_are_refused():
    telemetry = Telemetry(enabled=True)
    with pytest.raises(TelemetryRedactionError):
        telemetry.record("audit.run", "ses_1", detail={"inner": {"formula": "=A1"}})


def test_long_strings_are_refused_as_a_leak_vector():
    telemetry = Telemetry(enabled=True)
    with pytest.raises(TelemetryRedactionError) as excinfo:
        telemetry.record("audit.run", "ses_1", note="x" * 500)
    assert "long strings" in str(excinfo.value)


def test_counts_and_categories_are_allowed():
    telemetry = Telemetry(enabled=True)
    event = telemetry.record(
        "audit.run",
        "ses_1",
        duration_ms=120,
        finding_count=7,
        health_score=64,
        risk="high",
    )
    assert event is not None
    assert event.properties["finding_count"] == 7


def test_session_id_is_hashed_not_stored():
    telemetry = Telemetry(enabled=True)
    event = telemetry.record("session.started", "ses_secret_identifier", duration_ms=1)
    assert event is not None
    assert "ses_secret_identifier" not in event.to_json()
    assert event.session_hash == session_hash("ses_secret_identifier")


def test_hash_is_stable_and_correlates_events():
    telemetry = Telemetry(enabled=True)
    first = telemetry.record("session.started", "ses_1", duration_ms=1)
    second = telemetry.record("audit.run", "ses_1", duration_ms=2)
    assert first is not None and second is not None
    assert first.session_hash == second.session_hash


def test_audit_recorder_captures_rule_ids_but_no_addresses():
    telemetry = Telemetry(enabled=True)
    record_audit(
        telemetry,
        "ses_1",
        duration_ms=45,
        findings_by_rule={"AUD-001": 2, "AUD-004": 1},
        health_score=72,
        formula_cells=1200,
        coverage_complete=False,
    )
    payload = telemetry.events[0].to_json()
    assert "AUD-001" in payload
    assert "finding_count" in payload
    assert telemetry.events[0].properties["finding_count"] == 3


def test_changeset_recorder_captures_risk_and_size():
    telemetry = Telemetry(enabled=True)
    record_changeset(
        telemetry,
        "ses_1",
        "changeset.applied",
        risk="high",
        edit_count=12,
        affected_cells=340,
        duration_ms=900,
        repair_attempts=1,
    )
    properties = telemetry.events[0].properties
    assert properties["risk"] == "high"
    assert properties["repair_attempts"] == 1


def test_event_buffer_is_bounded():
    telemetry = Telemetry(enabled=True)
    telemetry.max_retained = 10
    for index in range(50):
        telemetry.record("audit.run", "ses_1", duration_ms=index)
    assert len(telemetry.events) == 10


def test_counts_summarize_by_kind():
    telemetry = Telemetry(enabled=True)
    telemetry.record("audit.run", "ses_1", duration_ms=1)
    telemetry.record("audit.run", "ses_1", duration_ms=2)
    telemetry.record("changeset.applied", "ses_1", risk="low")
    assert telemetry.counts() == {"audit.run": 2, "changeset.applied": 1}
