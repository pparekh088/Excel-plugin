"""Agent endpoint tests: planning, change-set records, cost metering."""

CHANGE_SET = {
    "change_set_id": "cs_abc123",
    "intent": "Change FY2025 growth to 5.2%",
    "summary": "Set the FY2025 growth driver",
    "risk": "medium",
    "status": "proposed",
    "edit_count": 1,
    "affected_cells": 42,
    "explanation": "What changed: Assumptions!C2 0.10 -> 0.052",
    "diff": [{"address": "Assumptions!C2", "before": 0.1, "after": 0.052}],
}


def test_plan_routes_through_the_gateway_and_meters_cost(client, session_id):
    response = client.post(
        f"/api/v1/sessions/{session_id}/agent/plan",
        json={
            "intent": "Change growth to 5.2%",
            "wil": "WORKBOOK Test\nSHEET Assumptions",
        },
    )
    assert response.status_code == 200
    body = response.json()
    # The mock provider returns an empty plan rather than a plausible guess.
    assert body["text"] == '{"steps": []}'
    assert body["tier"] == "strong"
    assert body["input_tokens"] > 0


def test_planner_role_routes_to_strong_executor_to_fast(client, session_id):
    strong = client.post(
        f"/api/v1/sessions/{session_id}/agent/plan",
        json={"intent": "x", "wil": "w", "role": "planner"},
    ).json()
    fast = client.post(
        f"/api/v1/sessions/{session_id}/agent/plan",
        json={"intent": "x", "wil": "w", "role": "executor"},
    ).json()
    cheap = client.post(
        f"/api/v1/sessions/{session_id}/agent/plan",
        json={"intent": "x", "wil": "w", "role": "classifier"},
    ).json()
    assert strong["tier"] == "strong"
    assert fast["tier"] == "fast"
    assert cheap["tier"] == "cheap"


def test_plan_rejects_unknown_fields(client, session_id):
    response = client.post(
        f"/api/v1/sessions/{session_id}/agent/plan",
        json={"intent": "x", "wil": "w", "workbook": "raw grid"},
    )
    assert response.status_code == 422


def test_plan_requires_an_intent(client, session_id):
    response = client.post(
        f"/api/v1/sessions/{session_id}/agent/plan",
        json={"intent": "", "wil": "w"},
    )
    assert response.status_code == 422


def test_change_set_is_recorded_and_retrievable(client, session_id):
    created = client.post(
        f"/api/v1/sessions/{session_id}/agent/change-sets", json=CHANGE_SET
    )
    assert created.status_code == 201
    assert created.json()["change_set_id"] == "cs_abc123"

    listed = client.get(f"/api/v1/sessions/{session_id}/agent/change-sets").json()
    assert len(listed["change_sets"]) == 1
    assert listed["change_sets"][0]["intent"] == CHANGE_SET["intent"]

    fetched = client.get(
        f"/api/v1/sessions/{session_id}/agent/change-sets/cs_abc123"
    ).json()
    assert fetched["affected_cells"] == 42
    assert fetched["explanation"].startswith("What changed")


def test_recording_the_same_change_set_advances_its_status(client, session_id):
    client.post(f"/api/v1/sessions/{session_id}/agent/change-sets", json=CHANGE_SET)
    applied = {**CHANGE_SET, "status": "applied"}
    client.post(f"/api/v1/sessions/{session_id}/agent/change-sets", json=applied)

    listed = client.get(f"/api/v1/sessions/{session_id}/agent/change-sets").json()
    # One record, latest state — not a duplicate.
    assert len(listed["change_sets"]) == 1
    assert listed["change_sets"][0]["status"] == "applied"


def test_rolled_back_change_sets_keep_their_record(client, session_id):
    client.post(f"/api/v1/sessions/{session_id}/agent/change-sets", json=CHANGE_SET)
    rolled_back = {
        **CHANGE_SET,
        "status": "rolled-back",
        "failure_reason": "verification failed after 3 repairs",
    }
    client.post(f"/api/v1/sessions/{session_id}/agent/change-sets", json=rolled_back)

    fetched = client.get(
        f"/api/v1/sessions/{session_id}/agent/change-sets/cs_abc123"
    ).json()
    assert fetched["status"] == "rolled-back"
    assert "verification failed" in fetched["failure_reason"]


def test_unknown_change_set_is_404(client, session_id):
    response = client.get(
        f"/api/v1/sessions/{session_id}/agent/change-sets/cs_nope"
    )
    assert response.status_code == 404


def test_invalid_risk_tier_is_rejected(client, session_id):
    response = client.post(
        f"/api/v1/sessions/{session_id}/agent/change-sets",
        json={**CHANGE_SET, "risk": "catastrophic"},
    )
    assert response.status_code == 422


def test_cost_accumulates_across_calls(client, session_id):
    before = client.get(f"/api/v1/sessions/{session_id}/agent/cost").json()
    assert before["total_cost_usd"] == 0

    for _ in range(3):
        client.post(
            f"/api/v1/sessions/{session_id}/agent/plan",
            json={"intent": "x", "wil": "w"},
        )

    after = client.get(f"/api/v1/sessions/{session_id}/agent/cost").json()
    # Keyed by provider/model now, and the call count is real rather than
    # inferred from how many distinct tiers were touched.
    assert after["total_calls"] == 3
    assert len(after["by_model"]) >= 1
    assert sum(entry["inputTokens"] for entry in after["by_model"].values()) > 0


def test_agent_endpoints_are_scoped_to_the_session_owner(client):
    response = client.get("/api/v1/sessions/ses_someone_else/agent/change-sets")
    assert response.status_code == 404
