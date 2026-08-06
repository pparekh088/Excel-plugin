"""The Phase 0 gate test: range.read envelope round trip with INV-1 validation."""

VALID_PARAMS = {"sheet": "Sheet1", "a1": "A1:D10", "include": ["values", "formulas"]}

VALID_RESULT = {
    "sheet": "Sheet1",
    "a1": "Sheet1!A1:D10",
    "rowCount": 2,
    "columnCount": 2,
    "cellCount": 4,
    "chunkCount": 1,
    "values": [[1, "x"], [True, None]],
    "formulas": [["=B1", "x"], ["TRUE", ""]],
}


def _invoke(client, session_id, tool="range.read", params=None):
    return client.post(
        f"/api/v1/sessions/{session_id}/tool-calls",
        json={"tool": tool, "params": VALID_PARAMS if params is None else params},
    )


def test_full_roundtrip(client, session_id):
    accepted = _invoke(client, session_id)
    assert accepted.status_code == 201
    call = accepted.json()
    assert call["tool_call_id"].startswith("tc_")
    assert call["access"] == "read"
    assert call["status"] == "validated"

    ack = client.post(
        f"/api/v1/sessions/{session_id}/tool-calls/{call['tool_call_id']}/result",
        json={"ok": True, "result": VALID_RESULT},
    )
    assert ack.status_code == 200
    assert ack.json()["status"] == "completed"

    info = client.get(f"/api/v1/sessions/{session_id}").json()
    assert info["tool_call_count"] == 1
    assert info["tool_calls"][0]["status"] == "completed"
    assert info["tool_calls"][0]["result"]["cellCount"] == 4


def test_unknown_tool_rejected(client, session_id):
    response = _invoke(client, session_id, tool="range.hack")
    assert response.status_code == 404
    assert response.json()["detail"]["code"] == "unknown_tool"


def test_invalid_params_rejected(client, session_id):
    # Missing required field.
    response = _invoke(client, session_id, params={"a1": "A1:B2"})
    assert response.status_code == 422
    assert response.json()["detail"]["code"] == "invalid_params"

    # Malformed A1 reference.
    response = _invoke(client, session_id, params={"sheet": "S", "a1": "not-a-range"})
    assert response.status_code == 422

    # Unknown extra property (schemas are strict).
    response = _invoke(
        client, session_id, params={**VALID_PARAMS, "unexpected": 1}
    )
    assert response.status_code == 422

    # Sheet-qualified a1 is rejected (sheet is a separate field).
    response = _invoke(client, session_id, params={"sheet": "S", "a1": "Sheet1!A1"})
    assert response.status_code == 422


def test_invalid_result_rejected_and_call_stays_open(client, session_id):
    call_id = _invoke(client, session_id).json()["tool_call_id"]
    bad_result = {**VALID_RESULT, "rowCount": "two"}
    response = client.post(
        f"/api/v1/sessions/{session_id}/tool-calls/{call_id}/result",
        json={"ok": True, "result": bad_result},
    )
    assert response.status_code == 422
    assert response.json()["detail"]["code"] == "invalid_result"

    # The call is still open, so a corrected result can be posted.
    ok = client.post(
        f"/api/v1/sessions/{session_id}/tool-calls/{call_id}/result",
        json={"ok": True, "result": VALID_RESULT},
    )
    assert ok.status_code == 200


def test_error_result_marks_call_failed(client, session_id):
    call_id = _invoke(client, session_id).json()["tool_call_id"]
    ack = client.post(
        f"/api/v1/sessions/{session_id}/tool-calls/{call_id}/result",
        json={
            "ok": False,
            "error": {"code": "RANGE_TOO_LARGE", "message": "3M cells exceeds budget"},
        },
    )
    assert ack.status_code == 200
    assert ack.json()["status"] == "failed"


def test_result_for_unknown_call_is_404(client, session_id):
    response = client.post(
        f"/api/v1/sessions/{session_id}/tool-calls/tc_nope/result",
        json={"ok": True, "result": VALID_RESULT},
    )
    assert response.status_code == 404


def test_double_resolution_is_409(client, session_id):
    call_id = _invoke(client, session_id).json()["tool_call_id"]
    url = f"/api/v1/sessions/{session_id}/tool-calls/{call_id}/result"
    assert client.post(url, json={"ok": True, "result": VALID_RESULT}).status_code == 200
    assert client.post(url, json={"ok": True, "result": VALID_RESULT}).status_code == 409
