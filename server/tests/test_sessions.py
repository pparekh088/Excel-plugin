def test_create_and_get_session(client):
    created = client.post("/api/v1/sessions")
    assert created.status_code == 201
    body = created.json()
    assert body["session_id"].startswith("ses_")
    assert body["auth_mode"] == "dev"

    fetched = client.get(f"/api/v1/sessions/{body['session_id']}")
    assert fetched.status_code == 200
    assert fetched.json()["tool_call_count"] == 0


def test_get_unknown_session_is_404(client):
    assert client.get("/api/v1/sessions/ses_does_not_exist").status_code == 404
