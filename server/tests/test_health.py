def test_healthz_reports_status_and_tools(client):
    response = client.get("/healthz")
    assert response.status_code == 200
    body = response.json()
    assert body["status"] == "ok"
    assert body["auth_mode"] == "dev"
    assert body["session_store"] == "memory"
    assert "range.read" in body["tools"]
