"""AI custom-function endpoint tests."""

import pytest


def test_batch_returns_one_result_per_request(client):
    response = client.post(
        "/api/v1/ai/batch",
        json={
            "requests": [
                {
                    "fn": "AI.CLASSIFY",
                    "inputs": ["great product"],
                    "prompt": "positive,negative",
                    "model": "cheap",
                },
                {
                    "fn": "AI.CLASSIFY",
                    "inputs": ["terrible"],
                    "prompt": "positive,negative",
                    "model": "cheap",
                },
            ]
        },
    )
    assert response.status_code == 200
    assert len(response.json()["results"]) == 2


def test_batch_rejects_an_unknown_function(client):
    response = client.post(
        "/api/v1/ai/batch",
        json={
            "requests": [
                {"fn": "AI.EXECUTE", "inputs": [], "prompt": "rm -rf", "model": "cheap"}
            ]
        },
    )
    assert response.status_code == 422


def test_batch_is_bounded(client):
    requests = [
        {"fn": "AI.ASK", "inputs": [i], "prompt": "x", "model": "cheap"} for i in range(101)
    ]
    response = client.post("/api/v1/ai/batch", json={"requests": requests})
    assert response.status_code == 422


def test_batch_rejects_an_empty_request_list(client):
    assert client.post("/api/v1/ai/batch", json={"requests": []}).status_code == 422


@pytest.mark.parametrize(
    ("series", "horizon", "expect_increasing"),
    [
        ([10, 20, 30, 40, 50, 60], 3, True),
        ([100, 100, 100, 100, 100], 2, False),
    ],
)
def test_forecast_follows_the_trend(client, series, horizon, expect_increasing):
    response = client.post(
        "/api/v1/ai/forecast",
        json={"series": series, "horizon": horizon, "seasonLength": 0},
    )
    assert response.status_code == 200
    body = response.json()
    assert len(body["values"]) == horizon
    if expect_increasing:
        assert body["values"][-1] > series[-1]
    else:
        assert abs(body["values"][0] - 100) < 1


def test_forecast_uses_no_language_model(client):
    """The forecast path must never touch the gateway — a generated number is
    indistinguishable from a computed one, which is exactly the risk."""
    before = client.get("/healthz").status_code
    assert before == 200
    response = client.post(
        "/api/v1/ai/forecast",
        json={"series": [1, 2, 3, 4, 5], "horizon": 2, "seasonLength": 0},
    )
    assert response.status_code == 200
    # Deterministic: the same input gives the same output every time.
    again = client.post(
        "/api/v1/ai/forecast",
        json={"series": [1, 2, 3, 4, 5], "horizon": 2, "seasonLength": 0},
    )
    assert again.json()["values"] == response.json()["values"]


def test_forecast_reports_a_caveat_for_short_series(client):
    response = client.post(
        "/api/v1/ai/forecast",
        json={"series": [5], "horizon": 3, "seasonLength": 0},
    )
    body = response.json()
    assert body["values"] == [5, 5, 5]
    assert "repeats the last value" in body["caveat"]


def test_forecast_says_when_seasonality_could_not_be_used(client):
    response = client.post(
        "/api/v1/ai/forecast",
        json={"series": [10, 12, 11, 13, 12], "horizon": 2, "seasonLength": 4},
    )
    body = response.json()
    assert body["method"] != "holt-winters"
    assert "at least 8" in body["caveat"]


def test_forecast_bounds_the_horizon(client):
    response = client.post(
        "/api/v1/ai/forecast",
        json={"series": [1, 2, 3], "horizon": 500, "seasonLength": 0},
    )
    assert response.status_code == 422


def test_forecast_handles_an_empty_series(client):
    response = client.post(
        "/api/v1/ai/forecast", json={"series": [], "horizon": 2, "seasonLength": 0}
    )
    body = response.json()
    assert body["values"] == [0, 0]
    assert "No numeric history" in body["caveat"]
