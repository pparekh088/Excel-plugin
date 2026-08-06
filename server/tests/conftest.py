import pytest
from fastapi.testclient import TestClient

from ledger_server.config import Settings
from ledger_server.main import create_app


@pytest.fixture
def settings() -> Settings:
    # In-memory session store, dev auth, schemas from the checked-in shared dir.
    return Settings(auth_mode="dev", redis_url="")


@pytest.fixture
def client(settings: Settings):
    app = create_app(settings)
    with TestClient(app) as test_client:
        yield test_client


@pytest.fixture
def session_id(client: TestClient) -> str:
    response = client.post("/api/v1/sessions")
    assert response.status_code == 201
    return response.json()["session_id"]
