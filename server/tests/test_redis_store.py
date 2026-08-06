"""Redis-backed session store tests. Skipped automatically when Redis is not running."""

import pytest
import redis

from ledger_server.config import Settings
from ledger_server.models import Session, new_id
from ledger_server.sessions import RedisSessionStore

REDIS_URL = "redis://localhost:6379/15"


def _redis_available() -> bool:
    try:
        redis.Redis.from_url(REDIS_URL, socket_connect_timeout=0.5).ping()
        return True
    except Exception:
        return False


pytestmark = pytest.mark.skipif(not _redis_available(), reason="Redis not running")


@pytest.mark.anyio
async def test_redis_session_persistence():
    store = RedisSessionStore(REDIS_URL, ttl_seconds=60)
    session = Session(id=new_id("ses"), principal_subject="tester")
    await store.create(session)
    loaded = await store.get(session.id)
    assert loaded is not None
    assert loaded.id == session.id
    assert loaded.principal_subject == "tester"
    assert await store.get("ses_missing") is None


@pytest.fixture
def anyio_backend():
    return "asyncio"


def test_settings_pick_redis_store():
    from ledger_server.sessions import build_session_store

    store = build_session_store(Settings(redis_url=REDIS_URL))
    assert store.name == "redis"
