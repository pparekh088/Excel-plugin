"""Session store: Redis-backed with an in-memory fallback for local dev.

Sessions hold the tool-call log (and, from Phase 3, snapshots + change sets).
Both stores expose the same async interface; the factory picks Redis when
LEDGER_REDIS_URL is set. The in-memory store is single-process only and says
so in /healthz.
"""

import time
from typing import Protocol

import redis.asyncio as aioredis

from .config import Settings
from .models import Session


class SessionStore(Protocol):
    name: str

    async def create(self, session: Session) -> None: ...

    async def get(self, session_id: str) -> Session | None: ...

    async def save(self, session: Session) -> None: ...

    async def ping(self) -> bool: ...


class InMemorySessionStore:
    """Dev-only fallback. Not shared across processes; TTL enforced lazily."""

    name = "memory"

    def __init__(self, ttl_seconds: int):
        self._ttl = ttl_seconds
        self._data: dict[str, tuple[float, str]] = {}

    def _evict_expired(self) -> None:
        now = time.monotonic()
        expired = [sid for sid, (expires, _) in self._data.items() if expires < now]
        for sid in expired:
            del self._data[sid]

    async def create(self, session: Session) -> None:
        await self.save(session)

    async def get(self, session_id: str) -> Session | None:
        self._evict_expired()
        entry = self._data.get(session_id)
        if entry is None:
            return None
        return Session.model_validate_json(entry[1])

    async def save(self, session: Session) -> None:
        self._evict_expired()
        self._data[session.id] = (time.monotonic() + self._ttl, session.model_dump_json())

    async def ping(self) -> bool:
        return True


class RedisSessionStore:
    name = "redis"

    def __init__(self, url: str, ttl_seconds: int):
        self._redis = aioredis.from_url(url, decode_responses=True)
        self._ttl = ttl_seconds

    @staticmethod
    def _key(session_id: str) -> str:
        return f"ledger:session:{session_id}"

    async def create(self, session: Session) -> None:
        await self.save(session)

    async def get(self, session_id: str) -> Session | None:
        raw = await self._redis.get(self._key(session_id))
        if raw is None:
            return None
        return Session.model_validate_json(raw)

    async def save(self, session: Session) -> None:
        await self._redis.set(self._key(session.id), session.model_dump_json(), ex=self._ttl)

    async def ping(self) -> bool:
        try:
            return bool(await self._redis.ping())
        except Exception:
            return False


def build_session_store(settings: Settings) -> SessionStore:
    if settings.redis_url:
        return RedisSessionStore(settings.redis_url, settings.session_ttl_seconds)
    return InMemorySessionStore(settings.session_ttl_seconds)
