"""Ledger Agent API entry point.

Run locally:
    uvicorn ledger_server.main:app --reload --port 8000
"""

from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from .config import Settings, get_settings
from .llm import build_gateway
from .routers import agent, ai, health, sessions, tools
from .schema_registry import ToolRegistry
from .sessions import build_session_store


def create_app(settings: Settings | None = None) -> FastAPI:
    settings = settings or get_settings()

    if settings.env == "production" and settings.auth_mode == "dev":
        raise RuntimeError("auth_mode=dev is not allowed with LEDGER_ENV=production")

    @asynccontextmanager
    async def lifespan(app: FastAPI):
        app.state.tool_registry = ToolRegistry.load(settings.schemas_dir)
        app.state.session_store = build_session_store(settings)
        app.state.llm_gateway = build_gateway()
        # Per-session cost meters and change-set records. In-process for now;
        # they move to Redis alongside sessions when the deployment target is
        # multi-worker (tracked in PROGRESS).
        app.state.cost_meters = {}
        app.state.change_sets = {}
        # AI cost meters keyed by (principal, session) — never process-global.
        app.state.ai_cost_meters = {}
        yield

    app = FastAPI(title="Ledger Agent API", version="0.0.1", lifespan=lifespan)
    app.add_middleware(
        CORSMiddleware,
        allow_origins=settings.cors_origins,
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )
    app.include_router(health.router)
    app.include_router(sessions.router)
    app.include_router(tools.router)
    app.include_router(agent.router)
    app.include_router(ai.router)
    return app


app = create_app()
