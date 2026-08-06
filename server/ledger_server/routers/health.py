from fastapi import APIRouter, Request

from .. import __version__
from ..config import get_settings
from ..models import Health

router = APIRouter(tags=["health"])


@router.get("/healthz", response_model=Health)
async def healthz(request: Request) -> Health:
    settings = get_settings()
    return Health(
        status="ok",
        version=__version__,
        auth_mode=settings.auth_mode,
        session_store=request.app.state.session_store.name,
        tools=request.app.state.tool_registry.tool_names,
    )
