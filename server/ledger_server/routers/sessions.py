from fastapi import APIRouter, HTTPException, Request

from ..auth import CurrentPrincipal
from ..config import get_settings
from ..models import Session, SessionCreated, SessionInfo, new_id
from ..sessions import SessionStore

router = APIRouter(prefix="/api/v1/sessions", tags=["sessions"])


def store_of(request: Request) -> SessionStore:
    return request.app.state.session_store


async def load_session_or_404(
    request: Request, session_id: str, principal: CurrentPrincipal
) -> Session:
    session = await store_of(request).get(session_id)
    # A session is private to the principal that created it.
    if session is None or session.principal_subject != principal.subject:
        raise HTTPException(status_code=404, detail="Session not found")
    return session


@router.post("", response_model=SessionCreated, status_code=201)
async def create_session(request: Request, principal: CurrentPrincipal) -> SessionCreated:
    session = Session(id=new_id("ses"), principal_subject=principal.subject)
    await store_of(request).create(session)
    return SessionCreated(
        session_id=session.id,
        created_at=session.created_at,
        auth_mode=get_settings().auth_mode,
    )


@router.get("/{session_id}", response_model=SessionInfo)
async def get_session(
    request: Request, session_id: str, principal: CurrentPrincipal
) -> SessionInfo:
    session = await load_session_or_404(request, session_id, principal)
    return SessionInfo(
        session_id=session.id,
        created_at=session.created_at,
        tool_call_count=len(session.tool_calls),
        tool_calls=session.tool_calls,
    )
