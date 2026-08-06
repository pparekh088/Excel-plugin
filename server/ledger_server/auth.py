"""Authentication.

Two modes (LEDGER_AUTH_MODE):

- dev:   no token required. Every request gets a fixed dev principal. Refused
         when LEDGER_ENV=production so it can never leak into a deployment.
- entra: Entra ID (Azure AD) bearer-token validation via the tenant JWKS.
         The add-in acquires tokens with MSAL/NAA (nested app authentication)
         and sends them as Authorization: Bearer <token>.

The Entra path follows the standard app-registration pattern: the add-in's SPA
registration requests an access token for this API's app registration
(audience = api://<client-id> or the client id itself, both accepted).
"""

from dataclasses import dataclass
from functools import lru_cache
from typing import Annotated

import jwt
from fastapi import Depends, HTTPException, Request
from jwt import PyJWKClient

from .config import Settings, get_settings


@dataclass(frozen=True)
class Principal:
    subject: str
    tenant_id: str
    name: str
    is_dev: bool = False


DEV_PRINCIPAL = Principal(subject="dev-user", tenant_id="dev", name="Local Dev", is_dev=True)


@lru_cache
def _jwk_client(tenant_id: str) -> PyJWKClient:
    return PyJWKClient(
        f"https://login.microsoftonline.com/{tenant_id}/discovery/v2.0/keys",
        cache_keys=True,
    )


def _validate_entra_token(token: str, settings: Settings) -> Principal:
    if not settings.entra_tenant_id or not settings.entra_client_id:
        raise HTTPException(
            status_code=500,
            detail="Server misconfigured: auth_mode=entra requires "
            "LEDGER_ENTRA_TENANT_ID and LEDGER_ENTRA_CLIENT_ID.",
        )
    try:
        signing_key = _jwk_client(settings.entra_tenant_id).get_signing_key_from_jwt(token)
        claims = jwt.decode(
            token,
            signing_key.key,
            algorithms=["RS256"],
            audience=[settings.entra_client_id, f"api://{settings.entra_client_id}"],
            options={"require": ["exp", "aud", "sub"]},
        )
    except jwt.PyJWTError as exc:
        raise HTTPException(status_code=401, detail=f"Invalid token: {exc}") from exc
    return Principal(
        subject=claims["sub"],
        tenant_id=claims.get("tid", ""),
        name=claims.get("name", claims.get("preferred_username", "")),
    )


def get_principal(
    request: Request, settings: Annotated[Settings, Depends(get_settings)]
) -> Principal:
    if settings.auth_mode == "dev":
        return DEV_PRINCIPAL

    auth = request.headers.get("Authorization", "")
    if not auth.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="Missing bearer token")
    return _validate_entra_token(auth.removeprefix("Bearer "), settings)


CurrentPrincipal = Annotated[Principal, Depends(get_principal)]
