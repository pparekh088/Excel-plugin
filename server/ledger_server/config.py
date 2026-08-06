"""Server configuration.

All settings are overridable via environment variables with the LEDGER_ prefix,
e.g. LEDGER_AUTH_MODE=entra, LEDGER_REDIS_URL=redis://localhost:6379/0.
"""

from functools import lru_cache
from pathlib import Path
from typing import Literal

from pydantic_settings import BaseSettings, SettingsConfigDict

# Repo root is two levels up from this file (server/ledger_server/config.py).
_REPO_ROOT = Path(__file__).resolve().parent.parent.parent


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_prefix="LEDGER_", env_file=".env", extra="ignore")

    # dev: no auth required (local development only; the server refuses to start
    #      in dev mode if LEDGER_ENV=production).
    # entra: validate Entra ID (Azure AD) bearer tokens via JWKS.
    auth_mode: Literal["dev", "entra"] = "dev"
    env: Literal["development", "production"] = "development"

    # Entra ID settings (required when auth_mode == "entra").
    entra_tenant_id: str = ""
    entra_client_id: str = ""  # the API's app registration (audience)

    # Empty -> in-memory session store (single-process dev only).
    redis_url: str = ""
    session_ttl_seconds: int = 8 * 3600

    # Directory containing the generated JSON Schemas for the typed tool surface.
    schemas_dir: Path = _REPO_ROOT / "shared" / "schemas"

    # CORS origins for the add-in dev server / hosted add-in.
    cors_origins: list[str] = ["https://localhost:3000"]

    # Bound the per-session tool-call log kept in the session store.
    max_tool_calls_per_session: int = 500


@lru_cache
def get_settings() -> Settings:
    return Settings()
