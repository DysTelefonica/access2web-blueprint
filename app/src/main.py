"""FastAPI application composition root — Lanzadera module (W58, #515).

Wires the ``LanzaderaContainer`` (W55, #49) into the FastAPI application
so every admin route goes through the use-case layer (W54, #43). The
lifespan creates the container once at startup; ``app.state.container``
carries it through the request lifecycle.

The bootstrap step (``run_bootstrap``) runs idempotently at startup
(D91 + CA-F4): if no ``GLOBAL_ADMIN_EMAILS`` are configured the step is
a no-op, and a DB unreachable at startup is logged but does not block
the event loop (the readiness probe will surface the unavailability).

Design choices (D73, D91, DA-1):
- The container is built inside the lifespan context so the ``AsyncEngine``
  is created before any request hits a route.
- ``DATABASE_URL`` is required; the app refuses to start without it.
- ``FERNET_KEY`` or ``SECRET_KEY`` configures the DNI encryption key.
- The admin router is mounted under ``/admin`` with the container injected.
"""

from __future__ import annotations

import logging
import os
from contextlib import asynccontextmanager
from pathlib import Path
from typing import AsyncIterator

from fastapi import FastAPI, Request, status
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates

from app.src.modules.lanzadera.adapters.bootstrap.env_admin_source_adapter import (
    EnvAdminSourceAdapter,
)
from app.src.modules.lanzadera.adapters.crypto.credential_hasher_argon2id import (
    CredentialHasherArgon2id,
)
from app.src.modules.lanzadera.adapters.cross.secret_manager import EnvSecretManagerAdapter
from app.src.modules.lanzadera.adapters.persistence.async_session_factory import (
    async_session_factory,
)
from app.src.modules.lanzadera.delivery.http.admin_routes import register_routes
from app.src.modules.lanzadera.di.bootstrap import run_bootstrap
from app.src.modules.lanzadera.di.container import LanzaderaContainer

_logger = logging.getLogger(__name__)

# Resolve the templates directory relative to this file so the app works
# regardless of the current working directory.
_TEMPLATES_DIR = Path(__file__).parent / "modules" / "lanzadera" / "delivery" / "http" / "templates"


# ---------------------------------------------------------------------------
# Lifespan — container creation and teardown
# ---------------------------------------------------------------------------


@asynccontextmanager
async def lifespan(app: FastAPI) -> AsyncIterator[None]:
    db_url = os.environ.get("DATABASE_URL")
    if not db_url:
        raise RuntimeError("DATABASE_URL environment variable is required")
    _engine, session_factory = async_session_factory(db_url)
    secret_manager = EnvSecretManagerAdapter()
    password_hasher = CredentialHasherArgon2id()
    bootstrap_source = EnvAdminSourceAdapter()
    container = LanzaderaContainer(
        session_factory=session_factory,
        secret_manager=secret_manager,
        password_hasher=password_hasher,
        bootstrap_source=bootstrap_source,
    )
    try:
        created = await run_bootstrap(container)
        if created:
            _logger.info("bootstrap: %d global-admin rows created", created)
    except Exception as exc:  # noqa: BLE001
        # Bootstrap is idempotent. A DB unreachable at startup is a deployment
        # error, not a startup failure — the readiness probe will surface it.
        _logger.warning("bootstrap skipped at startup: %s", exc)
    app.state.container = container
    yield
    await _engine.dispose()


# ---------------------------------------------------------------------------
# App factory
# ---------------------------------------------------------------------------

app = FastAPI(
    lifespan=lifespan,
    title="Platform — Lanzadera MVP",
    version="0.1.0",
    description=(
        "Hexagonal Python platform for the Lanzadera admin module. "
        "W58 (#515) wired the LanzaderaContainer (W55, #49) to the "
        "HTTP delivery layer so every admin route runs through the "
        "use-case layer (W54, #43)."
    ),
)

# ---------------------------------------------------------------------------
# Healthcheck
# ---------------------------------------------------------------------------


@app.get("/health", status_code=status.HTTP_200_OK, tags=["health"])
async def healthcheck() -> dict[str, str]:
    return {"status": "ok", "phase": "1", "service": "lanzadera"}


# ---------------------------------------------------------------------------
# Admin routes — wired to the LanzaderaContainer
# ---------------------------------------------------------------------------

_admin_templates = Jinja2Templates(directory=str(_TEMPLATES_DIR))


@app.on_event("startup")  # type: ignore[unused-ignore]
async def _register_admin_routes() -> None:
    """Attach admin routes after the app is built.

    The container lives in ``app.state.container`` (set by the lifespan).
    Registering the routes here rather than at import-time avoids a circular
    import with the admin module.
    """
    # Import here to avoid circular imports at module-load time.
    from fastapi import APIRouter

    container: LanzaderaContainer = app.state.container  # type: ignore[unused-ignore]
    admin_router = APIRouter(prefix="/admin", tags=["admin"])
    register_routes(
        admin_router,
        templates=_admin_templates,
        container=container,
    )
    app.include_router(admin_router)
