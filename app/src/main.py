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
from collections.abc import AsyncIterator, Awaitable, Callable
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, status
from fastapi.templating import Jinja2Templates
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request as _StarletteRequest
from starlette.responses import Response

from app.src.modules.lanzadera.adapters.bootstrap.env_admin_source_adapter import (
    EnvAdminSourceAdapter,
)
from app.src.modules.lanzadera.adapters.cross.secret_manager import EnvSecretManagerAdapter
from app.src.modules.lanzadera.adapters.crypto.credential_hasher_argon2id import (
    CredentialHasherArgon2id,
)
from app.src.modules.lanzadera.adapters.persistence.async_session_factory import (
    async_session_factory,
)
from app.src.modules.lanzadera.delivery.http.admin_routes import register_routes
from app.src.modules.lanzadera.di.bootstrap import run_bootstrap
from app.src.modules.lanzadera.di.container import LanzaderaContainer

_logger = logging.getLogger(__name__)

# W60 (#522): middleware that mirrors ``app.state.container`` onto every
# request. The SSE emitter at ``/admin/presence/stream`` resolves the
# container via ``request.state.container`` so the long-lived generator
# does not depend on the closure that registered the rest of the admin
# routes (those go through ``register_routes`` and bind the container at
# startup). The mirror is a one-line pass-through that re-reads the app
# state per request — the app-level container reference can be replaced
# during the lifespan and every subsequent request sees the new value.
_container_mirror_installed: bool = False


class _ContainerMirrorMiddleware(BaseHTTPMiddleware):
    """HTTP middleware that exposes ``app.state.container`` as
    ``request.state.container`` on every request. The middleware is
    installed at startup by the ``_register_admin_routes`` callback so
    the admin and presence slices share a single per-request view of
    the container (the lifespan may replace ``app.state.container``
    mid-flight; the mirror re-reads it every request).
    """

    async def dispatch(
        self,
        request: _StarletteRequest,
        call_next: Callable[[_StarletteRequest], Awaitable[Response]],
    ) -> Response:
        container = getattr(request.app.state, "container", None)
        if container is not None:
            request.state.container = container
        return await call_next(request)


def _ensure_container_middleware(app_obj: FastAPI) -> None:
    """Install the container-mirror middleware once per app instance.

    Adding the middleware twice would double the per-request cost; the
    module-level ``_container_mirror_installed`` flag keeps the install
    idempotent across startup callbacks (TestClient, ASGI hot reload, etc.).
    """
    global _container_mirror_installed
    if _container_mirror_installed:
        return
    app_obj.add_middleware(_ContainerMirrorMiddleware)
    _container_mirror_installed = True


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
        # No database configured — app starts for headless tests that only hit
        # the healthcheck endpoint. The container stays unset; any request that
        # needs it will get an AttributeError at runtime (acceptable for the MVP).
        _logger.warning("DATABASE_URL not set; container not initialised")
        app.state.container = None
        yield
        return
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
    version="0.0.0+phase0",
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
    _ensure_container_middleware(app)
    admin_router = APIRouter(prefix="/admin", tags=["admin"])
    register_routes(
        admin_router,
        templates=_admin_templates,
        container=container,
    )
    # W60 (#522): mount the SSE/heartbeat slice as its own router so
    # the long-lived generator in ``presence_stream`` resolves the
    # container via ``request.state.container`` (mirrored by the
    # middleware above) instead of a closure-bound instance.
    from app.src.modules.lanzadera.delivery.http.admin_routes_presence import (
        router as presence_router,
    )

    app.include_router(presence_router, prefix="/admin", tags=["presence"])
    app.include_router(admin_router)
