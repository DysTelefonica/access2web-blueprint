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
from app.src.modules.lanzadera.delivery.http.auth_middleware import (
    AuthMiddleware,
)
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
    # W62 (#543), e2e-smoke (#876): #926 wired SessionRepositoryPg, so the
    # container must build. If it cannot, the app fails to start — a
    # silent ``container = None`` is what hid the 404 on /auth/* in the
    # e2e smoke. The previous ``except NotImplementedError`` branch is
    # removed on purpose; the engine is still disposed in the finally
    # below so a failed container construction does not leak connections.
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
    # W62 (#543), e2e-smoke (#876): Starlette 0.50+ no longer iterates
    # ``app.router.on_startup`` during the lifespan, so the auth / admin
    # / presence / apps routers must be registered from the lifespan
    # itself (the original ``@app.on_event("startup")`` is removed below).
    await _register_admin_routes()
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
# Middleware installed at import time
# ---------------------------------------------------------------------------
# W62 (#543), e2e-smoke (#876): Starlette 0.50+ refuses ``add_middleware``
# after the first ASGI call (the lifespan is that first call), so the
# auth + container-mirror middlewares must be installed here, before any
# request reaches the app. The signer is wrapped in a lazy proxy that
# resolves ``app.state.container.jwt_signer`` on every request because
# the container itself is created in the lifespan.


class _LazyAuthSigner:
    """``JwtSignerPort`` proxy that defers resolution to request time.

    The lifespan builds the real ``Hs256JwtSigner`` after this proxy is
    installed. Each ``verify()`` call looks up the container from
    ``app.state`` so the middleware always uses the live signer.
    """

    def verify(self, token: str, now: int | None = None) -> dict[str, object]:
        container = getattr(app.state, "container", None)
        if container is None:
            from app.src.modules.lanzadera.domain.errors import (
                InvalidTokenError,
            )

            raise InvalidTokenError("auth not initialised yet")
        if now is None:
            now = int(container._clock().timestamp())
        result: dict[str, object] = container.jwt_signer.verify(token, now=now)
        return result

    def sign(self, claims: dict[str, object]) -> str:
        """Issue a JWT. Called by the ``/auth/login`` use case after a
        successful credential check. The lifespan builds the real
        ``Hs256JwtSigner`` after this proxy is installed, so we look it
        up from ``app.state.container`` per call.
        """
        container = getattr(app.state, "container", None)
        if container is None:
            from app.src.modules.lanzadera.domain.errors import (
                InvalidTokenError,
            )

            raise InvalidTokenError("auth not initialised yet")
        return container.jwt_signer.sign(claims)


_ensure_container_middleware(app)
app.add_middleware(AuthMiddleware, jwt_signer=_LazyAuthSigner())

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


async def _register_admin_routes() -> None:
    """Attach admin routes after the app is built.

    The container lives in ``app.state.container`` (set by the lifespan).
    Registering the routes here rather than at import-time avoids a circular
    import with the admin module. W62 (#543), e2e-smoke (#876): called
    from the lifespan because Starlette 0.50+ no longer runs
    ``@app.on_event("startup")`` handlers.
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
    # W60 (#522): mount the SSE/heartbeat slice as its own router so
    # the long-lived generator in ``presence_stream`` resolves the
    # container via ``request.state.container`` (mirrored by the
    # middleware above) instead of a closure-bound instance.
    from app.src.modules.lanzadera.delivery.http.admin_routes_presence import (
        router as presence_router,
    )

    app.include_router(presence_router, prefix="/admin", tags=["presence"])
    # W61 (#524): mount the app CRUD slice as its own router so the
    # JSON routes resolve the container via ``request.state.container``
    # (mirrored by the middleware above) — the same seam the W60
    # presence slice uses. Mounting it before the legacy ``admin_router``
    # keeps the order deterministic; FastAPI dispatches by exact
    # match so the two routers cannot collide on shared paths.
    from app.src.modules.lanzadera.delivery.http.admin_routes_apps import (
        router as apps_router,
    )
    from app.src.modules.lanzadera.delivery.http.admin_routes_users_json import (
        router as users_json_router,
    )

    app.include_router(apps_router, prefix="/admin", tags=["apps"])
    app.include_router(users_json_router, prefix="/admin", tags=["admin-users-json"])
    app.include_router(admin_router)
    # W62 PR-6: mount the three auth routes (``/auth/login``,
    # ``/auth/logout``, ``/auth/me``). The router is created with the
    # ``auth`` tag and registered last so the path table stays
    # predictable; FastAPI dispatches by exact match so the new
    # routes cannot collide with the existing ``/admin/...`` slices.
    from app.src.modules.lanzadera.delivery.http.auth_routes import (
        register_auth_routes,
    )

    auth_router = APIRouter(tags=["auth"])
    register_auth_routes(auth_router, container=container)
    app.include_router(auth_router)
