# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 — lanzadera-mvp PR 2
# DA-1, D91 — Admin HTTP delivery placeholder gate (W62 PR-6 replaces it).
"""Admin HTTP delivery — D91 / W62 PR-6 placeholder gate (DL2, issue #55).

W58 (#515) wired the HTTP routes (``admin_routes.py``,
``admin_routes_users.py``, ``admin_routes_misc.py``) to the
``LanzaderaContainer`` directly. The router is now mounted by
``main.py`` through ``register_routes(router, templates=..., container=...)``
during application startup.

W62 PR-6 replaces the D91 placeholder ``require_global_admin`` with a
real implementation that:
- Reads ``request.state.user_id`` (populated by ``AuthMiddleware``,
  PR-5).
- Raises HTTP 401 when the state is empty.
- Calls ``container.global_admin_repo.is_global_admin(user_id)``.
- Raises HTTP 403 when the caller is not a global admin.

The module keeps the ``admin.require_global_admin`` attribute so the
existing ``auth_bypass`` test fixture (PR-7 retires it) can still
monkeypatch the function for the integration tests that need to skip
the gate in the transition.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from fastapi import HTTPException, status
from starlette.requests import Request

if TYPE_CHECKING:
    pass


__all__ = ["require_global_admin", "require_capability"]


async def require_global_admin(request: Request) -> None:
    """Verify the caller is a global admin (W62 PR-6).

    AD-W62-2 silent-failure contract: the middleware leaves
    ``request.state.user_id`` empty when the token is missing or invalid.
    This gate is where the empty state becomes an HTTP response: 401.

    Steps:
        1. Read ``request.state.user_id`` (set by ``AuthMiddleware``).
        2. If ``None``, raise HTTP 401.
        3. Resolve the container via ``request.app.state.container``.
        4. Call ``container.global_admin_repo.is_global_admin(user_id)``.
        5. If ``False``, raise HTTP 403.
    """
    user_id = getattr(request.state, "user_id", None)
    if user_id is None:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="authentication required",
        )
    container = getattr(request.app.state, "container", None)
    if container is None:
        # The container is not initialised (e.g. ``DATABASE_URL`` not
        # configured). The lifespan logs a warning; the route cannot
        # gate and rejects with 503 to avoid a silent permission bypass.
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="admin gate unavailable",
        )

    is_admin = await container.global_admin_repo.is_global_admin(user_id)
    if not is_admin:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="admin privileges required",
        )


async def require_capability(
    request: Request,
    *,
    capability: str,
) -> None:
    """Verify the caller holds ``capability`` for the app on ``request.app.state.app_id``.

    Usage in a route::

        from fastapi import Depends
        from app.src.modules.lanzadera.delivery.http.admin import require_capability

        @router.get("/apps/{app_id}/records",
                    dependencies=[Depends(require_capability(capability="write"))])
        async def list_records(request: Request) -> ...:

    The admin router sets ``request.app.state.app_id`` on each mounted
    sub-router so the capability gate knows which app the caller operates on.

    Raises HTTP 401 when no token is present. Raises HTTP 403 when the
    caller does not hold ``capability`` for the target app. Raises HTTP 503
    when the container is not available.
    """
    user_id = getattr(request.state, "user_id", None)
    if user_id is None:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="authentication required",
        )

    container = getattr(request.app.state, "container", None)
    if container is None:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="container not available",
        )

    app_id = getattr(request.app.state, "app_id", None)
    if app_id is None:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="app_id not set on request state",
        )

    caps = await container.use_cases["get_my_capabilities_for_app"](
        user_id=user_id, app_id=app_id
    )
    if capability not in caps:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail=f"capability '{capability}' required",
        )
