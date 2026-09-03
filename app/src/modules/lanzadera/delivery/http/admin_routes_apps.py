# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W62 PR-7
# W61 (#524) — admin JSON routes for the app catalog.
# W62 PR-7 (#544) — replaced the X-Admin-User-ID header stub with
# request.state.user_id populated by AuthMiddleware (PR-5).
"""Admin HTTP route handlers for the W61 (#524) app CRUD slice.

Four ``/admin/apps/...`` JSON endpoints:

- ``POST   /admin/apps``               — register a new catalog row.
- ``GET    /admin/apps/{app_id}``      — read the catalog row.
- ``PATCH  /admin/apps/{app_id}``      — apply a partial patch.
- ``DELETE /admin/apps/{app_id}``      — retire the row (status=retired).

The destructive routes (``POST``, ``PATCH``, ``DELETE``) record the
audit actor from ``request.state.user_id`` populated by
``AuthMiddleware`` (W62 PR-5). The ``_actor_id`` helper resolves
the value or returns ``None`` when the request carries no auth
session — the W61 contract continues to accept a missing actor
(rather than 401-ing) because ``require_global_admin`` (PR-6) is the
gate that enforces presence of a global admin on the destructive
routes; ``actor_id`` is the audit-trail field that flows into DA-11.

The router is mounted directly in ``app/src/main.py`` (W61 ships its
own ``/admin`` prefix), mirroring the W60 presence slice. The
container is resolved via ``request.state.container`` (set by the
FastAPI app's lifespan or by the small middleware that mirrors
``app.state.container`` onto every request) so the JSON routes do not
need a closure binding — every request picks up the current container
state and the test path stays free of any startup wiring.
"""

from __future__ import annotations

from uuid import UUID

from fastapi import APIRouter, HTTPException, Request, status

from app.src.modules.lanzadera.domain.app import (
    App,
    AppTopology,
)

router = APIRouter()


def _container(request: Request) -> object:
    """Return the container attached to the request (or 503 if absent).

    FastAPI's ``app.state.container`` is set by the lifespan; the
    ``request.state.container`` it forwards to is populated by the
    ``_ContainerMirrorMiddleware``. The W61 routes never reach for
    ``app.state.container`` directly so the container reference can be
    replaced mid-flight (the middleware reads ``app.state`` at request
    time).
    """
    container = getattr(request.state, "container", None)
    if container is None:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="app routes require the container (lifespan not ready)",
        )
    return container


def _actor_id(request: Request) -> UUID | None:
    """Return the authenticated user ID from the request state.

    W62 PR-7 (#544): the W61 ``X-Admin-User-ID`` header stub is
    replaced by ``request.state.user_id`` populated by
    ``AuthMiddleware`` (PR-5). When the request has no valid auth
    session (no Bearer token, expired token, or invalid signature
    under the AD-W62-2 silent-failure contract), the middleware
    leaves the state ``None`` and this helper propagates the
    ``None`` so the audit-actor field on the use case stays
    explicit. ``require_global_admin`` (PR-6) is the gate that
    turns an empty state into HTTP 401 before the destructive
    routes ever call ``_actor_id``.
    """
    return getattr(request.state, "user_id", None)


def _serialise(app: App) -> dict[str, object]:
    """Render an ``App`` as the JSON shape the W61 routes return.

    Keys are stable (``id`` / ``name`` / ``short_code`` /
    ``deployment_topology`` / ``requires_office_presence`` /
    ``registration_status``); values are JSON-safe — enum fields are
    stringified via ``.value`` so the JSON payload does not carry the
    StrEnum subclass (which ``json.dumps`` cannot serialise).
    """
    return {
        "id": app.id,
        "name": app.name,
        "short_code": app.short_code,
        "deployment_topology": app.deployment_topology.value,
        "requires_office_presence": app.requires_office_presence,
        "registration_status": app.registration_status.value,
    }


@router.post("/apps", status_code=status.HTTP_201_CREATED)
async def create_app(request: Request) -> dict[str, object]:
    """Register a new application in the catalog.

    Body shape (JSON, required):
        ``name``                     — non-empty string
        ``short_code``               — non-empty string
        ``deployment_topology``      — ``"central"`` or ``"office-nas"``
        ``requires_office_presence`` — optional bool (defaults to ``False``)
    """
    container = _container(request)
    payload = await request.json()
    name = payload.get("name")
    short_code = payload.get("short_code")
    topology_raw = payload.get("deployment_topology")
    requires_office = payload.get("requires_office_presence", False)
    if not isinstance(name, str) or not name.strip():
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="name must be a non-empty string",
        )
    if not isinstance(short_code, str) or not short_code.strip():
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="short_code must be a non-empty string",
        )
    try:
        topology = AppTopology(topology_raw)
    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=(
                f"deployment_topology must be one of "
                f"{[t.value for t in AppTopology]}; got {topology_raw!r}"
            ),
        ) from exc
    actor_id = _actor_id(request)
    use_cases = container.use_cases  # type: ignore[attr-defined]
    app = await use_cases["create_app"](
        name=name,
        short_code=short_code,
        deployment_topology=topology,
        requires_office_presence=bool(requires_office),
        actor_id=actor_id,
    )
    return _serialise(app)


@router.get("/apps/{app_id}")
async def get_app(request: Request, app_id: int) -> dict[str, object]:
    """Read the catalog row with ``app_id`` (404 if it does not exist)."""
    container = _container(request)
    app_repo = container.app_repo  # type: ignore[attr-defined]
    app = await app_repo.get_by_id(app_id)
    if app is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f"app not found for id={app_id}",
        )
    return _serialise(app)


@router.patch("/apps/{app_id}")
async def update_app(request: Request, app_id: int) -> dict[str, object]:
    """Apply a partial patch to the catalog row with ``app_id``.

    Body shape (JSON, all keys optional):
        ``name``                     — non-empty string
        ``deployment_topology``      — ``"central"`` or ``"office-nas"``
        ``requires_office_presence`` — bool
    """
    container = _container(request)
    payload = await request.json()
    name: str | None = payload.get("name")
    topology: AppTopology | None = None
    if "deployment_topology" in payload:
        try:
            topology = AppTopology(payload["deployment_topology"])
        except ValueError as exc:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=(
                    f"deployment_topology must be one of "
                    f"{[t.value for t in AppTopology]}; got {payload['deployment_topology']!r}"
                ),
            ) from exc
    requires_office: bool | None = payload.get("requires_office_presence")
    if name is not None and (not isinstance(name, str) or not name.strip()):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="name must be a non-empty string when present",
        )
    actor_id = _actor_id(request)
    use_cases = container.use_cases  # type: ignore[attr-defined]
    app = await use_cases["update_app"](
        app_id=app_id,
        name=name,
        deployment_topology=topology,
        requires_office_presence=requires_office,
        actor_id=actor_id,
    )
    return _serialise(app)


@router.delete("/apps/{app_id}", status_code=status.HTTP_204_NO_CONTENT)
async def disable_app(request: Request, app_id: int) -> None:
    """Retire the catalog row with ``app_id`` (status=retired)."""
    container = _container(request)
    _actor_id(request)
    use_cases = container.use_cases  # type: ignore[attr-defined]
    await use_cases["disable_app"](app_id=app_id)
    return None


__all__ = ["router"]
