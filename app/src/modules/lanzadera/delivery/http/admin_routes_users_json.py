# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp Phase 2 E2E
# Issue #605 — Phase 2: Admin users JSON routes for E2E testing.
"""Admin JSON route handlers for the user-management endpoints.

Covers the HTTP surface needed by ``e2e/admin-users.spec.ts`` (Phase 2 E2E):

- ``GET  /admin/users``              — paginated list, JSON array
- ``POST /admin/users``             — create user, returns 201 + body
- ``GET  /admin/users/{id}``       — user detail
- ``PATCH /admin/users/{id}``       — update (status toggle)
- ``DELETE /admin/users/{id}/assignments/{app_id}`` — revoke all assignments for user+app

Each endpoint calls ``require_global_admin`` as the first statement (W62 PR-6).
The auth middleware populates ``request.state.user_id`` from the JWT; if absent,
the middleware leaves it ``None`` and ``require_global_admin`` raises 401.

This module lives alongside ``admin_routes_users.py`` (HTML Jinja2 routes)
and ``admin_routes_apps.py`` (JSON app CRUD). HTML and JSON routes coexist:
HTML serves the browser UI; JSON serves the E2E Playwright suite and API clients.
"""

from __future__ import annotations

from uuid import UUID

from fastapi import APIRouter, HTTPException, Request, status

from app.src.modules.lanzadera.delivery.http import admin as _admin
from app.src.modules.lanzadera.domain.user import UserStatus

router = APIRouter(prefix="/admin/users", tags=["admin-users-json"])


def _container(request: Request) -> object:
    """Return the container attached to the request (or 503 if absent).

    Mirrors the pattern from ``admin_routes_apps.py`` — returning ``object``
    avoids a delivery->di layer-gate import so the container reference is
    resolved at call time via ``request.state.container``.
    """
    container = getattr(request.state, "container", None)
    if container is None:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="container not ready",
        )
    return container


def _actor_id(request: Request) -> UUID | None:
    return getattr(request.state, "user_id", None)


def _get_user_or_404(container: object, user_id: UUID) -> object:
    """Fetch a user by ID; raise 404 if absent."""
    user = container.users.get_by_id(user_id)  # type: ignore[attr-defined]
    if user is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f"user not found: {user_id}",
        )
    return user


def _serialise_user(user: object) -> dict[str, object]:
    _status = getattr(user, "status", None)
    _created = getattr(user, "created_at", None)
    return {
        "id": str(getattr(user, "id", "")),
        "email": getattr(user, "email", ""),
        "name": getattr(user, "name", ""),
        # UserStatus has .value; the fallback (None) has no .value so guard
        "status": _status.value if hasattr(_status, "value") else str(_status),  # type: ignore[union-attr]
        # datetime has .isoformat(); None/str do not
        "created_at": _created.isoformat() if hasattr(_created, "isoformat") else str(_created),  # type: ignore[union-attr]
    }


@router.get("", status_code=status.HTTP_200_OK)
async def list_users(request: Request, limit: int = 50, offset: int = 0) -> dict[str, object]:
    """Paginated user list as JSON."""
    await _admin.require_global_admin(request)
    container = _container(request)
    users, total = await container.list_all_users(limit=limit, offset=offset)  # type: ignore[attr-defined]
    return {
        "users": [_serialise_user(u) for u in users],
        "total": total,
        "limit": limit,
        "offset": offset,
    }


@router.post("", status_code=status.HTTP_201_CREATED)
async def create_user(
    request: Request,
    *,
    email: str,
    name: str,
    national_id: str,
) -> dict[str, object]:
    """Create a new platform user.

    The user is created with ``status=password_reset_required``.
    A separate reset flow activates the account.
    """
    await _admin.require_global_admin(request)
    container = _container(request)
    actor = _actor_id(request)
    try:
        user = await container.create_user(  # type: ignore[attr-defined]
            email=email.strip().lower(),
            name=name.strip(),
            national_id=national_id.strip(),
            actor_id=actor,
        )
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=str(exc),
        ) from exc
    return _serialise_user(user)


@router.get("/{user_id}", status_code=status.HTTP_200_OK)
async def get_user(request: Request, user_id: UUID) -> dict[str, object]:
    """Return a single user by ID (404 if not found)."""
    await _admin.require_global_admin(request)
    container = _container(request)
    user = _get_user_or_404(container, user_id)
    return _serialise_user(user)


@router.patch("/{user_id}", status_code=status.HTTP_200_OK)
async def update_user(
    request: Request,
    user_id: UUID,
    *,
    active: bool | None = None,
    status_value: str | None = None,
) -> dict[str, object]:
    """Partial update: toggle ``active`` (maps to DISABLED/ACTIVE status)."""
    await _admin.require_global_admin(request)
    container = _container(request)
    _get_user_or_404(container, user_id)

    if active is not None:
        new_status = UserStatus.ACTIVE if active else UserStatus.DISABLED
        await container.users.update_status(user_id, new_status)  # type: ignore[attr-defined]

    if status_value is not None:
        await container.users.update_status(user_id, UserStatus(status_value))  # type: ignore[attr-defined]

    user = _get_user_or_404(container, user_id)
    return _serialise_user(user)


@router.delete("/{user_id}/assignments/{app_id}", status_code=status.HTTP_204_NO_CONTENT)
async def revoke_assignment(
    request: Request,
    user_id: UUID,
    app_id: int,
) -> None:
    """Revoke all assignments for ``user_id`` on ``app_id``.

    Maps to ``AssignmentRepositoryPort.revoke(user_id, app_id)`` which deletes
    all rows from ``user_app_assignments`` for that (user, app) pair.
    """
    await _admin.require_global_admin(request)
    container = _container(request)
    from datetime import UTC, datetime

    now = datetime.now(UTC)
    await container.assignment_repo.revoke(  # type: ignore[attr-defined]
        user_id=user_id,
        app_id=app_id,
        now=now,
    )
    return None


__all__ = ["router"]
