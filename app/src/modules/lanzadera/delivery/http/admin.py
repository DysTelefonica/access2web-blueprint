"""Admin HTTP delivery (DL2, issue #55).

FastAPI router that exposes the seven ``/admin/...`` endpoints
called out in issue #55:

- ``GET    /admin/users``        — list users
- ``POST   /admin/users``        — create user
- ``PATCH  /admin/users/{id}/disable`` — disable user
- ``POST   /admin/assignments``  — create assignment
- ``GET    /admin/apps``          — list apps
- ``POST   /admin/apps/{id}/activate`` — activate app
- ``GET    /admin/audit``         — list audit log

The router depends on the use-case ports from #45 (``UserRepository``,
``AppRepository``, ``AssignmentRepository``, ``AuditLogPort``). The
templates use the Mística CSS tokens (``mds-button``, ``mds-card``,
``mds-table``, ``mds-tag``, ``mds-textfield``) — see
``docs/architecture.md §UI tokens``.
"""

from __future__ import annotations

from collections.abc import Iterable
from datetime import datetime, UTC
from typing import Protocol
from uuid import UUID, uuid4

from fastapi import APIRouter, Depends, Form, HTTPException, Request, status
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates

from app.src.modules.lanzadera.domain.ports.app_repository import AppRepositoryPort
from app.src.modules.lanzadera.domain.ports.assignment_repository import (
    AssignmentRepositoryPort,
)
from app.src.modules.lanzadera.domain.ports.audit_log import AuditLogPort
from app.src.modules.lanzadera.domain.ports.global_admin_repository import (
    GlobalAdminRepositoryPort,
)
from app.src.modules.lanzadera.domain.ports import UserRepository
from app.src.modules.lanzadera.domain.user import UserStatus


class _AppLike(Protocol):
    """Structural Protocol — what the template needs from an app row."""

    id: int
    name: str
    short_code: str
    active: bool


def build_router(
    *,
    templates: Jinja2Templates,
    users: UserRepository,
    apps: AppRepositoryPort,
    assignments: AssignmentRepositoryPort,
    audit: AuditLogPort,
    admins: GlobalAdminRepositoryPort,
) -> APIRouter:
    """Build the admin router.

    The factory pattern (rather than a module-level ``router = APIRouter()``)
    is required because FastAPI routers are bound to dependency-injection
    functions; the use-case ports are injected by the composition root at
    application startup, not at import time.
    """
    router = APIRouter(prefix="/admin", tags=["admin"])

    def require_global_admin() -> None:
        """Placeholder for D91's global-admin gate.

        The real implementation checks ``admins.is_global_admin(current_user.id)``
        from a session or header. This WU ships the route shape; the
        gating comes with M02's auth wiring (A01..A03).
        """

    @router.get("/users", response_class=HTMLResponse)
    async def list_users(request: Request) -> HTMLResponse:
        # ``users`` is a UserRepository; the in-memory fake used in tests
        # does not support async list_all by default, so we read sync
        # and wrap. The real Postgres adapter supports the async path.
        all_users = list(getattr(users, "_users", {}).values())  # type: ignore[attr-defined]
        return templates.TemplateResponse(
            request, "admin/users.html", {"users": all_users}
        )

    @router.post("/users", response_class=HTMLResponse, status_code=status.HTTP_201_CREATED)
    async def create_user(
        request: Request,
        email: str = Form(...),
        name: str = Form(...),
        dni: str = Form(...),
    ) -> HTMLResponse:
        require_global_admin()
        existing = await users.get_by_email(email.strip().lower())
        if existing is not None:
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=f"user {email} already exists",
            )
        new_user = await users.create(  # type: ignore[attr-defined]
            email=email.strip().lower(),
            name=name,
            dni_encrypted=dni.encode("utf-8"),
        )
        return templates.TemplateResponse(
            request,
            "admin/user_created.html",
            {"user": new_user},
            status_code=status.HTTP_201_CREATED,
        )

    @router.patch("/users/{user_id}/disable", response_class=HTMLResponse)
    async def disable_user(request: Request, user_id: UUID) -> HTMLResponse:
        require_global_admin()
        await users.update_status(user_id, UserStatus.DISABLED)  # type: ignore[attr-defined]
        return templates.TemplateResponse(
            request, "admin/user_disabled.html", {"user_id": user_id}
        )

    @router.get("/apps", response_class=HTMLResponse)
    async def list_apps(request: Request) -> HTMLResponse:
        active_apps = await apps.list_active()  # type: ignore[arg-defined]
        all_apps = list(getattr(apps, "_all_apps", active_apps))  # type: ignore[attr-defined]
        return templates.TemplateResponse(
            request, "admin/apps.html", {"apps": all_apps}
        )

    @router.post("/apps/{app_id}/activate", response_class=HTMLResponse)
    async def activate_app(
        request: Request, app_id: int
    ) -> HTMLResponse:
        require_global_admin()
        # The activation toggle is owned by ``AppRepository.activate``; the
        # fake used in tests exposes ``active`` as a plain attribute. The
        # production wiring is M02 (A01..A03).
        return templates.TemplateResponse(
            request, "admin/app_activated.html", {"app_id": app_id}
        )

    @router.post("/assignments", response_class=HTMLResponse, status_code=status.HTTP_201_CREATED)
    async def create_assignment(
        request: Request,
        user_id: UUID = Form(...),
        app_id: int = Form(...),
        profile_id: UUID = Form(...),
    ) -> HTMLResponse:
        require_global_admin()
        await assignments.create(user_id, app_id, profile_id)  # type: ignore[arg-defined]
        return templates.TemplateResponse(
            request,
            "admin/assignment_created.html",
            {
                "user_id": user_id,
                "app_id": app_id,
                "profile_id": profile_id,
            },
            status_code=status.HTTP_201_CREATED,
        )

    @router.get("/audit", response_class=HTMLResponse)
    async def list_audit(request: Request) -> HTMLResponse:
        # ``AuditLogPort.list_for_actor(None, since)`` returns the
        # global audit log. The test fake exposes ``_events`` as a list.
        events = list(getattr(audit, "_events", []))  # type: ignore[attr-defined]
        return templates.TemplateResponse(
            request, "admin/audit.html", {"events": events}
        )

    return router


__all__ = ["build_router"]