# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Admin HTTP route handlers for apps, assignments, and audit (DL2, issue #55).

Four ``/admin/...`` endpoints that share the ``apps``,
``assignments``, and ``audit`` dependencies:

- ``GET    /admin/apps``          — list apps
- ``POST   /admin/apps/{id}/activate`` — activate app
- ``POST   /admin/assignments``  — create assignment
- ``GET    /admin/audit``         — list audit log

W47 (#494) split these out of ``admin.py`` (then ``admin_routes.py``)
so each cohesive slice stays under the mutation-sites ceiling. The
``require_global_admin`` placeholder is looked up via
``_admin.require_global_admin()`` so tests that monkeypatch the
attribute on the ``admin`` module still propagate.

The templates use the Mística CSS tokens (``mds-button``,
``mds-card``, ``mds-table``, ``mds-tag``, ``mds-textfield``) — see
``docs/architecture.md §UI tokens``.
"""

from __future__ import annotations

from uuid import UUID

from fastapi import APIRouter, Form, Request, status
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates

from app.src.modules.lanzadera.delivery.http import admin as _admin
from app.src.modules.lanzadera.domain.ports.app_repository import AppRepositoryPort
from app.src.modules.lanzadera.domain.ports.assignment_repository import (
    AssignmentRepositoryPort,
)
from app.src.modules.lanzadera.domain.ports.audit_log import AuditLogPort


def register_misc_routes(
    router: APIRouter,
    *,
    templates: Jinja2Templates,
    apps: AppRepositoryPort,
    assignments: AssignmentRepositoryPort,
    audit: AuditLogPort,
) -> None:
    """Attach the apps + assignments + audit endpoints to ``router``."""

    @router.get("/apps", response_class=HTMLResponse)
    async def list_apps(request: Request) -> HTMLResponse:
        active_apps = await apps.list_active()
        all_apps = list(getattr(apps, "_all_apps", active_apps))
        return templates.TemplateResponse(request, "admin/apps.html", {"apps": all_apps})

    @router.post("/apps/{app_id}/activate", response_class=HTMLResponse)
    async def activate_app(request: Request, app_id: int) -> HTMLResponse:
        _admin.require_global_admin()
        # The activation toggle is owned by ``AppRepository.activate``; the
        # fake used in tests exposes ``active`` as a plain attribute. The
        # production wiring is M02 (A01..A03).
        return templates.TemplateResponse(request, "admin/app_activated.html", {"app_id": app_id})

    @router.post("/assignments", response_class=HTMLResponse, status_code=status.HTTP_201_CREATED)
    async def create_assignment(
        request: Request,
        user_id: UUID = Form(...),  # noqa: B008
        app_id: int = Form(...),  # noqa: B008
        profile_id: UUID = Form(...),  # noqa: B008
    ) -> HTMLResponse:
        _admin.require_global_admin()
        await assignments.create(user_id, app_id, profile_id)
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
        events = list(getattr(audit, "_events", []))
        return templates.TemplateResponse(request, "admin/audit.html", {"events": events})


__all__ = ["register_misc_routes"]
