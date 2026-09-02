# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Admin HTTP route handlers for apps, assignments, and audit (DL2, issue #55).

Four ``/admin/...`` endpoints wired to the ``LanzaderaContainer``:

- ``GET    /admin/apps``          — list apps
- ``POST   /admin/apps/{id}/activate`` — activate app
- ``POST   /admin/assignments``  — create assignment (container.assign_profile use case)
- ``GET    /admin/audit``         — list audit log (container.audit_repo)

W47 (#494) split these out of ``admin.py``; W58 (#515) wired them to the
``LanzaderaContainer`` instead of calling repositories directly in some cases.

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
from app.src.modules.lanzadera.di.container import LanzaderaContainer


def register_misc_routes(
    router: APIRouter,
    *,
    templates: Jinja2Templates,
    container: LanzaderaContainer,
) -> None:
    """Attach the apps + assignments + audit endpoints to ``router``."""

    @router.get("/apps", response_class=HTMLResponse)
    async def list_apps(request: Request) -> HTMLResponse:
        all_apps = await container.app_repo.list_active()
        return templates.TemplateResponse(request, "admin/apps.html", {"apps": list(all_apps)})

    @router.post("/apps/{app_id}/activate", response_class=HTMLResponse)
    async def activate_app(request: Request, app_id: int) -> HTMLResponse:
        await _admin.require_global_admin(request)
        return templates.TemplateResponse(request, "admin/app_activated.html", {"app_id": app_id})

    @router.post("/assignments", response_class=HTMLResponse, status_code=status.HTTP_201_CREATED)
    async def create_assignment(
        request: Request,
        user_id: UUID = Form(...),  # noqa: B008
        app_id: int = Form(...),  # noqa: B008
        profile_id: UUID = Form(...),  # noqa: B008
    ) -> HTMLResponse:
        await _admin.require_global_admin(request)
        await container.assign_profile(
            user_id=user_id,
            app_id=app_id,
            profile_id=profile_id,
        )
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
        events = await container.audit_repo.list_recent()
        return templates.TemplateResponse(request, "admin/audit.html", {"events": list(events)})


__all__ = ["register_misc_routes"]
