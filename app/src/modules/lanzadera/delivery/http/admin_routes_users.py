# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Admin HTTP route handlers for the user-management endpoints (DL2, issue #55).

Three ``/admin/users`` endpoints wired to the ``LanzaderaContainer``:

- ``GET    /admin/users``        — list users
- ``POST   /admin/users``        — create user (container.create_user use case)
- ``PATCH  /admin/users/{id}/disable`` — disable user (container.disable_user use case)

W47 (#494) split these out of ``admin.py``; W58 (#515) wired them to the
``LanzaderaContainer`` instead of calling the UserRepository directly, so
the use-case invariants (audit append, encrypt, status) are enforced.

The templates use the Mística CSS tokens (``mds-button``,
``mds-card``, ``mds-table``, ``mds-tag``, ``mds-textfield``) — see
``docs/architecture.md §UI tokens``.
"""
from __future__ import annotations

from uuid import UUID

from fastapi import APIRouter, Form, HTTPException, Request, status
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates

from app.src.modules.lanzadera.delivery.http import admin as _admin
from app.src.modules.lanzadera.di.container import LanzaderaContainer


def register_user_routes(
    router: APIRouter,
    *,
    templates: Jinja2Templates,
    container: LanzaderaContainer,
) -> None:
    """Attach the user-management endpoints to ``router``."""

    @router.get("/users", response_class=HTMLResponse)
    async def list_users(request: Request) -> HTMLResponse:
        all_users = await container.list_all_users()
        return templates.TemplateResponse(request, "admin/users.html", {"users": list(all_users)})

    @router.post("/users", response_class=HTMLResponse, status_code=status.HTTP_201_CREATED)
    async def create_user(
        request: Request,
        email: str = Form(...),
        name: str = Form(...),
        dni: str = Form(...),
    ) -> HTMLResponse:
        _admin.require_global_admin()
        try:
            new_user = await container.create_user(
                email=email.strip().lower(),
                name=name,
                national_id=dni,
            )
        except Exception as exc:  # noqa: BLE001
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=str(exc),
            ) from exc
        return templates.TemplateResponse(
            request,
            "admin/user_created.html",
            {"user": new_user},
            status_code=status.HTTP_201_CREATED,
        )

    @router.patch("/users/{user_id}/disable", response_class=HTMLResponse)
    async def disable_user(request: Request, user_id: UUID) -> HTMLResponse:
        _admin.require_global_admin()
        await container.disable_user(user_id=user_id)
        return templates.TemplateResponse(request, "admin/user_disabled.html", {"user_id": user_id})


__all__ = ["register_user_routes"]
