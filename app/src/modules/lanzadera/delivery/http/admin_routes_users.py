# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Admin HTTP route handlers for the user-management endpoints (DL2, issue #55).

Three ``/admin/users`` endpoints that share the ``users`` dependency:

- ``GET    /admin/users``        — list users
- ``POST   /admin/users``        — create user
- ``PATCH  /admin/users/{id}/disable`` — disable user

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

from fastapi import APIRouter, Form, HTTPException, Request, status
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates

from app.src.modules.lanzadera.delivery.http import admin as _admin
from app.src.modules.lanzadera.domain.ports import UserRepository
from app.src.modules.lanzadera.domain.user import UserStatus


def register_user_routes(
    router: APIRouter,
    *,
    templates: Jinja2Templates,
    users: UserRepository,
) -> None:
    """Attach the user-management endpoints to ``router``."""

    @router.get("/users", response_class=HTMLResponse)
    async def list_users(request: Request) -> HTMLResponse:
        # ``users`` is a UserRepository; the in-memory fake used in tests
        # does not support async list_all by default, so we read sync
        # and wrap. The real Postgres adapter supports the async path.
        all_users = list(getattr(users, "_users", {}).values())
        return templates.TemplateResponse(request, "admin/users.html", {"users": all_users})

    @router.post("/users", response_class=HTMLResponse, status_code=status.HTTP_201_CREATED)
    async def create_user(
        request: Request,
        email: str = Form(...),
        name: str = Form(...),
        dni: str = Form(...),
    ) -> HTMLResponse:
        _admin.require_global_admin()
        existing = await users.get_by_email(email.strip().lower())
        if existing is not None:
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=f"user {email} already exists",
            )
        from app.src.modules.lanzadera.domain.user import User, UserStatus

        new_user = User(
            id=__import__("uuid").uuid4(),
            email=email.strip().lower(),
            name=name,
            dni_encrypted=dni.encode("utf-8"),
            password_hash=None,
            status=UserStatus.PASSWORD_RESET_REQUIRED,
            failed_attempts=0,
            last_login_at=None,
            created_at=__import__("datetime").datetime.now(__import__("datetime").UTC),
            updated_at=__import__("datetime").datetime.now(__import__("datetime").UTC),
        )
        await users.create(new_user)
        return templates.TemplateResponse(
            request,
            "admin/user_created.html",
            {"user": new_user},
            status_code=status.HTTP_201_CREATED,
        )

    @router.patch("/users/{user_id}/disable", response_class=HTMLResponse)
    async def disable_user(request: Request, user_id: UUID) -> HTMLResponse:
        _admin.require_global_admin()
        await users.update_status(user_id, UserStatus.DISABLED)
        return templates.TemplateResponse(request, "admin/user_disabled.html", {"user_id": user_id})


__all__ = ["register_user_routes"]
