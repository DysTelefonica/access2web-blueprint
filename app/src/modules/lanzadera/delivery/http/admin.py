# mypy: disable-error-code=unused-ignore
"""Admin HTTP delivery (DL2, issue #55).

FastAPI router that exposes the seven ``/admin/...`` endpoints called out
in issue #55. The factory pattern (rather than a module-level
``router = APIRouter()``) is required because FastAPI routers are
bound to dependency-injection functions; the use-case ports are
injected by the composition root at application startup, not at import
time.

W47 (#494) split the seven route handlers out to
``admin_routes.py``. This module now owns only the ``build_router``
factory and the ``require_global_admin`` placeholder (the destructive
commands look it up via ``_admin.require_global_admin()`` so tests
that monkeypatch the attribute on the ``admin`` module still
propagate).
"""

from __future__ import annotations

from fastapi import APIRouter
from fastapi.templating import Jinja2Templates

from app.src.modules.lanzadera.delivery.http.admin_ports import (
    AppRepositoryPort,
    AssignmentRepositoryPort,
    AuditLogPort,
    GlobalAdminRepositoryPort,
    UserRepository,
)
from app.src.modules.lanzadera.delivery.http.admin_routes import register_routes


def require_global_admin() -> None:
    """Placeholder for D91's global-admin gate.

    The real implementation checks ``admins.is_global_admin(current_user.id)``
    from a session or header. This WU ships the route shape; the
    gating comes with M02's auth wiring (A01..A03).
    """


def build_router(
    *,
    templates: Jinja2Templates,
    users: UserRepository,
    apps: AppRepositoryPort,
    assignments: AssignmentRepositoryPort,
    audit: AuditLogPort,
    admins: GlobalAdminRepositoryPort,
) -> APIRouter:
    """Build the admin router."""
    router = APIRouter(prefix="/admin", tags=["admin"])
    register_routes(
        router,
        templates=templates,
        users=users,
        apps=apps,
        assignments=assignments,
        audit=audit,
        admins=admins,
    )
    return router


__all__ = ["build_router", "require_global_admin"]
