# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Router orchestrator for the admin HTTP endpoints (DL2, issue #55).

Thin aggregator that delegates the seven ``/admin/...`` routes to two
cohesive slices:

- ``admin_routes_users.py`` — list/create/disable users (3 routes).
- ``admin_routes_misc.py`` — apps + assignments + audit (4 routes).

W47 (#494) extracted the routes from ``admin.py`` so each file stays
under the mutation-sites ceiling. ``admin.py`` keeps ``build_router`` +
the ``require_global_admin`` placeholder; this file is just the
per-router attachment dispatch.
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
from app.src.modules.lanzadera.delivery.http.admin_routes_misc import (
    register_misc_routes,
)
from app.src.modules.lanzadera.delivery.http.admin_routes_users import (
    register_user_routes,
)


def register_routes(
    router: APIRouter,
    *,
    templates: Jinja2Templates,
    users: UserRepository,
    apps: AppRepositoryPort,
    assignments: AssignmentRepositoryPort,
    audit: AuditLogPort,
    admins: GlobalAdminRepositoryPort,
) -> None:
    """Attach the seven admin endpoints to ``router``.

    ``admins`` is accepted for symmetry with ``build_router`` — the
    destructive commands look it up indirectly via the
    ``_admin.require_global_admin`` placeholder.
    """
    register_user_routes(router, templates=templates, users=users)
    register_misc_routes(
        router,
        templates=templates,
        apps=apps,
        assignments=assignments,
        audit=audit,
    )


__all__ = ["register_routes"]
