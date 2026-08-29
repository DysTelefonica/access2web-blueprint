# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Router orchestrator for the admin HTTP endpoints (DL2, issue #55).

Thin aggregator that delegates the seven ``/admin/...`` routes to two
cohesive slices:

- ``admin_routes_users.py`` — list/create/disable users (3 routes).
- ``admin_routes_misc.py`` — apps + assignments + audit (4 routes).

W47 (#494) split the routes from ``admin.py``; W58 (#515) re-wired
both slices to accept the ``LanzaderaContainer`` instead of individual
repository arguments, so use-case invariants (audit append, encrypt,
status transitions) are enforced end-to-end.

``admin.py`` keeps ``build_router`` + the ``require_global_admin``
placeholder; this file is just the per-router attachment dispatch.
"""
from __future__ import annotations

from fastapi import APIRouter
from fastapi.templating import Jinja2Templates

from app.src.modules.lanzadera.delivery.http.admin_routes_misc import (
    register_misc_routes,
)
from app.src.modules.lanzadera.delivery.http.admin_routes_users import (
    register_user_routes,
)
from app.src.modules.lanzadera.di.container import LanzaderaContainer


def register_routes(
    router: APIRouter,
    *,
    templates: Jinja2Templates,
    container: LanzaderaContainer,
) -> None:
    """Attach the seven admin endpoints to ``router``.

    All routes are wired to the ``LanzaderaContainer`` so the use-case
    layer (W54, issue #43) is in the call chain.
    """
    register_user_routes(router, templates=templates, container=container)
    register_misc_routes(router, templates=templates, container=container)


__all__ = ["register_routes"]
