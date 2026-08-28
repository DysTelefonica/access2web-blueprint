"""Lanzadera admin use cases (DA-1, D91).

Pure-function orchestrators over the domain ports. Each use case is a
``async def`` that takes its required ports as keyword-only dependencies
and returns either a domain entity or raises a domain error. The
``delivery`` (HTTP, CLI) layer calls these directly; the ``bootstrap``
adapter in ``di/`` calls them on process start.

W54 (#508) shipped the 9 use cases required by issue #43. W55 will
wire them into the composition root.

Writes flow through the user port, never via direct SQL — the test
suite uses the in-memory fakes in ``tests/lanzadera/application/_fakes.py``
to exercise the use cases end-to-end without the Postgres adapter.
"""

from app.src.modules.lanzadera.application.assign_profile import assign_profile
from app.src.modules.lanzadera.application.audit_append import audit_append
from app.src.modules.lanzadera.application.bootstrap_global_admins import (
    bootstrap_global_admins,
)
from app.src.modules.lanzadera.application.create_user import create_user
from app.src.modules.lanzadera.application.disable_user import disable_user
from app.src.modules.lanzadera.application.grant_global_admin import (
    grant_global_admin,
)
from app.src.modules.lanzadera.application.list_effective_apps import (
    list_effective_apps,
)
from app.src.modules.lanzadera.application.revoke_global_admin import (
    revoke_global_admin,
)
from app.src.modules.lanzadera.application.set_password import set_password

__all__ = [
    "assign_profile",
    "audit_append",
    "bootstrap_global_admins",
    "create_user",
    "disable_user",
    "grant_global_admin",
    "list_effective_apps",
    "revoke_global_admin",
    "set_password",
]
