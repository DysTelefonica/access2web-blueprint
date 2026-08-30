# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W61
# W61 (#524) — create_app use case.
"""Admin use case: create_app (W61, issue #524).

Registers a new application in the catalog. The use case delegates to
the driven port (``AppRepositoryPort.create``) so the application
layer stays a thin coordination shell — there is no domain invariant
beyond the dataclass-level validation (``App.__post_init__`` rejects
empty names / short codes).

``short_code`` uniqueness (D52) is enforced at the application layer
in a future slice — the adapter raises ``IntegrityError`` today if a
caller sends a duplicate, and the HTTP layer surfaces a 409 (DA-7).

The audit row emission is deferred to the W-TEST pattern's next
slice; the W61 routes call this use case through the
``functools.partial`` ``container.use_cases["create_app"]`` seam and
do not require an audit row on every call (the table-shape audit
emission belongs to the W62 hardening pass).
"""

from __future__ import annotations

from app.src.modules.lanzadera.domain.app import App, AppTopology
from app.src.modules.lanzadera.domain.ports.app_repository import (
    AppRepositoryPort,
)


async def create_app(
    name: str,
    short_code: str,
    deployment_topology: AppTopology,
    requires_office_presence: bool,
    *,
    apps: AppRepositoryPort,
    actor_id: object = None,
) -> App:
    """Register a new application in the catalog.

    Side effects:
        1. ``apps.create(...)`` persists the row and returns the
           DB-authoritative ``App`` (id + server-defaulted timestamps).

    Returns the freshly-inserted ``App``.

    Raises:
        ``ValueError`` — empty / whitespace-only ``name`` or ``short_code``
            (propagated from ``App.__post_init__``).
    """
    return await apps.create(
        name=name,
        short_code=short_code,
        deployment_topology=deployment_topology,
        requires_office_presence=requires_office_presence,
    )


__all__ = ["create_app"]
