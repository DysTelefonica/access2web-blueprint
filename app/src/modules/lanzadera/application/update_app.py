# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W61
# W61 (#524) — update_app use case.
"""Admin use case: update_app (W61, issue #524).

Applies a partial patch to a catalog row. ``None`` fields are skipped
(the column keeps its current value); the adapter's ``RETURNING`` clause
returns the post-update row so the caller never sees a stale read.

The audit row emission is deferred to the W-TEST pattern's next
slice — the W61 routes call this use case through the
``functools.partial`` ``container.use_cases["update_app"]`` seam and
do not require an audit row on every call (the table-shape audit
emission belongs to the W62 hardening pass).
"""

from __future__ import annotations

from app.src.modules.lanzadera.domain.app import App, AppTopology
from app.src.modules.lanzadera.domain.ports.app_repository import (
    AppRepositoryPort,
)


async def update_app(
    app_id: int,
    *,
    apps: AppRepositoryPort,
    name: str | None = None,
    deployment_topology: AppTopology | None = None,
    requires_office_presence: bool | None = None,
    actor_id: object = None,
) -> App:
    """Apply the partial patch to ``app_id`` and return the new row.

    Side effects:
        1. ``apps.update(...)`` writes the patch and returns the
           DB-authoritative ``App`` (with ``updated_at`` bumped).

    Returns the freshly-updated ``App``.

    Raises:
        ``RuntimeError`` — ``apps.update`` returned no row (the adapter
            translates the empty ``RETURNING`` set into this signal).
    """
    return await apps.update(
        app_id,
        name=name,
        deployment_topology=deployment_topology,
        requires_office_presence=requires_office_presence,
    )


__all__ = ["update_app"]
