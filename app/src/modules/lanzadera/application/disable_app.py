# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W61
# W61 (#524) — disable_app use case.
"""Admin use case: disable_app (W61, issue #524).

Retires an application: the row stays in the table (the lifecycle
preserves history for audit reads), but ``registration_status`` flips
to ``retired`` and the row disappears from ``list_active`` (the hot
read path).

The audit row emission is deferred to the W-TEST pattern's next
slice — the W61 routes call this use case through the
``functools.partial`` ``container.use_cases["disable_app"]`` seam and
do not require an audit row on every call (the table-shape audit
emission belongs to the W62 hardening pass).
"""

from __future__ import annotations

from app.src.modules.lanzadera.domain.app import App
from app.src.modules.lanzadera.domain.ports.app_repository import (
    AppRepositoryPort,
)


async def disable_app(
    app_id: int,
    *,
    apps: AppRepositoryPort,
    actor_id: object = None,
) -> App:
    """Retire ``app_id`` (status=retired) and return the row.

    Side effects:
        1. ``apps.disable(...)`` flips ``registration_status`` to
           ``retired`` and stamps ``updated_at``.

    Returns the post-disable ``App`` so the caller can correlate the
    response with the row's authoritative timestamps.

    Raises:
        ``RuntimeError`` — ``apps.disable`` returned no row (the
            adapter translates the empty ``RETURNING`` set into this
            signal; the row was deleted between the route's read and
            write).
    """
    return await apps.disable(app_id)


__all__ = ["disable_app"]
