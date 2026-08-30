# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W60
# W60 (#522) — track_presence use case.
"""Admin / delivery use case: track_presence (W60, #522).

Records a heartbeat for the user. Invoked from
``POST /presence/heartbeat``; the SSE emitter at
``GET /presence/stream`` only reads via ``get_connected_users``.

The use case delegates straight to the driven port (``heartbeat``) so
the application layer stays a thin coordination shell — there is no
domain invariant to enforce beyond "first time you ping this user, the
row exists" (which the adapter's ``ON CONFLICT DO UPDATE`` semantics
handle implicitly). The route layer is free to skip the call when the
user is not authenticated; the use case itself does not enforce an
identity check (that belongs to the auth middleware slated for W62).
"""

from __future__ import annotations

from uuid import UUID

from app.src.modules.lanzadera.domain.ports.presence_repository import (
    PresenceRepository,
)


async def track_presence(
    user_id: UUID,
    *,
    presence: PresenceRepository,
) -> None:
    """Record a heartbeat for ``user_id``.

    The hot path used by the SSE client's heartbeat ping. Calls
    ``PresenceRepository.heartbeat`` so the row's ``last_seen`` is
    bumped to ``now()``. If the row does not exist (a heartbeat without
    a prior ``track``), the call is a silent no-op at the SQL level —
    the next ``track`` call from the same client will mint a row.
    """
    await presence.heartbeat(user_id)


__all__ = ["track_presence"]
