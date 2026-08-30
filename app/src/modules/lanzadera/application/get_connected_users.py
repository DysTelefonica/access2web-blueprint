# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W60
# W60 (#522) — get_connected_users use case.
"""Delivery use case: get_connected_users (W60, #522).

Returns the most-recently-active connected users, ordered by
``last_seen`` DESC. Used by two routes:

- ``GET /presence`` — one-shot JSON snapshot for the admin console.
- ``GET /presence/stream`` — SSE emitter that re-broadcasts the snapshot
  every 5 seconds.

The use case delegates the read straight to the driven port with a
default ``limit=100``. Routes that need a different cap pass the value
through their query parameters; the use case accepts a keyword-only
``limit`` so the contract is explicit at the call site.
"""

from __future__ import annotations

from collections.abc import Sequence

from app.src.modules.lanzadera.domain.ports.presence_repository import (
    PresenceRepository,
)
from app.src.modules.lanzadera.domain.presence import ConnectedUser


async def get_connected_users(
    *,
    presence: PresenceRepository,
    limit: int = 100,
) -> Sequence[ConnectedUser]:
    """Return the currently-connected users, ordered by ``last_seen`` DESC.

    ``limit`` is the cap on result size (default 100). The SSE emitter
    keeps this default because the snapshot is small relative to a
    typical admin-team size; a future slice can parameterise the route
    to forward a smaller cap when the client requests it explicitly.
    """
    return await presence.list_connected(limit=limit)


__all__ = ["get_connected_users"]
