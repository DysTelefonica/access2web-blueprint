# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W60
# W60 (#522) — presence port.
"""Presence port (driven) for the W60 (#522) real-time presence slice.

The HTTP delivery layer (``POST /presence/heartbeat``,
``GET /presence``, ``GET /presence/stream``) talks to a repository
through this Protocol so the use case stays database-agnostic and the
in-memory fakes in ``tests/lanzadera/application/_fakes.py`` can stand
in at the adapter seam without spinning up a real Postgres.

The four operations map directly to the SQL the Postgres adapter emits:

- ``track(user_id, email)`` — INSERT ... ON CONFLICT (user_id) DO UPDATE
  SET last_seen = now(). Upsert so re-connections don't have to clear the
  row first; the email stays the value the most recent ``track`` call passed.
- ``heartbeat(user_id)`` — UPDATE last_seen = now() WHERE user_id. The
  hot path; called on every ``POST /presence/heartbeat`` ping from the
  client. Doesn't touch ``email``.
- ``disconnect(user_id)`` — DELETE WHERE user_id. Reserved for future
  cleanup paths (a cron that drops rows whose ``last_seen`` is older
  than a heartbeat TTL); the HTTP layer does not call it yet.
- ``list_connected(limit)`` — SELECT ORDER BY last_seen DESC LIMIT.
  The SSE emitter's read path; the ordering keeps the freshest at the
  top of the list.

The ``ConnectedUser`` return type on ``list_connected`` is referenced
by string annotation (``from __future__ import annotations``) so the
domain module can stay free of a circular import — the real
dataclass lives in ``app.src.modules.lanzadera.domain.presence``.
"""

from __future__ import annotations

from app.src.modules.lanzadera.domain.ports._imports import (
    UUID,
    Protocol,
    Sequence,
)

# W60 (#522): the ``list_connected`` annotation references the
# ``ConnectedUser`` dataclass from the domain layer. The runtime
# import here keeps mypy --strict happy (string annotation alone
# would still trigger the ``name-defined`` rule when the class is
# walked by the type checker). No circular import because
# ``presence.py`` depends on this module only through ``domain/_imports``,
# not on the ``Protocol`` class itself.
from app.src.modules.lanzadera.domain.presence import ConnectedUser


class PresenceRepository(Protocol):
    """Persistence boundary for the ``presence_sessions`` table.

    All four methods are async to match the rest of the driven ports
    (DA-1) — the Postgres adapter uses async SQLAlchemy; the fakes mirror
    the surface so tests run end-to-end without an event loop.
    """

    async def track(self, user_id: UUID, email: str) -> None:
        """Upsert a presence row for ``user_id`` with the given ``email``.

        Called once per (re)connection. The implementation bumps
        ``last_seen`` if the row already exists (the
        ``ON CONFLICT DO UPDATE`` path); ``connected_at`` is set only on
        the first insert.
        """
        ...

    async def heartbeat(self, user_id: UUID) -> None:
        """Refresh ``last_seen`` for ``user_id`` (no-op if no row exists)."""
        ...

    async def disconnect(self, user_id: UUID) -> None:
        """Delete the presence row for ``user_id`` (idempotent)."""
        ...

    async def list_connected(self, limit: int = 100) -> Sequence[ConnectedUser]:
        """Return the most recently active users, ordered by ``last_seen`` DESC.

        ``limit`` caps the result; the SSE emitter requests ``limit=100``
        by default (the same default the use-case layer forwards when the
        caller does not pass an explicit cap).
        """
        ...


__all__ = ["PresenceRepository"]
