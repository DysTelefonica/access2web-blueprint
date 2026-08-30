# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W60
# W60 (#522) — pure domain entity for the presence-tracking feature.
"""Connected-user presence record (lanzadera-mvp/presence).

W60 (#522) ships the real-time session-presence slice: every logged-in
user who heartbeats over ``/admin/presence/heartbeat`` writes a row in
``lanzadera.presence_sessions`` and the SSE stream at
``/admin/presence/stream`` polls the same table every 5 seconds and
re-broadcasts the active set as ``data: <json>\\n\\n``.

The domain entity is intentionally narrow: the SQLAlchemy table holds
four fields (``user_id``, ``email``, ``connected_at``, ``last_seen``);
the dataclass mirrors that shape so the adapter's row-to-domain mapper
is a straight ``ConnectedUser(...)`` constructor call. The ``last_seen``
field is mutable on purpose — every heartbeat bumps it, so the dataclass
stays ``@dataclass`` (not ``frozen=True``) like ``Session`` and
``AuditEvent``.
"""

from __future__ import annotations

from app.src.modules.lanzadera.domain._imports import (
    UUID,
    dataclass,
    datetime,
)


@dataclass
class ConnectedUser:
    """A currently-connected user (mutable entity).

    ``last_seen`` is bumped on every heartbeat from ``/presence/heartbeat``
    so the SSE emitter at ``/presence/stream`` can decide whether the
    row should still appear in the connected set. ``connected_at`` is
    set once at the first ``track`` call (the ``ON CONFLICT (user_id)
    DO UPDATE SET last_seen = now()`` semantics in the adapter leaves
    it untouched on conflict).

    The dataclass enforces no invariants beyond what the columns declare;
    the persistence layer is the source of truth for non-nullability.
    """

    user_id: UUID
    email: str
    connected_at: datetime
    last_seen: datetime


__all__ = ["ConnectedUser"]
