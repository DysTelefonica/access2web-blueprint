# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W60
# W60 (#522) — in-memory presence fake for the application/delivery tests.
"""In-memory ``PresenceRepository`` fake for the W60 (#522) tests.

The fake mirrors the production ``PresenceRepository`` Protocol surface
(``track`` / ``heartbeat`` / ``disconnect`` / ``list_connected``) so the
application and delivery tests run without spinning up a real Postgres.

The fake stores ``ConnectedUser`` rows in a dict keyed by ``user_id``;
``track`` is the upsert path (existing rows are overwritten in place so
``last_seen`` and ``email`` track the most recent call) and ``heartbeat``
bumps only ``last_seen`` without touching the email. Both methods
record their invocation in the ``*_calls`` lists so the use-case tests
can pin the exact mutation count.
"""

from __future__ import annotations

import dataclasses
from collections.abc import Sequence
from dataclasses import dataclass, field
from datetime import UTC, datetime
from typing import TYPE_CHECKING
from uuid import UUID, uuid4

# Imported at module level so dataclass field annotations resolve at
# runtime (no forward-reference gymnastics the type checker has to chase).
from app.src.modules.lanzadera.domain.presence import ConnectedUser

if TYPE_CHECKING:
    pass  # runtime imports above already cover the static checker


def _now() -> datetime:
    """Return a fresh UTC timestamp. Default clock for the fake's append/heartbeat paths."""
    return datetime.now(UTC)


@dataclass
class FakePresenceRepository:
    """In-memory ``PresenceRepository`` for the W60 application + delivery tests.

    Three call-recording lists (``track_calls`` / ``heartbeat_calls`` /
    ``disconnect_calls``) feed the use-case tests; ``list_connected``
    returns the rows ordered by ``last_seen`` DESC (the same ordering the
    Postgres adapter pins) so the SSE shape is exercised correctly.
    """

    # user_id -> ConnectedUser
    rows: dict[UUID, ConnectedUser] = field(default_factory=dict)
    track_calls: list[tuple[UUID, str]] = field(default_factory=list)
    heartbeat_calls: list[UUID] = field(default_factory=list)
    disconnect_calls: list[UUID] = field(default_factory=list)
    list_connected_calls: int = 0

    def add(self, user: ConnectedUser) -> None:
        """Synchronous seed helper for tests (test-side convenience, not in the Protocol)."""
        self.rows[user.user_id] = user

    async def track(self, user_id: UUID, email: str) -> None:
        self.track_calls.append((user_id, email))
        now = _now()
        existing = self.rows.get(user_id)
        if existing is None:
            self.rows[user_id] = ConnectedUser(
                user_id=user_id,
                email=email,
                connected_at=now,
                last_seen=now,
            )
            return
        # Upsert: keep ``connected_at`` on update, refresh ``email`` +
        # ``last_seen`` to match the SQL ``ON CONFLICT DO UPDATE`` semantics.
        self.rows[user_id] = dataclasses.replace(existing, email=email, last_seen=now)

    async def heartbeat(self, user_id: UUID) -> None:
        self.heartbeat_calls.append(user_id)
        existing = self.rows.get(user_id)
        if existing is None:
            # Adapter behaviour: no row, no-op. Replicating it here keeps
            # the contract symmetric end to end.
            return
        self.rows[user_id] = dataclasses.replace(existing, last_seen=_now())

    async def disconnect(self, user_id: UUID) -> None:
        self.disconnect_calls.append(user_id)
        # Idempotent: a disconnect call on a missing row is a no-op
        # (matches the Postgres DELETE WHERE which raises no error).
        self.rows.pop(user_id, None)

    async def list_connected(self, limit: int = 100) -> Sequence[ConnectedUser]:
        self.list_connected_calls += 1
        ordered = sorted(self.rows.values(), key=lambda row: row.last_seen, reverse=True)
        return ordered[:limit]

    @staticmethod
    def seed_row(
        *, user_id: UUID | None = None, email: str = "alice@enterprise.test"
    ) -> ConnectedUser:
        """Return a fresh ``ConnectedUser`` for test seeding helpers."""
        return ConnectedUser(
            user_id=user_id or uuid4(),
            email=email,
            connected_at=_now(),
            last_seen=_now(),
        )


__all__ = ["FakePresenceRepository"]
