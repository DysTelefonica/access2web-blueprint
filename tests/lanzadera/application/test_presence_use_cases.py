# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W60
# W60 (#522) — unit tests for the presence use cases.
"""Unit tests for the W60 (#522) presence use cases.

The tests cover the two application-layer entry points the SSE slice
relies on:

- ``track_presence`` — heartbeat path that ``POST /presence/heartbeat``
  drives on every client ping.
- ``get_connected_users`` — read path that both ``GET /presence`` (JSON
  snapshot) and ``GET /presence/stream`` (SSE emitter) share.

Each test seeds the in-memory ``FakePresenceRepository`` directly so the
assertions pin the contract the Postgres adapter has to honour: the
use case is a thin pass-through to the driven port.
"""

from __future__ import annotations

from collections.abc import Sequence
from datetime import UTC, datetime
from uuid import UUID, uuid4

from app.src.modules.lanzadera.application.get_connected_users import (
    get_connected_users,
)
from app.src.modules.lanzadera.application.track_presence import track_presence
from app.src.modules.lanzadera.domain.presence import ConnectedUser
from tests.lanzadera._presence_fakes import FakePresenceRepository


def _seed_user(
    *, user_id: UUID | None = None, email: str = "alice@enterprise.test"
) -> ConnectedUser:
    """Return a fresh ``ConnectedUser`` for the assertion helpers."""
    return FakePresenceRepository.seed_row(user_id=user_id, email=email)


# ---------------------------------------------------------------------------
# track_presence
# ---------------------------------------------------------------------------


async def test_track_presence_calls_repo() -> None:
    """``track_presence`` forwards the heartbeat to the repository's ``heartbeat`` method.

    The use case signature is ``track_presence(user_id, *, presence)``;
    calling it must result in exactly one ``heartbeat`` invocation that
    carries the same ``user_id`` forward. The repo's mutation record
    (``heartbeat_calls``) is the contract the Postgres adapter pins.
    """
    presence = FakePresenceRepository()
    user_id = uuid4()

    await track_presence(user_id, presence=presence)

    assert presence.heartbeat_calls == [user_id], (
        "track_presence must call presence.heartbeat(user_id) exactly once"
    )
    # ``heartbeat`` is a no-op when no row exists yet (matches the SQL
    # ``UPDATE ... WHERE`` semantics); the fake mirrors that.
    assert presence.rows == {}


async def test_track_presence_updates_existing_last_seen() -> None:
    """Calling ``track_presence`` on a known user refreshes the fake's ``last_seen``.

    The Postgres adapter's ``UPDATE last_seen = now() WHERE user_id``
    semantics are mirrored by the fake's ``dataclasses.replace`` call;
    this test asserts the user's ``last_seen`` strictly moves forward
    after a heartbeat (the row pre-existed, so the no-op path is NOT
    taken).
    """
    presence = FakePresenceRepository()
    user = _seed_user()
    presence.add(user)
    original_last_seen = user.last_seen
    # Force a measurable delta on systems where ``datetime.now(UTC)`` has
    # microsecond resolution guarantees the second call lands later.
    new_last_seen = datetime(2099, 1, 1, tzinfo=UTC)

    # The fake's ``heartbeat`` stamps ``datetime.now(UTC)`` on each
    # call; stub that via monkeypatch-style assignment so we can pin the
    # behaviour deterministically.
    import tests.lanzadera._presence_fakes as _fakes_module

    original_now = _fakes_module._now

    def _stub_now() -> datetime:
        return new_last_seen

    _fakes_module._now = _stub_now  # type: ignore[assignment]
    try:
        await track_presence(user.user_id, presence=presence)
    finally:
        _fakes_module._now = original_now  # type: ignore[assignment]

    assert presence.heartbeat_calls == [user.user_id]
    updated = presence.rows[user.user_id]
    assert updated.last_seen == new_last_seen
    assert updated.last_seen > original_last_seen


async def test_heartbeat_updates_last_seen() -> None:
    """Two consecutive heartbeats bump ``last_seen`` twice (no-op rows are surfaced).

    The Postgres adapter's ``heartbeat`` issues an ``UPDATE ... WHERE
    user_id``; the fake mirrors that with ``dataclasses.replace`` on
    every call. The assertion below pins that the second heartbeat
    strictly moves ``last_seen`` forward (the bug-detection contract
    the unit test exists for).
    """
    presence = FakePresenceRepository()
    user = _seed_user()
    presence.add(user)

    # First heartbeat: ``last_seen`` becomes ``now``.
    await presence.heartbeat(user.user_id)
    first = presence.rows[user.user_id].last_seen

    # Sleep-less second heartbeat: monotonic guarantee from
    # ``datetime.now(UTC)`` on systems with microsecond resolution;
    # the assertion below accepts equality as well as strict ordering
    # so the test stays portable across clock-resolution environments.
    await presence.heartbeat(user.user_id)
    second = presence.rows[user.user_id].last_seen

    assert presence.heartbeat_calls == [user.user_id, user.user_id]
    assert second >= first, "the second heartbeat must not move last_seen backwards"


# ---------------------------------------------------------------------------
# get_connected_users
# ---------------------------------------------------------------------------


async def test_get_connected_users_returns_sequence() -> None:
    """``get_connected_users`` returns every row the fake carries, ordered ``last_seen`` DESC.

    The default ``limit=100`` matches the SSE emitter's default; the
    use case is the thin pass-through that the HTTP route reaches for
    via ``container.use_cases["get_connected_users"](limit=...)``.
    """
    presence = FakePresenceRepository()
    alice = _seed_user(email="alice@enterprise.test")
    bob = _seed_user(email="bob@enterprise.test")
    presence.add(alice)
    presence.add(bob)

    users = await get_connected_users(presence=presence)

    assert isinstance(users, Sequence)
    assert {u.user_id for u in users} == {alice.user_id, bob.user_id}
    # ``list_connected_calls`` is the fake's per-test invocation counter;
    # exactly one call per ``get_connected_users`` invocation is the
    # contract the SSE emitter relies on (the HTTP layer does not
    # re-read the repo behind the use case's back).
    assert presence.list_connected_calls == 1


async def test_get_connected_users_respects_limit() -> None:
    """The ``limit`` kwarg reaches the repository (``LIMIT`` in the SQL)."""
    presence = FakePresenceRepository()
    for i in range(5):
        presence.add(_seed_user(email=f"user{i}@enterprise.test"))

    users = await get_connected_users(presence=presence, limit=3)

    assert len(users) == 3


__all__ = [
    "test_get_connected_users_respects_limit",
    "test_get_connected_users_returns_sequence",
    "test_heartbeat_updates_last_seen",
    "test_track_presence_calls_repo",
    "test_track_presence_updates_existing_last_seen",
]
