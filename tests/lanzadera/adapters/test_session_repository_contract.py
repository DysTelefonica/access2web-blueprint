# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W62 (#537)
"""Contract test for ``SessionRepositoryPort`` (W62, D-W62-2, DA-1).

Verifies the in-memory ``FakeSessionRepository`` satisfies the
``SessionRepositoryPort`` Protocol structurally (DA-1) and exercises
each method with positive and negative cases. A Postgres
``SessionRepositoryPg`` adapter will follow in a later slice; this
test is the safety net the adapter must satisfy.
"""

from __future__ import annotations

from datetime import UTC, datetime, timedelta
from uuid import uuid4

from app.src.modules.lanzadera.domain.ports.session_repository import (
    SessionRepositoryPort,
)
from app.src.modules.lanzadera.domain.session import Session
from tests.lanzadera._fakes import FakeSessionRepository

# ---------------------------------------------------------------------------
# Protocol conformance
# ---------------------------------------------------------------------------


def test_fake_session_repository_conforms_to_protocol() -> None:
    """The fake exposes every ``SessionRepositoryPort`` method on its instance."""
    # Mypy checks structural conformance statically. The runtime check
    # below is a belt-and-suspenders test in case mypy ever stops
    # inheriting async signatures.
    expected = {n for n in dir(SessionRepositoryPort) if not n.startswith("_")} - {
        "__init__",
        "__annotations__",
        "__subclasshook__",
    }
    actual = {n for n in dir(FakeSessionRepository) if not n.startswith("_")}
    missing = expected - actual
    assert not missing, f"FakeSessionRepository is missing Protocol methods: {missing}"


# ---------------------------------------------------------------------------
# Behaviour — create / get_by_id / revoke
# ---------------------------------------------------------------------------


async def test_create_stores_session_and_returns_it() -> None:
    fake = FakeSessionRepository()
    sid, user_id = uuid4(), uuid4()
    now = datetime.now(UTC)
    session = Session(
        id=sid,
        user_id=user_id,
        created_at=now,
        expires_at=now + timedelta(hours=1),
    )

    out = await fake.create(session)

    assert out is session
    assert fake.sessions[sid] is session
    assert fake.create_calls == [sid]


async def test_get_by_id_returns_none_for_missing_session() -> None:
    fake = FakeSessionRepository()
    assert await fake.get_by_id(uuid4()) is None


async def test_get_by_id_returns_session_regardless_of_expiry() -> None:
    """The repository does not silently filter expired rows (DA-11 contract).

    The auth middleware is responsible for the expiry check; the
    repository returns whatever is on disk so the audit consumer can
    still see the full lifecycle.
    """
    fake = FakeSessionRepository()
    sid, user_id = uuid4(), uuid4()
    past = datetime.now(UTC) - timedelta(hours=1)
    session = Session(id=sid, user_id=user_id, created_at=past, expires_at=past)
    await fake.create(session)

    assert await fake.get_by_id(sid) is session


async def test_revoke_shortens_expires_at() -> None:
    fake = FakeSessionRepository()
    sid, user_id = uuid4(), uuid4()
    now = datetime.now(UTC)
    session = Session(
        id=sid,
        user_id=user_id,
        created_at=now,
        expires_at=now + timedelta(hours=1),
    )
    await fake.create(session)
    past = now - timedelta(seconds=1)

    ok = await fake.revoke(sid, past)

    assert ok is True
    assert fake.sessions[sid].expires_at == past
    assert fake.revoke_calls == [sid]


async def test_revoke_returns_false_for_missing_session() -> None:
    fake = FakeSessionRepository()
    assert await fake.revoke(uuid4(), datetime.now(UTC)) is False
    assert fake.revoke_calls == []


async def test_revoke_is_idempotent_on_subsequent_calls() -> None:
    """Revoking twice does not extend a past expiry (DA-11: audit-safe).

    The state stays consistent: ``expires_at`` remains the earlier of
    the two timestamps; ``revoke_calls`` records both invocations so the
    test can pin the exact call count.
    """
    fake = FakeSessionRepository()
    sid, user_id = uuid4(), uuid4()
    now = datetime.now(UTC)
    session = Session(
        id=sid,
        user_id=user_id,
        created_at=now,
        expires_at=now + timedelta(hours=1),
    )
    await fake.create(session)
    first = now - timedelta(seconds=2)
    second = now - timedelta(seconds=1)

    await fake.revoke(sid, first)
    await fake.revoke(sid, second)

    assert fake.sessions[sid].expires_at == first
    assert fake.revoke_calls == [sid, sid]
