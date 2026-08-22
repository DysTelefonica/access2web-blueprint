"""Contract-conformance test for :class:`ResetTokenRepositoryPort`.

Pin both implementations against the same Protocol:
- the in-memory ``FakeResetTokenRepository`` (used by the rest of
  the test suite), and
- the Postgres ``ResetTokenRepositoryPg`` wired to the
  :func:`async_session_factory` seam (DA-1, W05 #42-subset).
"""

from __future__ import annotations

import dataclasses
from datetime import UTC, datetime, timedelta
from uuid import uuid4

from app.src.modules.lanzadera.adapters.persistence import (
    ResetTokenRepositoryPg,
    async_session_factory,
)
from app.src.modules.lanzadera.domain.ports.reset_token_repository import (
    ResetTokenRepositoryPort,
)
from tests.lanzadera.adapters._fakes import FakeResetTokenRepository


def test_reset_token_pg_satisfies_protocol() -> None:
    """Static structural check on the Postgres adapter's method set."""
    expected = {
        "insert",
        "find_unused",
        "mark_consumed",
        "mark_superseded",
        "purge_expired",
    }
    actual = set(dir(ResetTokenRepositoryPg))
    missing = expected - actual
    assert not missing, f"ResetTokenRepositoryPg is missing Protocol methods: {missing}"


async def test_fake_insert_then_find_unused_returns_token() -> None:
    fake: ResetTokenRepositoryPort = FakeResetTokenRepository()
    user_id = uuid4()
    future = datetime.now(UTC) + timedelta(hours=24)

    token = await fake.insert(user_id, token_hash="hash-abc", expires_at=future)

    found = await fake.find_unused("hash-abc")
    assert found is not None
    assert found.id == token.id
    assert found.user_id == user_id
    assert found.consumed_at is None
    assert found.superseded_at is None


async def test_fake_find_unused_returns_none_for_unknown_hash() -> None:
    fake = FakeResetTokenRepository()
    assert await fake.find_unused("never-inserted") is None


async def test_fake_mark_consumed_then_find_unused_returns_none() -> None:
    """DA-4: consumed tokens are single-use."""
    fake = FakeResetTokenRepository()
    future = datetime.now(UTC) + timedelta(hours=24)
    await fake.insert(uuid4(), token_hash="hash-consumed", expires_at=future)
    await fake.mark_consumed("hash-consumed", at=datetime.now(UTC))
    assert await fake.find_unused("hash-consumed") is None


async def test_fake_mark_superseded_invalidates_other_tokens_for_user() -> None:
    """DA-4: a fresh ``issue_reset_token`` invalidates prior live tokens for the same user."""
    fake = FakeResetTokenRepository()
    user_id = uuid4()
    future = datetime.now(UTC) + timedelta(hours=24)

    await fake.insert(user_id, token_hash="hash-old", expires_at=future)
    await fake.insert(user_id, token_hash="hash-new", expires_at=future)
    await fake.mark_superseded(user_id, at=datetime.now(UTC))

    assert await fake.find_unused("hash-old") is None
    assert fake.by_hash["hash-old"].superseded_at is not None
    assert await fake.find_unused("hash-new") is None
    assert fake.by_hash["hash-new"].superseded_at is not None


async def test_fake_purge_expired_removes_only_expired() -> None:
    """The fake deletes rows whose ``expires_at < now`` and leaves the rest."""
    fake = FakeResetTokenRepository()
    user_id = uuid4()
    now = datetime.now(UTC)

    # Live token
    await fake.insert(user_id, token_hash="hash-live", expires_at=now + timedelta(hours=24))
    # Expired-by-backdate
    expired = await fake.insert(
        user_id,
        token_hash="hash-expired",
        expires_at=now + timedelta(seconds=10),
    )
    backdated = dataclasses.replace(
        expired,
        created_at=now - timedelta(hours=26),
        expires_at=now - timedelta(seconds=1),
    )
    fake.by_hash["hash-expired"] = backdated

    purged = await fake.purge_expired(now)
    assert purged == 1
    assert "hash-expired" not in fake.by_hash
    assert "hash-live" in fake.by_hash


async def test_reset_token_pg_construction() -> None:
    """``ResetTokenRepositoryPg`` takes the factory seam and is constructable (DA-1)."""
    _, factory = async_session_factory("postgresql+asyncpg://nobody:nopwd@127.0.0.1:1/nodb")
    adapter = ResetTokenRepositoryPg(factory)
    assert adapter is not None
