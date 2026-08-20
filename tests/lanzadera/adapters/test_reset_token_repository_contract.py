"""Contract-conformance test for :class:`ResetTokenRepositoryPort`.

Async I/O uses ``asyncio.run`` (the project's ``pyproject.toml`` does
not enable ``pytest-asyncio``).
"""

from __future__ import annotations

import asyncio
import dataclasses
from datetime import UTC, datetime, timedelta
from uuid import uuid4

from app.src.modules.lanzadera.adapters.repos import ResetTokenRepositoryPg
from app.src.modules.lanzadera.domain.ports.reset_token_repository import (
    ResetTokenRepositoryPort,
)
from tests.lanzadera.adapters._fakes import FakeResetTokenRepository


def test_reset_token_pg_satisfies_protocol() -> None:
    adapter: ResetTokenRepositoryPort = ResetTokenRepositoryPg()
    for method in ("insert", "find_unused", "mark_consumed", "mark_superseded", "purge_expired"):
        assert hasattr(adapter, method), f"missing {method}"


def test_fake_insert_then_find_unused_returns_token() -> None:
    fake = FakeResetTokenRepository()
    user_id = uuid4()
    future = datetime.now(UTC) + timedelta(hours=24)

    async def _go() -> None:
        token = await fake.insert(user_id, token_hash="hash-abc", expires_at=future)

        found = await fake.find_unused("hash-abc")
        assert found is not None
        assert found.id == token.id
        assert found.user_id == user_id
        assert found.consumed_at is None
        assert found.superseded_at is None

    asyncio.run(_go())


def test_fake_find_unused_returns_none_for_unknown_hash() -> None:
    fake = FakeResetTokenRepository()

    async def _go() -> None:
        found = await fake.find_unused("never-inserted")
        assert found is None

    asyncio.run(_go())


def test_fake_mark_consumed_then_find_unused_returns_none() -> None:
    """DA-4: consumed tokens are single-use."""
    fake = FakeResetTokenRepository()
    future = datetime.now(UTC) + timedelta(hours=24)

    async def _go() -> None:
        await fake.insert(uuid4(), token_hash="hash-consumed", expires_at=future)
        await fake.mark_consumed("hash-consumed", at=datetime.now(UTC))

        found = await fake.find_unused("hash-consumed")
        assert found is None

    asyncio.run(_go())


def test_fake_mark_superseded_invalidates_other_tokens_for_user() -> None:
    """DA-4: a fresh ``issue_reset_token`` invalidates prior live tokens for the same user."""
    fake = FakeResetTokenRepository()
    user_id = uuid4()
    future = datetime.now(UTC) + timedelta(hours=24)

    async def _go() -> None:
        await fake.insert(user_id, token_hash="hash-old", expires_at=future)
        await fake.insert(user_id, token_hash="hash-new", expires_at=future)
        await fake.mark_superseded(user_id, at=datetime.now(UTC))

        # ``find_unused`` returns None for both — superseded tokens are not
        # findable. The fake replaces the frozen ``ResetToken`` instance with a
        # new one carrying ``superseded_at`` set, so we read the live store
        # back rather than holding a stale reference to the original token.
        assert (await fake.find_unused("hash-old")) is None
        assert fake.by_hash["hash-old"].superseded_at is not None
        assert (await fake.find_unused("hash-new")) is None
        assert fake.by_hash["hash-new"].superseded_at is not None

    asyncio.run(_go())


def test_fake_purge_expired_removes_only_expired() -> None:
    """The fake deletes rows whose ``expires_at < now`` and leaves the rest.

    Because the domain ``ResetToken`` requires ``expires_at > created_at``,
    we cannot hand the fake an actually-expired token via :meth:`insert`.
    We use :func:`dataclasses.replace` to backdate both fields so the
    domain invariant still holds while ``expires_at`` lands in the past.
    """
    fake = FakeResetTokenRepository()
    user_id = uuid4()
    now = datetime.now(UTC)

    async def _go() -> None:
        # Live token: ``expires_at`` 24h into the future.
        await fake.insert(user_id, token_hash="hash-live", expires_at=now + timedelta(hours=24))

        # Expired token: ``expires_at`` and ``created_at`` are both backdated
        # so the token is in the past relative to ``now``, while the domain
        # invariant ``expires_at > created_at`` is preserved (1-second gap).
        expired = await fake.insert(
            user_id, token_hash="hash-expired", expires_at=now + timedelta(seconds=10)
        )
        backdated = dataclasses.replace(
            expired,
            created_at=now - timedelta(hours=26),
            expires_at=now - timedelta(seconds=1),
        )
        # ``ResetToken`` is frozen; ``dataclasses.replace`` returns a new
        # instance, so re-key the fake's store under the same hash.
        fake.by_hash["hash-expired"] = backdated

        purged = await fake.purge_expired(now)
        assert purged == 1
        assert "hash-expired" not in fake.by_hash
        assert "hash-live" in fake.by_hash

    asyncio.run(_go())


def test_pg_skeleton_raises_not_implemented() -> None:
    adapter = ResetTokenRepositoryPg()

    async def _go() -> None:
        await adapter.insert(uuid4(), "hash", datetime.now(UTC))

    try:
        asyncio.run(_go())
    except NotImplementedError:
        return
    raise AssertionError("expected NotImplementedError")
