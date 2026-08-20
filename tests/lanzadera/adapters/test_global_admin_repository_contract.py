"""Contract-conformance test for :class:`GlobalAdminRepositoryPort`.

Async I/O uses ``asyncio.run`` (the project's ``pyproject.toml`` does
not enable ``pytest-asyncio``).
"""

from __future__ import annotations

import asyncio
from uuid import uuid4

from app.src.modules.lanzadera.adapters.repos import GlobalAdminRepositoryPg
from app.src.modules.lanzadera.domain.ports.global_admin_repository import (
    GlobalAdminRepositoryPort,
)
from tests.lanzadera.adapters._fakes import FakeGlobalAdminRepository


def test_global_admin_pg_satisfies_protocol() -> None:
    adapter: GlobalAdminRepositoryPort = GlobalAdminRepositoryPg()
    for method in ("list_all", "is_global_admin", "grant", "revoke"):
        assert hasattr(adapter, method), f"missing {method}"


def test_fake_grant_then_is_global_admin_then_list() -> None:
    fake = FakeGlobalAdminRepository()
    user_id = uuid4()

    async def _go() -> None:
        assert await fake.is_global_admin(user_id) is False
        await fake.grant(user_id)
        assert await fake.is_global_admin(user_id) is True

        listed = list(await fake.list_all())
        assert len(listed) == 1
        assert listed[0].user_id == user_id

    asyncio.run(_go())


def test_fake_grant_idempotency_fails_when_already_admin() -> None:
    fake = FakeGlobalAdminRepository()
    user_id = uuid4()

    async def _go() -> None:
        await fake.grant(user_id)
        try:
            await fake.grant(user_id)
        except ValueError as exc:
            assert "already a global admin" in str(exc)
            return
        raise AssertionError("expected ValueError")

    asyncio.run(_go())


def test_fake_revoke_removes_admin() -> None:
    fake = FakeGlobalAdminRepository()
    user_id = uuid4()
    other = uuid4()

    async def _go() -> None:
        await fake.grant(user_id)
        await fake.grant(other)
        await fake.revoke(user_id)
        assert await fake.is_global_admin(user_id) is False
        assert await fake.is_global_admin(other) is True

    asyncio.run(_go())


def test_fake_revoke_last_admin_rejected() -> None:
    """D42: the system must always have at least one global admin."""
    fake = FakeGlobalAdminRepository()
    last = uuid4()

    async def _go() -> None:
        await fake.grant(last)
        try:
            await fake.revoke(last)
        except ValueError as exc:
            assert "last global admin" in str(exc)
            assert await fake.is_global_admin(last) is True
            return
        raise AssertionError("expected ValueError")

    asyncio.run(_go())


def test_pg_skeleton_raises_not_implemented() -> None:
    adapter = GlobalAdminRepositoryPg()

    async def _go() -> None:
        await adapter.grant(uuid4())

    try:
        asyncio.run(_go())
    except NotImplementedError:
        return
    raise AssertionError("expected NotImplementedError")
