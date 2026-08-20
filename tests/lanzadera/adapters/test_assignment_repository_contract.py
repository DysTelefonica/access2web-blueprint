"""Contract-conformance test for :class:`AssignmentRepositoryPort`.

Pin the adapter skeleton (``AssignmentRepositoryPg``) AND the in-memory
fake (``FakeAssignmentRepository``) against the same Protocol. If the
Protocol drifts, both tests fail. If a fake drifts from the contract,
the fake test fails.

Async I/O uses ``asyncio.run`` (the project's ``pyproject.toml`` does
not enable ``pytest-asyncio`` — see ``tests/lanzadera/migrations/``).
"""

from __future__ import annotations

import asyncio
from uuid import uuid4

from app.src.modules.lanzadera.adapters.repos import AssignmentRepositoryPg
from app.src.modules.lanzadera.domain.ports.assignment_repository import (
    AssignmentRepositoryPort,
)
from tests.lanzadera.adapters._fakes import FakeAssignmentRepository


def test_assignment_pg_satisfies_protocol() -> None:
    adapter: AssignmentRepositoryPort = AssignmentRepositoryPg()
    assert hasattr(adapter, "create")
    assert hasattr(adapter, "list_for_user")
    assert hasattr(adapter, "list_for_app")
    assert hasattr(adapter, "effective_permissions")


def test_fake_satisfies_protocol_create_then_list() -> None:
    """End-to-end smoke for the fake: create → list_for_user finds it."""
    fake: AssignmentRepositoryPort = FakeAssignmentRepository()
    user_id = uuid4()

    async def _go() -> None:
        assignment = await fake.create(user_id=user_id, app_id=1, profile_id=uuid4())
        assert assignment.user_id == user_id
        assert assignment.app_id == 1
        assert assignment.revoked_at is None  # DA-12 — fresh rows are live

        listed = list(await fake.list_for_user(user_id))
        assert len(listed) == 1
        assert listed[0].id == assignment.id

    asyncio.run(_go())


def test_fake_list_for_app_isolates_users() -> None:
    """``list_for_app`` returns assignments for ``app_id`` regardless of user."""
    fake = FakeAssignmentRepository()
    user_a, user_b = uuid4(), uuid4()

    async def _go() -> None:
        await fake.create(user_a, app_id=42, profile_id=uuid4())
        await fake.create(user_b, app_id=42, profile_id=uuid4())
        await fake.create(user_a, app_id=7, profile_id=uuid4())

        app42 = list(await fake.list_for_app(42))
        assert len(app42) == 2
        assert {a.user_id for a in app42} == {user_a, user_b}

        app7 = list(await fake.list_for_app(7))
        assert len(app7) == 1
        assert app7[0].user_id == user_a

    asyncio.run(_go())


def test_fake_effective_permissions_returns_empty_when_no_profiles() -> None:
    """Without profile lookup wired in, the fake returns ``[]`` — this is the
    contract the application layer must tolerate (DA-8).
    """
    fake = FakeAssignmentRepository()

    async def _go() -> list[str]:
        return list(await fake.effective_permissions(uuid4(), app_id=1))

    assert asyncio.run(_go()) == []


def test_pg_skeleton_raises_not_implemented() -> None:
    """The Pg adapter is a skeleton — calling any method raises.

    This is the contract today: the body awaits migrations 0003 + 0006.
    """
    adapter = AssignmentRepositoryPg()

    async def _go() -> None:
        await adapter.create(uuid4(), 1, uuid4())

    try:
        asyncio.run(_go())
    except NotImplementedError:
        return
    raise AssertionError("expected NotImplementedError")
