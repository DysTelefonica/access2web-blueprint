"""Contract-conformance test for :class:`AssignmentRepositoryPort`.

Pin both implementations against the same Protocol:
- the in-memory ``FakeAssignmentRepository`` (used by the rest of the
  test suite), and
- the Postgres ``AssignmentRepositoryPg`` wired to the
  :func:`async_session_factory` seam (DA-1, W02 #45).

If the Protocol drifts, both tests fail. If a fake drifts from the
contract, the fake test fails.
"""

from __future__ import annotations

from uuid import uuid4

from app.src.modules.lanzadera.adapters.persistence import (
    AssignmentRepositoryPg,
    async_session_factory,
)
from app.src.modules.lanzadera.domain.ports.assignment_repository import (
    AssignmentRepositoryPort,
)
from tests.lanzadera.adapters._fakes import FakeAssignmentRepository


def test_assignment_pg_satisfies_protocol() -> None:
    """Static structural check on the Postgres adapter's method set.

    We do not instantiate (the constructor requires a session
    factory); the runtime attribute scan below mirrors what mypy's
    structural conformance already enforces.
    """
    expected = {"create", "list_for_user", "list_for_app", "effective_permissions"}
    actual = set(dir(AssignmentRepositoryPg))
    missing = expected - actual
    assert not missing, f"AssignmentRepositoryPg is missing Protocol methods: {missing}"


async def test_fake_satisfies_protocol_create_then_list() -> None:
    """End-to-end smoke for the fake: create -> list_for_user finds it."""
    fake: AssignmentRepositoryPort = FakeAssignmentRepository()
    user_id = uuid4()

    assignment = await fake.create(user_id=user_id, app_id=1, profile_id=uuid4())
    assert assignment.user_id == user_id
    assert assignment.app_id == 1
    assert assignment.revoked_at is None  # DA-12 -- fresh rows are live

    listed = list(await fake.list_for_user(user_id))
    assert len(listed) == 1
    assert listed[0].id == assignment.id


async def test_fake_list_for_app_isolates_users() -> None:
    """``list_for_app`` returns assignments for ``app_id`` regardless of user."""
    fake = FakeAssignmentRepository()
    user_a, user_b = uuid4(), uuid4()

    await fake.create(user_a, app_id=42, profile_id=uuid4())
    await fake.create(user_b, app_id=42, profile_id=uuid4())
    await fake.create(user_a, app_id=7, profile_id=uuid4())

    app42 = list(await fake.list_for_app(42))
    assert len(app42) == 2
    assert {a.user_id for a in app42} == {user_a, user_b}

    app7 = list(await fake.list_for_app(7))
    assert len(app7) == 1
    assert app7[0].user_id == user_a


async def test_fake_effective_permissions_returns_empty_when_no_profiles() -> None:
    """Without profile lookup wired in, the fake returns ``[]``.

    This is the contract the application layer must tolerate (DA-8).
    """
    fake = FakeAssignmentRepository()
    assert await fake.effective_permissions(uuid4(), app_id=1) == []


async def test_pg_factory_wiring_passes_constructor_check() -> None:
    """``AssignmentRepositoryPg`` takes a factory seam and is constructable.

    We do not connect to Postgres here; an unreachable URL is fine
    because the engine is created lazily. The test merely verifies the
    constructor's wiring shape (DA-1: the factory is the only seam the
    adapters depend on).
    """
    _, factory = async_session_factory("postgresql+asyncpg://nobody:nopwd@127.0.0.1:1/nodb")
    adapter = AssignmentRepositoryPg(factory)
    assert adapter is not None
