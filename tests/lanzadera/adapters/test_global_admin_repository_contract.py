"""Contract-conformance test for :class:`GlobalAdminRepositoryPort`.

Pin both implementations against the same Protocol:
- the in-memory ``FakeGlobalAdminRepository`` (used by the rest of the
  test suite), and
- the Postgres ``GlobalAdminRepositoryPg`` wired to the
  :func:`async_session_factory` seam (DA-1, W04 #21).
"""

from __future__ import annotations

from uuid import uuid4

from app.src.modules.lanzadera.adapters.persistence import (
    GlobalAdminRepositoryPg,
    async_session_factory,
)
from app.src.modules.lanzadera.domain.ports.global_admin_repository import (
    GlobalAdminRepositoryPort,
)
from tests.lanzadera.adapters._fakes import FakeGlobalAdminRepository


def test_global_admin_pg_satisfies_protocol() -> None:
    """Static structural check on the Postgres adapter's method set."""
    expected = {"list_all", "is_global_admin", "grant", "revoke"}
    actual = set(dir(GlobalAdminRepositoryPg))
    missing = expected - actual
    assert not missing, f"GlobalAdminRepositoryPg is missing Protocol methods: {missing}"


async def test_fake_grant_then_is_global_admin_then_list() -> None:
    fake: GlobalAdminRepositoryPort = FakeGlobalAdminRepository()
    user_id = uuid4()

    assert await fake.is_global_admin(user_id) is False
    await fake.grant(user_id)
    assert await fake.is_global_admin(user_id) is True
    assert await fake.list_all()


async def test_global_admin_pg_construction() -> None:
    """``GlobalAdminRepositoryPg`` takes the factory seam and is constructable (DA-1).

    We do not connect to Postgres here; an unreachable URL is fine
    because the engine is created lazily.
    """
    _, factory = async_session_factory("postgresql+asyncpg://nobody:nopwd@127.0.0.1:1/nodb")
    adapter = GlobalAdminRepositoryPg(factory)
    assert adapter is not None
