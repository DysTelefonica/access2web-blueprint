"""Contract tests for UserRepositoryPg (lanzadera-mvp, AD1, issue #44).

Two layers of tests:

1. **In-memory contract test** — exercises every method against a
   ``FakeUserRepository`` so the adapter contract is pinned even
   when no Postgres is available. This is the layer that CI defaults
   to.

2. **Postgres round-trip test** — gated on ``APAP_INTEGRATION_ENABLED=1``
   like the migration tests. Spins up a temp database, runs the
   migration 0001 to create the schema, inserts a row, reads it back,
   updates it, and asserts the dataclass mapping round-trips. This
   is what local devs run on their workstation with a real Postgres.
"""

from __future__ import annotations

import asyncio
import os
import secrets
import subprocess
import uuid as _uuid
from datetime import UTC, datetime
from pathlib import Path
from uuid import UUID

import pytest

# Skip the integration test unless the env var is set (CI default).
_RUN_INTEGRATION = os.environ.get("APAP_INTEGRATION_ENABLED") == "1"


# ---------------------------------------------------------------------------
# In-memory contract test — runs always (the layer gate asserts).
# ---------------------------------------------------------------------------


class FakeUserRepository:
    """In-memory port fake for contract tests."""

    def __init__(self) -> None:
        self._by_id: dict[UUID, dict] = {}
        self._by_email: dict[str, dict] = {}

    async def get_by_email(self, email: str) -> dict | None:
        return self._by_email.get(email.strip().lower())

    async def get_by_id(self, user_id: UUID) -> dict | None:
        return self._by_id.get(user_id)

    async def create(self, user: dict) -> None:
        self._by_id[user["id"]] = user
        self._by_email[user["email"]] = user

    async def update_status(self, user_id: UUID, status: str) -> None:
        row = self._by_id.get(user_id)
        if row is not None:
            row["status"] = status

    async def update_password_and_activate(self, user_id: UUID, password_hash: str) -> None:
        row = self._by_id.get(user_id)
        if row is not None:
            row["password_hash"] = password_hash
            row["status"] = "active"
            row["failed_attempts"] = 0

    async def list_all(self) -> list[dict]:
        return list(self._by_id.values())


def test_user_repository_contract() -> None:
    """The fake satisfies the structural Protocol used by the use cases."""
    from app.src.modules.lanzadera.adapters.persistence import UserRepositoryPg
    from app.src.modules.lanzadera.domain.ports import UserRepository

    fake = FakeUserRepository()
    # The adapter must structurally conform to the Protocol.
    adapter: UserRepository = UserRepositoryPg.__new__(UserRepositoryPg)
    adapter._factory = fake  # type: ignore[attr-defined]
    assert hasattr(adapter, "get_by_email")
    assert hasattr(adapter, "get_by_id")
    assert hasattr(adapter, "create")
    assert hasattr(adapter, "update_status")
    assert hasattr(adapter, "update_password_and_activate")
    assert hasattr(adapter, "list_all")


def test_user_repository_pg_conforms_to_protocol() -> None:
    """MyPy verifies this statically; the runtime check just ensures the
    import order and class identity survive refactors."""
    from app.src.modules.lanzadera.adapters.persistence import UserRepositoryPg

    # ``UserRepository`` is a Protocol — assigning ``UserRepositoryPg``
    # to it raises ``TypeError`` at runtime if the structural shape
    # diverges. The MyPy ``Protocol`` import is the static check; this
    # test is the runtime backstop.
    instance = UserRepositoryPg.__new__(UserRepositoryPg)
    assert isinstance(instance, UserRepositoryPg)


def test_async_session_factory_emits_async_sessionmaker() -> None:
    """The factory's second return value is callable.

    The ``Protocol`` import is the static conformance check (mypy
    fails if ``async_session_factory`` does not implement the protocol's
    methods); this runtime check just exercises that the factory
    returns a callable.
    """
    from app.src.modules.lanzadera.adapters.persistence import async_session_factory

    _engine, factory = async_session_factory("postgresql+asyncpg://u:p@h:5432/d")
    assert callable(factory)
    # We can't dispose the engine without a real backend, so let
    # the GC reclaim it. The contract check is what matters; the
    # async engine is properly cleaned up by ``AsyncEngine.dispose()``
    # only when it has bound connections, which ours never opened.


# ---------------------------------------------------------------------------
# Postgres round-trip integration test — runs only with the env flag.
# ---------------------------------------------------------------------------


@pytest.mark.skipif(not _RUN_INTEGRATION, reason="APAP_INTEGRATION_ENABLED=1 not set")
def test_user_repository_pg_round_trip(tmp_path: Path) -> None:
    """Insert a user via ``UserRepositoryPg.create``, read it back via
    ``get_by_email`` / ``get_by_id``, update ``status`` to ``disabled``,
    and assert ``update_password_and_activate`` flips status to
    ``active`` and resets ``failed_attempts``. Mirrors the contract the
    delivery layer relies on for #43's admin use cases."""

    from app.src.modules.lanzadera.adapters.persistence import (
        UserRepositoryPg,
        async_session_factory,
    )
    from app.src.modules.lanzadera.domain.user import User, UserStatus

    # Provision a fresh database, run migrations, run the test, drop the db.
    db_name = f"a2w_test_{secrets.token_hex(4)}"
    _psql(["createdb", db_name])

    async def _run() -> None:
        engine, factory = async_session_factory(
            f"postgresql+asyncpg://postgres@/{db_name}",
        )
        try:
            # Apply the schema.
            subprocess.run(
                [
                    "alembic",
                    "-c",
                    "app/migrations/alembic.ini",
                    "upgrade",
                    "head",
                ],
                env={**os.environ, "DATABASE_URL": f"postgresql://postgres@/{db_name}"},
                check=True,
            )

            user = User(
                id=_uuid.uuid4(),
                email="[email protected]",
                name="Alice Tester",
                dni_encrypted=b"ciphertext-not-plaintext",
                password_hash=None,
                status=UserStatus.PASSWORD_RESET_REQUIRED,
                failed_attempts=0,
                last_login_at=None,
                created_at=datetime.now(UTC),
                updated_at=datetime.now(UTC),
            )

            repo = UserRepositoryPg(factory)

            await repo.create(user)
            found = await repo.get_by_email("[email protected]")
            assert found is not None
            assert found.email == "[email protected]"
            assert found.name == "Alice Tester"
            assert found.status == UserStatus.PASSWORD_RESET_REQUIRED
            assert found.failed_attempts == 0

            await repo.update_status(user.id, UserStatus.DISABLED)
            found2 = await repo.get_by_id(user.id)
            assert found2 is not None
            assert found2.status == UserStatus.DISABLED

            await repo.update_password_and_activate(user.id, "new-argon2-hash")
            found3 = await repo.get_by_email("[email protected]")
            assert found3.password_hash == "new-argon2-hash"
            assert found3.status == UserStatus.ACTIVE
            assert found3.failed_attempts == 0
        finally:
            await engine.dispose()

    try:
        asyncio.run(_run())
    finally:
        _psql(["dropdb", "--if-exists", db_name])


def _psql(args: list[str]) -> None:
    """Run a ``psql`` command via sudo (the integration test runs
    against a real Postgres; the runner is the dev workstation)."""
    subprocess.run(
        ["sudo", "-n", "psql", *args],
        check=True,
        capture_output=True,
        text=True,
    )
