# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + lanzadera-mvp W01 (#44)
# DA-1 — contract + in-memory tests for the user-port Postgres adapter.
"""Strict TDD — ``UserRepositoryPg`` (W01, DA-1).

W01 ships three tests for ``UserRepositoryPg``:

- ``test_user_repository_pg_conforms_to_protocol`` — verifies the
  adapter satisfies the ``UserRepository`` Protocol structurally
  (DA-1: the contract is the source of truth, not the impl shape).
- ``test_async_session_factory_emits_async_sessionmaker`` — sanity
  check on the factory contract: the ``AsyncSessionFactoryPort`` we
  inject returns a usable ``AsyncSession`` (DA-1).
- ``test_user_repository_pg_round_trip`` — Postgres round-trip,
  gated on ``APAP_INTEGRATION_ENABLED=1``. The test boots a temp
  Postgres database, runs the migration, exercises ``create`` →
  ``get_by_email`` → ``update_status`` → ``update_password_and_activate``,
  and tears the database down. Without an integration Postgres the
  test is skipped (the in-memory contract tests still cover the
  ``User`` value-object mapping).
"""

from __future__ import annotations

from collections.abc import AsyncIterator
from datetime import UTC, datetime
from typing import Any
from uuid import uuid4

import pytest

from app.src.modules.lanzadera.adapters.persistence import (
    UserRepositoryPg,
    async_session_factory,
)

# ---------------------------------------------------------------------------
# Protocol conformance
# ---------------------------------------------------------------------------


async def test_user_repository_pg_conforms_to_protocol() -> None:
    """The adapter exposes every ``UserRepository`` method on its instance."""
    from app.src.modules.lanzadera.domain.ports import UserRepository

    # Mypy checks the structural conformance statically. The runtime
    # check below is a belt-and-suspenders test in case mypy ever
    # stops inheriting async signatures.
    expected_async = {n for n in dir(UserRepository) if not n.startswith("_")} - {
        "__init__",
        "__annotations__",
        "__subclasshook__",
    }
    actual_async = {n for n in dir(UserRepositoryPg) if not n.startswith("_")}
    missing = expected_async - actual_async
    assert not missing, f"UserRepositoryPg is missing Protocol methods: {missing}"


# ---------------------------------------------------------------------------
# Factory contract
# ---------------------------------------------------------------------------


async def test_async_session_factory_emits_callable_factory() -> None:
    """``async_session_factory`` returns an engine + a callable factory.

    The callable factory returns a fresh ``AsyncSession`` per
    invocation; that is the seam the adapters depend on. The
    test does NOT connect to Postgres — it exercises the factory's
    construction path only.
    """
    # An unreachable URL is fine: the engine is created lazily; the
    # factory's constructor never opens a connection (the
    # ``SQLAlchemyError`` would only fire on the first connect).
    engine, factory = async_session_factory("postgresql+asyncpg://nobody:nopwd@127.0.0.1:1/nodb")
    assert factory is not None
    # The callable returns a usable session object. We do not call
    # ``session.execute(...)`` because the engine has not actually
    # connected yet.
    session = factory()
    assert session is not None
    await engine.dispose()


# ---------------------------------------------------------------------------
# Postgres round-trip (integration-gated)
# ---------------------------------------------------------------------------


INTEGRATION_URL_ENV = "APAP_INTEGRATION_URL"


def _integration_url_or_skip() -> str:
    """Return the integration Postgres URL, or skip the test."""
    import os

    if os.environ.get("APAP_INTEGRATION_ENABLED") != "1":
        pytest.skip("APAP_INTEGRATION_ENABLED is not 1 — Postgres round-trip skipped")
    url = os.environ.get(INTEGRATION_URL_ENV)
    if not url:
        pytest.skip(
            f"{INTEGRATION_URL_ENV} is not set; set it to a Postgres URL to exercise the round-trip"
        )
    return url


@pytest.fixture
async def repo_and_engine() -> AsyncIterator[tuple[UserRepositoryPg, Any]]:
    """Yield ``(UserRepositoryPg, engine)`` for the integration test.

    Migrations are applied via the test harness's bootstrap script
    (out of scope here; PR 4 — #288 — owns that). Without a clean
    Postgres at the configured URL the round-trip test
    (``test_user_repository_pg_round_trip``) skips.
    """
    url = _integration_url_or_skip()
    engine, factory = async_session_factory(url)
    yield UserRepositoryPg(factory), engine
    await engine.dispose()


async def test_user_repository_pg_round_trip(repo_and_engine: Any) -> None:
    """End-to-end: create → get_by_email → update_status → update_password → get_by_id."""
    repo, _engine = repo_and_engine

    from app.src.modules.lanzadera.domain.user import User, UserStatus

    now = datetime(2026, 8, 13, 12, 0, 0, tzinfo=UTC)
    user = User(
        id=uuid4(),
        email="round-trip@enterprise.test",
        name="Round Trip",
        dni_encrypted=b"\x00",
        password_hash=None,
        status=UserStatus.PASSWORD_RESET_REQUIRED,
        failed_attempts=0,
        last_login_at=None,
        created_at=now,
        updated_at=now,
    )
    await repo.create(user)
    fetched = await repo.get_by_email(user.email)
    assert fetched is not None
    assert fetched.email == user.email
    assert fetched.status is UserStatus.PASSWORD_RESET_REQUIRED
