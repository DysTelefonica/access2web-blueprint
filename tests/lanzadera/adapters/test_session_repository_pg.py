# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W62 (#921)
"""Postgres integration tests for ``SessionRepositoryPg`` (W62, #921).

Gated like ``test_user_repository_pg.py``: the suite needs
``APAP_INTEGRATION_ENABLED=1`` and ``APAP_INTEGRATION_URL`` pointing at a
Postgres with the core schema applied. Without them the tests fail closed
(issue #882), so a missing database cannot pass as green.

Each test seeds its own user (``sessions.user_id`` is a foreign key) and
removes it on teardown; the ``ON DELETE CASCADE`` drops its sessions.
"""

from __future__ import annotations

import os
from collections.abc import AsyncIterator
from datetime import UTC, datetime, timedelta
from typing import Any
from uuid import UUID, uuid4

import pytest
import sqlalchemy as sa

from app.src.modules.lanzadera.adapters.persistence import async_session_factory
from app.src.modules.lanzadera.adapters.persistence.repositories.session_repository_pg import (
    SessionRepositoryPg,
)
from app.src.modules.lanzadera.domain.session import Session

INTEGRATION_URL_ENV = "APAP_INTEGRATION_URL"


def _integration_url_or_fail() -> str:
    """Return the integration Postgres URL, or fail the test (issue #882)."""
    if os.environ.get("APAP_INTEGRATION_ENABLED") != "1":
        pytest.fail(
            "Issue #882: APAP_INTEGRATION_ENABLED is not 1. The integration "
            "DSN fixture must fail closed, not skip."
        )
    url = os.environ.get(INTEGRATION_URL_ENV)
    if not url:
        pytest.fail(f"Issue #882: {INTEGRATION_URL_ENV} is not set.")
    return url


@pytest.fixture
async def repo_and_user() -> AsyncIterator[tuple[SessionRepositoryPg, UUID, Any]]:
    """Yield ``(SessionRepositoryPg, seeded user id, engine)``."""
    engine, factory = async_session_factory(_integration_url_or_fail())
    user_id = uuid4()
    async with engine.begin() as conn:
        await conn.execute(
            sa.text(
                "INSERT INTO lanzadera.users (id, email, name, dni_encrypted) "
                "VALUES (:id, :email, 'Session Test', :dni)"
            ),
            {"id": user_id, "email": f"session-{user_id.hex}@test.invalid", "dni": b"x"},
        )
    try:
        yield SessionRepositoryPg(factory), user_id, engine
    finally:
        async with engine.begin() as conn:
            await conn.execute(
                sa.text("DELETE FROM lanzadera.users WHERE id = :id"), {"id": user_id}
            )
        await engine.dispose()


def _new_session(user_id: UUID, expires_in: timedelta = timedelta(hours=1)) -> Session:
    now = datetime.now(UTC)
    return Session(id=uuid4(), user_id=user_id, created_at=now, expires_at=now + expires_in)


@pytest.mark.integration
async def test_create_then_get_by_id_round_trip(
    repo_and_user: tuple[SessionRepositoryPg, UUID, Any],
) -> None:
    repo, user_id, _ = repo_and_user
    session = _new_session(user_id)

    out = await repo.create(session)
    fetched = await repo.get_by_id(session.id)

    assert out is session
    assert fetched is not None
    assert fetched.id == session.id
    assert fetched.user_id == user_id
    assert fetched.expires_at == session.expires_at


@pytest.mark.integration
async def test_get_by_id_returns_none_for_missing_session(
    repo_and_user: tuple[SessionRepositoryPg, UUID, Any],
) -> None:
    repo, _, _ = repo_and_user
    assert await repo.get_by_id(uuid4()) is None


@pytest.mark.integration
async def test_get_by_id_returns_expired_row(
    repo_and_user: tuple[SessionRepositoryPg, UUID, Any],
) -> None:
    """The repository does not filter expiry; the middleware does (DA-11)."""
    repo, user_id, _ = repo_and_user
    session = _new_session(user_id, expires_in=-timedelta(hours=1))
    await repo.create(session)

    fetched = await repo.get_by_id(session.id)

    assert fetched is not None
    assert fetched.expires_at < datetime.now(UTC)


@pytest.mark.integration
async def test_revoke_shortens_expires_at(
    repo_and_user: tuple[SessionRepositoryPg, UUID, Any],
) -> None:
    repo, user_id, _ = repo_and_user
    session = _new_session(user_id)
    await repo.create(session)
    at = datetime.now(UTC) - timedelta(seconds=1)

    ok = await repo.revoke(session.id, at)
    fetched = await repo.get_by_id(session.id)

    assert ok is True
    assert fetched is not None
    assert fetched.expires_at == at


@pytest.mark.integration
async def test_revoke_never_extends_an_earlier_expiry(
    repo_and_user: tuple[SessionRepositoryPg, UUID, Any],
) -> None:
    """A second, later revocation keeps the earlier timestamp (DA-11)."""
    repo, user_id, _ = repo_and_user
    session = _new_session(user_id)
    await repo.create(session)
    first = datetime.now(UTC) - timedelta(seconds=2)
    second = datetime.now(UTC) - timedelta(seconds=1)

    await repo.revoke(session.id, first)
    await repo.revoke(session.id, second)
    fetched = await repo.get_by_id(session.id)

    assert fetched is not None
    assert fetched.expires_at == first


@pytest.mark.integration
async def test_revoke_returns_false_for_missing_session(
    repo_and_user: tuple[SessionRepositoryPg, UUID, Any],
) -> None:
    repo, _, _ = repo_and_user
    assert await repo.revoke(uuid4(), datetime.now(UTC)) is False
