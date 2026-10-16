"""Issue #773 — ``/auth/logout`` must invalidate the JWT.

Behavioural coverage against real Postgres (#921 ships
``SessionRepositoryPg``; CI job #901 wires the suite to a real DB).
Without the fix these four cases fail: ``AuthMiddleware`` verifies the
JWT signature but never asks the session repo whether the row is
alive. Each case goes through a probe route that returns ``401`` when
``request.state.user_id`` is ``None`` — the same check
``require_global_admin`` and ``/auth/me`` perform on the same state —
so the assertion pins the public contract, not middleware internals.
The DSN fixture fails closed (issue #882) when the integration URL is
missing; there is no silent green.
"""

from __future__ import annotations

import os
from collections.abc import AsyncIterator
from datetime import UTC, datetime, timedelta
from uuid import UUID, uuid4

import pytest
import sqlalchemy as sa
from fastapi import FastAPI, HTTPException, Request
from httpx import ASGITransport, AsyncClient

from app.src.modules.lanzadera.adapters.crypto.jwt import Hs256JwtSigner
from app.src.modules.lanzadera.adapters.persistence import (
    async_session_factory,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.session_repository_pg import (
    SessionRepositoryPg,
)
from app.src.modules.lanzadera.delivery.http.auth_middleware import (
    AuthMiddleware,
)

_JWT_SECRET: bytes = b"x" * 32  # noqa: S105


def _integration_url_or_fail() -> str:
    """Issue #882: a missing DSN fails the gate instead of silently passing."""
    if os.environ.get("APAP_INTEGRATION_ENABLED") != "1":
        pytest.fail("APAP_INTEGRATION_ENABLED is not 1 (issue #882 fail-closed).")
    url = os.environ.get("APAP_INTEGRATION_URL")
    if not url:
        pytest.fail("APAP_INTEGRATION_URL is not set.")
    return url


_INSERT_USER_SQL = (
    "INSERT INTO lanzadera.users "
    "(id, email, name, dni_encrypted, password_hash, status) "
    "VALUES (:id, :email, :name, :dni, :hash, 'active')"
)
_INSERT_SESSION_SQL = (
    "INSERT INTO lanzadera.sessions (id, user_id, created_at, expires_at) "
    "VALUES (:id, :uid, :created, :expires)"
)


@pytest.fixture
async def seeded_user() -> AsyncIterator[UUID]:
    """Seed an active platform user; FK cascade drops the test sessions on cleanup."""
    engine, _ = async_session_factory(_integration_url_or_fail())
    user_id = uuid4()
    async with engine.begin() as conn:
        await conn.execute(
            sa.text(_INSERT_USER_SQL),
            {
                "id": user_id,
                "email": f"lanzadera-773-{user_id.hex}@enterprise.test",
                "name": "Lanzadera 773",
                "dni": b"encrypted:dni:773",
                "hash": "middleware-test-only",
            },
        )
    try:
        yield user_id
    finally:
        async with engine.begin() as conn:
            await conn.execute(
                sa.text("DELETE FROM lanzadera.users WHERE id = :id"),
                {"id": user_id},
            )
        await engine.dispose()


def _build_app(session_repo: SessionRepositoryPg | object) -> FastAPI:
    """A minimal FastAPI app with the middleware mounted against ``session_repo``.

    The probe route mirrors ``require_global_admin`` and ``/auth/me``:
    ``state.user_id is None`` ⇒ ``401``, otherwise ``200``.
    """

    class _Container:
        sessions = session_repo  # middleware reads it via request.app.state.container

    app = FastAPI()
    app.state.container = _Container()

    @app.get("/probe")
    async def _probe(req: Request) -> dict[str, str]:
        if getattr(req.state, "user_id", None) is None:
            raise HTTPException(status_code=401, detail="authentication required")
        return {"user_id": str(req.state.user_id)}

    app.add_middleware(
        AuthMiddleware,
        jwt_signer=Hs256JwtSigner(_JWT_SECRET),
        now=lambda: int(datetime.now(UTC).timestamp()),
    )
    return app


async def _insert_session(
    engine: object, *, session_id: UUID, user_id: UUID, expires_at: datetime
) -> None:
    async with engine.begin() as conn:
        await conn.execute(
            sa.text(_INSERT_SESSION_SQL),
            {
                "id": session_id,
                "uid": user_id,
                "created": datetime.now(UTC),
                "expires": expires_at,
            },
        )


def _mint(sub: UUID, *, exp_offset_seconds: int = 3600) -> str:
    now = int(datetime.now(UTC).timestamp())
    return Hs256JwtSigner(_JWT_SECRET).sign(
        {"sub": str(sub), "iat": now - 60, "exp": now + exp_offset_seconds}
    )


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------


@pytest.mark.integration
async def test_revoked_session_returns_401_on_probe(
    seeded_user: UUID,
) -> None:
    """Seeded session → 200; after ``SessionRepositoryPg.revoke`` with same token → 401.

    ``revoke`` is what ``POST /auth/logout`` calls. The bug is that the
    middleware, consulting only the JWT signature, does not notice
    when the session row has been revoked.
    """
    engine, factory = async_session_factory(_integration_url_or_fail())
    real_repo = SessionRepositoryPg(factory)
    session_id = uuid4()
    await _insert_session(
        engine,
        session_id=session_id,
        user_id=seeded_user,
        expires_at=datetime.now(UTC) + timedelta(hours=24),
    )
    app, token = _build_app(real_repo), _mint(session_id)
    bearer = {"Authorization": f"Bearer {token}"}

    async with AsyncClient(transport=ASGITransport(app=app), base_url="http://test") as client:
        assert (await client.get("/probe", headers=bearer)).status_code == 200
        await real_repo.revoke(session_id, datetime.now(UTC))
        response = await client.get("/probe", headers=bearer)
        assert response.status_code == 401, response.text
    await engine.dispose()


@pytest.mark.integration
async def test_unknown_session_id_returns_401_on_probe(
    seeded_user: UUID,
) -> None:
    """JWT whose ``sub`` is a UUID absent from ``lanzadera.sessions`` → 401.

    A valid signature is not enough — the middleware must refuse a
    token whose session row does not exist.
    """
    engine, factory = async_session_factory(_integration_url_or_fail())
    app = _build_app(SessionRepositoryPg(factory))
    async with AsyncClient(transport=ASGITransport(app=app), base_url="http://test") as client:
        response = await client.get("/probe", headers={"Authorization": f"Bearer {_mint(uuid4())}"})
        assert response.status_code == 401
    await engine.dispose()


@pytest.mark.integration
async def test_expired_session_row_returns_401_on_probe(
    seeded_user: UUID,
) -> None:
    """``sessions.expires_at`` in the past → 401, even with a valid JWT.

    Distinct from an expired-JWT case: the JWT signature and ``exp``
    claim verify cleanly, but the session row's ``expires_at`` is in
    the past. Only the middleware consulting the repo can catch this.
    """
    engine, factory = async_session_factory(_integration_url_or_fail())
    real_repo = SessionRepositoryPg(factory)
    session_id = uuid4()
    await _insert_session(
        engine,
        session_id=session_id,
        user_id=seeded_user,
        expires_at=datetime.now(UTC) + timedelta(minutes=5),
    )
    app = _build_app(real_repo)
    bearer = {"Authorization": f"Bearer {_mint(session_id, exp_offset_seconds=3600)}"}

    async with AsyncClient(transport=ASGITransport(app=app), base_url="http://test") as client:
        assert (await client.get("/probe", headers=bearer)).status_code == 200
        async with engine.begin() as conn:
            await conn.execute(
                sa.text("UPDATE lanzadera.sessions SET expires_at = :p WHERE id = :i"),
                {"p": datetime.now(UTC) - timedelta(minutes=1), "i": session_id},
            )
        response = await client.get("/probe", headers=bearer)
        assert response.status_code == 401, response.text
    await engine.dispose()


@pytest.mark.integration
async def test_session_repo_unreachable_fails_closed(
    seeded_user: UUID,
) -> None:
    """When ``sessions.get_by_id`` raises, ``/probe`` → 401 (never 200).

    Fail-closed (#773): a Postgres outage that prevents confirming
    the session cannot be treated as "session absent → allow through".
    """

    class _BrokenRepo:
        async def get_by_id(self, _session_id: UUID) -> None:
            raise RuntimeError("simulated postgres outage (issue #773 fail-closed)")

    app = _build_app(_BrokenRepo())  # type: ignore[arg-type]
    async with AsyncClient(transport=ASGITransport(app=app), base_url="http://test") as client:
        response = await client.get("/probe", headers={"Authorization": f"Bearer {_mint(uuid4())}"})
        assert response.status_code == 401, response.text
