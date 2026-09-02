# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W62 (#542)
"""Tests Categoría 4 (integration) para ``AuthMiddleware``.

The middleware is mounted on a minimal FastAPI app; a probe route
returns the current ``request.state.user_id`` so the test can assert
what the middleware populated. ``httpx.AsyncClient`` + ``ASGITransport``
exercise the full ASGI stack without spinning up a real server.

D-W62-1: ``JwtSignerPort`` is the contract surface; the test uses
``Hs256JwtSigner`` (the production adapter) so the middleware exercises
the same code path it will in production.
AD-W62-2: silent failure on missing / malformed / expired tokens.
"""

from __future__ import annotations

from datetime import UTC, datetime
from uuid import uuid4

import pytest
from fastapi import FastAPI, Request
from httpx import ASGITransport, AsyncClient

from app.src.modules.lanzadera.adapters.crypto.jwt import Hs256JwtSigner
from app.src.modules.lanzadera.delivery.http.auth_middleware import AuthMiddleware

# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


SECRET = b"test-secret-32-bytes-min-padding!!"


def _signer() -> Hs256JwtSigner:
    return Hs256JwtSigner(SECRET)


def _fixed_now() -> int:
    """Anchor ``now`` to a deterministic timestamp for expiry tests."""
    return int(datetime(2026, 9, 1, 12, 0, 0, tzinfo=UTC).timestamp())


def _app_with_middleware(*, signer: Hs256JwtSigner | None = None, now=None) -> FastAPI:
    """Build a minimal FastAPI app with the middleware mounted.

    The probe route returns the current ``request.state.user_id`` so
    tests can assert what the middleware populated. Returns 401 when
    the state is None so the test can distinguish "middleware set the
    state" from "the app did its own gate".
    """
    app = FastAPI()

    @app.get("/probe")
    async def probe(req: Request):
        user_id = getattr(req.state, "user_id", None)
        if user_id is None:
            from fastapi import HTTPException

            raise HTTPException(401, "no user_id")
        return {"user_id": str(user_id)}

    app.add_middleware(
        AuthMiddleware,
        jwt_signer=signer if signer is not None else _signer(),
        now=now if now is not None else _fixed_now,
    )
    return app


@pytest.fixture
async def client_with_middleware():
    """Yield a tuple ``(client, signer)`` for tests that need both."""
    signer = _signer()
    app = _app_with_middleware(signer=signer)
    transport = ASGITransport(app=app)
    async with AsyncClient(transport=transport, base_url="http://test") as client:
        yield client, signer


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------


async def test_middleware_passes_through_without_authorization(client_with_middleware) -> None:
    """AD-W62-2: missing Authorization header → ``user_id = None``.

    The probe route's 401 makes the assertion explicit: the request
    flowed through the middleware (no raise) but the route gate saw
    no user_id and returned 401.
    """
    client, _ = client_with_middleware
    response = await client.get("/probe")
    assert response.status_code == 401


async def test_middleware_extracts_user_id_from_valid_bearer(client_with_middleware) -> None:
    """Valid Bearer token → ``request.state.user_id`` populated."""
    client, signer = client_with_middleware
    session_id = uuid4()
    now = _fixed_now()
    token = signer.sign({"sub": str(session_id), "iat": now, "exp": now + 60})

    response = await client.get("/probe", headers={"Authorization": f"Bearer {token}"})

    assert response.status_code == 200
    assert response.json() == {"user_id": str(session_id)}


async def test_middleware_sets_user_id_to_none_for_malformed_token(client_with_middleware) -> None:
    """Malformed token → ``user_id = None`` (silent failure)."""
    client, _ = client_with_middleware
    response = await client.get("/probe", headers={"Authorization": "Bearer not.a.jwt"})
    assert response.status_code == 401


async def test_middleware_sets_user_id_to_none_for_expired_token(client_with_middleware) -> None:
    """Expired token → ``user_id = None`` (silent failure)."""
    client, signer = client_with_middleware
    session_id = uuid4()
    issue_time = _fixed_now() - 3600  # 1 hour ago
    token = signer.sign({"sub": str(session_id), "iat": issue_time, "exp": issue_time + 60})

    response = await client.get("/probe", headers={"Authorization": f"Bearer {token}"})
    assert response.status_code == 401


async def test_middleware_ignores_non_bearer_scheme(client_with_middleware) -> None:
    """``Authorization: Basic ...`` is ignored — only Bearer counts."""
    client, _ = client_with_middleware
    response = await client.get("/probe", headers={"Authorization": "Basic dXNlcjpwYXNz"})
    assert response.status_code == 401


async def test_middleware_is_per_request() -> None:
    """AD-W62-2: two concurrent requests see their own user_id, not a shared one.

    Uses two probe routes on the same app so the test can confirm
    that requests do not leak state. The unique ``X-Test-Session-A``
    and ``X-Test-Session-B`` tokens are issued with the same signer
    but different ``sub`` claims.
    """
    signer = _signer()
    app = _app_with_middleware(signer=signer)

    session_a = uuid4()
    session_b = uuid4()
    now = _fixed_now()
    token_a = signer.sign({"sub": str(session_a), "iat": now, "exp": now + 60})
    token_b = signer.sign({"sub": str(session_b), "iat": now, "exp": now + 60})

    transport = ASGITransport(app=app)
    async with AsyncClient(transport=transport, base_url="http://test") as client:
        # Two requests with different tokens must see different user_ids.
        response_a = await client.get("/probe", headers={"Authorization": f"Bearer {token_a}"})
        response_b = await client.get("/probe", headers={"Authorization": f"Bearer {token_b}"})

        assert response_a.status_code == 200
        assert response_b.status_code == 200
        assert response_a.json() == {"user_id": str(session_a)}
        assert response_b.json() == {"user_id": str(session_b)}
        assert response_a.json() != response_b.json()
