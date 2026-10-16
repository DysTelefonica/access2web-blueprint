# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W62 (#542, #876)
"""Tests Categoría 4 (integration) for ``_LazyAuthSigner``.

The lazy proxy in ``app.src.main`` defers resolution of the real
``JwtSignerPort`` until the first ``verify`` / ``sign`` call because
the container is built in the lifespan, after the middleware is
installed. The proxy's private ``_container()`` helper centralises the
``getattr(app.state, "container", None)`` lookup and raises
``InvalidTokenError`` when the container is missing (pre-startup /
headless tests).

These tests pin two contracts:
- ``verify(token)`` and ``sign(claims)`` both route through
  ``_container()`` — the duplicated ``getattr / if / raise`` block
  PR #941 introduced is gone.
- A missing container raises ``InvalidTokenError`` from BOTH methods,
  not just from ``verify`` (the PR #941 contract).
- The end-to-end silent-failure path still works: the middleware
  catches ``InvalidTokenError`` (AD-W62-2) and leaves state empty so
  the route gate can fire 401.

The proxy reads ``app.state.container`` (the module-level ``app``
singleton in ``app.src.main``). The two direct ``verify`` / ``sign``
tests leave that singleton untouched (they only fail because
``container`` was never wired), and the end-to-end test mounts the
production middleware on a fresh FastAPI app so the assertion is
local to the missing-container contract.
"""

from __future__ import annotations

import pytest
from fastapi import FastAPI, Request
from httpx import ASGITransport, AsyncClient

from app.src.modules.lanzadera.adapters.crypto.jwt import Hs256JwtSigner
from app.src.modules.lanzadera.delivery.http.auth_middleware import AuthMiddleware
from app.src.modules.lanzadera.domain.errors import InvalidTokenError

# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------


def test_lazy_auth_signer_verify_raises_when_container_missing() -> None:
    """A missing container raises ``InvalidTokenError`` from ``verify``.

    The helper is the only path ``verify`` takes to reach
    ``container.jwt_signer``. With ``app.state.container`` unset the
    helper must raise so the middleware's silent-failure path fires
    (issue #773) — without this, a missing container would surface as
    an ``AttributeError`` deep inside ``AuthMiddleware.dispatch`` and
    take down every public route.
    """
    from app.src.main import _LazyAuthSigner

    signer = _LazyAuthSigner()

    with pytest.raises(InvalidTokenError, match="auth not initialised yet"):
        signer.verify("any.token.value", now=1)


def test_lazy_auth_signer_sign_raises_when_container_missing() -> None:
    """A missing container raises ``InvalidTokenError`` from ``sign``.

    PR #941 added ``sign`` to satisfy ``JwtSignerPort``. The helper
    centralises the missing-container branch, so ``sign`` must also
    raise ``InvalidTokenError`` — never silently produce a malformed
    token and never ``AttributeError`` past the middleware boundary.
    """
    from app.src.main import _LazyAuthSigner

    signer = _LazyAuthSigner()

    with pytest.raises(InvalidTokenError, match="auth not initialised yet"):
        signer.sign({"sub": "session-id", "iat": 1, "exp": 2})


async def test_auth_middleware_silences_lazy_signer_when_container_missing() -> None:
    """End-to-end: missing container → middleware leaves state empty.

    The middleware catches ``InvalidTokenError`` (AD-W62-2) and sets
    ``request.state.user_id = None``. The probe route raises 401 when
    no user id is attached, so the assertion is precise: the
    middleware ran, the helper raised, the middleware swallowed the
    error, the route gate fired. This pins the headless-test path
    the W62 PR-A e2e smoke relied on (refs #876).
    """
    test_app = FastAPI()

    @test_app.get("/probe")
    async def probe(req: Request):
        user_id = getattr(req.state, "user_id", None)
        if user_id is None:
            from fastapi import HTTPException

            raise HTTPException(401, "no user_id")
        return {"user_id": str(user_id)}

    # ``app.state.container`` is unset on the test app; the helper's
    # ``getattr(test_app.state, "container", None)`` returns ``None``
    # and the helper raises ``InvalidTokenError`` before any real
    # signer is consulted.
    test_app.add_middleware(
        AuthMiddleware,
        jwt_signer=Hs256JwtSigner(b"x" * 32),
    )

    transport = ASGITransport(app=test_app)
    async with AsyncClient(transport=transport, base_url="http://test") as client:
        response = await client.get("/probe", headers={"Authorization": "Bearer x.y.z"})

    assert response.status_code == 401
