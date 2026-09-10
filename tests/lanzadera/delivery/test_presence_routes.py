# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W60
# W60 (#522) — integration tests for the presence HTTP routes.
# Issue #577 — list_connected / presence_stream now gate on
# require_global_admin; heartbeat stays ungated (out of scope, reports
# only the caller's own presence).
"""Integration tests for the W60 (#522) ``/admin/presence/...`` routes.

The SSE stream (``GET /admin/presence/stream``) is out-of-scope for this
test file for its streaming behaviour. Its 5-second polling loop, SSE
framing, and graceful client-disconnect handling live behind an async
generator whose cancellation contract requires either a real HTTP client
(playwright/curl) or a dedicated load-test fixture. The W60 slice ships
the generator correctly; a future ``test_presence_routes_e2e.py``
exercises it under load. The auth gate (issue #577) is exercised here
because it runs before the generator is ever created.

Issue #577: ``list_connected`` and ``presence_stream`` now call
``require_global_admin`` as their first statement, so the test app wires
the same ``_TestAuthMiddleware`` the app/users/misc route test files use
— it reads ``X-Test-Session-Id`` and looks up the seeded session in the
container's ``session_repo`` to populate ``request.state.user_id``.
``heartbeat`` is intentionally left out of this gate (a different,
lower-risk endpoint reporting the caller's own presence) and its tests
below remain unauthenticated.
"""

from __future__ import annotations

from collections.abc import Awaitable, Callable, Iterator
from typing import TYPE_CHECKING
from uuid import UUID, uuid4

import pytest
from fastapi import APIRouter, FastAPI, HTTPException
from fastapi.testclient import TestClient
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request as _StarletteRequest
from starlette.responses import Response

from app.src.modules.lanzadera.delivery.http.admin_routes_presence import (
    presence_stream,
)
from app.src.modules.lanzadera.delivery.http.admin_routes_presence import (
    router as presence_router,
)
from app.src.modules.lanzadera.di.container import LanzaderaContainer
from tests.lanzadera._presence_fakes import FakePresenceRepository

if TYPE_CHECKING:
    from tests.lanzadera.conftest import FakeFixtures


class _TestAuthMiddleware(BaseHTTPMiddleware):
    """Test-only middleware that populates ``request.state.user_id`` from ``X-Test-Session-Id``.

    Mirrors the production ``AuthMiddleware`` (PR-5) end-to-end coverage
    without minting + verifying a JWT per test — see
    ``tests/lanzadera/delivery/test_app_routes_integration.py`` for the
    original of this pattern.
    """

    async def dispatch(
        self,
        request: _StarletteRequest,
        call_next: Callable[[_StarletteRequest], Awaitable[Response]],
    ) -> Response:
        request.state.user_id = None
        request.state.session_id = None

        session_id_header = request.headers.get("X-Test-Session-Id")
        if session_id_header:
            try:
                session_id = UUID(session_id_header)
            except ValueError:
                session_id = None
            if session_id is not None:
                container: LanzaderaContainer = request.app.state.container  # type: ignore[attr-defined]
                session = await container.sessions.get_by_id(session_id)
                if session is not None:
                    request.state.session_id = session.id
                    request.state.user_id = session.user_id

        return await call_next(request)


def _auth_headers(session_id: UUID) -> dict[str, str]:
    """Return the ``X-Test-Session-Id`` header that ``_TestAuthMiddleware`` reads."""
    return {"X-Test-Session-Id": str(session_id)}


def _build_app(container: LanzaderaContainer) -> FastAPI:
    """Minimal FastAPI app with the presence router wired to ``container``."""
    app = FastAPI()
    app.state.container = container  # type: ignore[attr-defined]

    @app.middleware("http")
    async def _mirror_container(request, call_next):  # type: ignore[no-untyped-def]
        request.state.container = app.state.container
        return await call_next(request)

    router = APIRouter(prefix="/admin")
    router.include_router(presence_router)
    app.include_router(router)
    app.add_middleware(_TestAuthMiddleware)
    return app


@pytest.fixture
def client(container: LanzaderaContainer) -> Iterator[TestClient]:
    """Yield a ``TestClient`` whose presence routes use the fakes-backed container."""
    app = _build_app(container)
    with TestClient(app) as test_client:
        yield test_client


@pytest.fixture
def presence_fake(fake_fixtures: FakeFixtures) -> FakePresenceRepository:
    """Expose the presence fake directly for assertions."""
    return fake_fixtures.presence


# ---------------------------------------------------------------------------
# POST /admin/presence/heartbeat
# ---------------------------------------------------------------------------


def test_heartbeat_accepts_user_id_header(
    client: TestClient, presence_fake: FakePresenceRepository
) -> None:
    """``POST /admin/presence/heartbeat`` reads X-User-ID and forwards the heartbeat call."""
    user_id = uuid4()
    response = client.post(
        "/admin/presence/heartbeat",
        headers={"X-User-ID": str(user_id)},
    )
    assert response.status_code == 200
    assert response.json() == {"ok": True}
    assert presence_fake.heartbeat_calls == [user_id]


def test_heartbeat_without_header_is_noop(
    client: TestClient, presence_fake: FakePresenceRepository
) -> None:
    """Missing X-User-ID is treated as a no-op (200 OK, no repo call)."""
    response = client.post("/admin/presence/heartbeat")
    assert response.status_code == 200
    assert response.json() == {"ok": True}
    assert presence_fake.heartbeat_calls == []


def test_heartbeat_rejects_invalid_uuid_header(client: TestClient) -> None:
    """A non-UUID X-User-ID returns 400 Bad Request."""
    response = client.post(
        "/admin/presence/heartbeat",
        headers={"X-User-ID": "not-a-uuid"},
    )
    assert response.status_code == 400
    assert "X-User-ID" in response.json()["detail"]


# ---------------------------------------------------------------------------
# GET /admin/presence
# ---------------------------------------------------------------------------


@pytest.mark.usefixtures("auth_session")
def test_list_connected_returns_json(
    client: TestClient, presence_fake: FakePresenceRepository, auth_session: dict[str, UUID]
) -> None:
    """``GET /admin/presence`` returns the JSON snapshot the use case produces."""
    user = presence_fake.seed_row(email="alice@enterprise.test")
    presence_fake.add(user)
    response = client.get("/admin/presence", headers=_auth_headers(auth_session["session_id"]))
    assert response.status_code == 200
    body = response.json()
    assert isinstance(body, list)
    assert len(body) == 1
    row = body[0]
    assert row["user_id"] == str(user.user_id)
    assert row["email"] == "alice@enterprise.test"
    assert row["connected_at"] == user.connected_at.isoformat()
    assert row["last_seen"] == user.last_seen.isoformat()
    assert presence_fake.list_connected_calls == 1


@pytest.mark.usefixtures("auth_session")
def test_list_connected_empty(
    client: TestClient, presence_fake: FakePresenceRepository, auth_session: dict[str, UUID]
) -> None:
    """``GET /admin/presence`` returns ``[]`` when no users are connected."""
    response = client.get("/admin/presence", headers=_auth_headers(auth_session["session_id"]))
    assert response.status_code == 200
    assert response.json() == []


@pytest.mark.usefixtures("auth_session")
def test_list_connected_forwards_limit(
    client: TestClient, presence_fake: FakePresenceRepository, auth_session: dict[str, UUID]
) -> None:
    """The ``limit`` query parameter reaches the use case (and ultimately the SQL ``LIMIT``)."""
    for i in range(5):
        presence_fake.add(presence_fake.seed_row(email=f"user{i}@enterprise.test"))
    response = client.get(
        "/admin/presence?limit=2", headers=_auth_headers(auth_session["session_id"])
    )
    assert response.status_code == 200
    assert len(response.json()) == 2


def test_list_connected_requires_admin_without_session(
    client: TestClient, presence_fake: FakePresenceRepository
) -> None:
    """Issue #577: ``GET /admin/presence`` returns 401 without an auth session.

    Before the fix, an unauthenticated caller could list every connected
    user's ``user_id`` / ``email`` / presence timestamps.
    """
    presence_fake.add(presence_fake.seed_row(email="alice@enterprise.test"))

    response = client.get("/admin/presence")

    assert response.status_code == 401


# ---------------------------------------------------------------------------
# GET /admin/presence/stream
#
# The streaming behaviour is out-of-scope for this test file (see module
# docstring); the auth gate (issue #577) runs before the stream starts, so
# it is covered here like any other route.
# ---------------------------------------------------------------------------


async def test_presence_stream_requires_admin_without_session(
    container: LanzaderaContainer,
) -> None:
    """Issue #577: ``presence_stream`` raises 401 without an auth session.

    This route is deliberately NOT exercised through ``TestClient`` /
    ``_TestAuthMiddleware`` like the rest of this file: before the fix,
    an unauthenticated call falls through into ``event_generator``'s
    ``while True`` polling loop, and Starlette's ``BaseHTTPMiddleware``
    (the base of ``_TestAuthMiddleware``) is documented to deadlock when
    it wraps a never-ending ``StreamingResponse`` that the test client
    never disconnects from — the test process hangs rather than failing.
    Calling the route coroutine directly with a hand-built ``Request``
    reproduces the exact vulnerability (an unauthenticated caller reaches
    the SSE body) without ever entering that generator, since
    ``require_global_admin`` — when present — raises before the
    ``StreamingResponse`` is constructed.
    """
    app = FastAPI()
    scope = {
        "type": "http",
        "method": "GET",
        "path": "/admin/presence/stream",
        "headers": [],
        "app": app,
    }
    request = _StarletteRequest(scope)
    request.state.container = container  # type: ignore[attr-defined]

    with pytest.raises(HTTPException) as exc_info:
        await presence_stream(request)

    assert exc_info.value.status_code == 401


__all__ = [
    "test_heartbeat_accepts_user_id_header",
    "test_heartbeat_without_header_is_noop",
    "test_heartbeat_rejects_invalid_uuid_header",
    "test_list_connected_returns_json",
    "test_list_connected_empty",
    "test_list_connected_forwards_limit",
    "test_list_connected_requires_admin_without_session",
    "test_presence_stream_requires_admin_without_session",
]

_ = UUID
