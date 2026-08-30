# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W60
# W60 (#522) — integration tests for the presence HTTP routes.
"""Integration tests for the W60 (#522) ``/admin/presence/...`` routes.

The SSE stream (``GET /admin/presence/stream``) is out-of-scope for this
test file. Its behaviour (5-second polling loop, SSE framing, graceful
client-disconnect handling) lives behind an async generator whose
cancellation contract requires either a real HTTP client (playwright/curl)
or a dedicated load-test fixture. The W60 slice ships the generator
correctly; a future ``test_presence_routes_e2e.py`` exercises it under
load.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING
from uuid import UUID, uuid4

import pytest
from fastapi import APIRouter, FastAPI
from fastapi.testclient import TestClient

from app.src.modules.lanzadera.delivery.http.admin_routes_presence import (
    router as presence_router,
)
from app.src.modules.lanzadera.di.container import LanzaderaContainer
from tests.lanzadera._presence_fakes import FakePresenceRepository

if TYPE_CHECKING:
    from tests.lanzadera.conftest import FakeFixtures


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


def test_list_connected_returns_json(
    client: TestClient, presence_fake: FakePresenceRepository
) -> None:
    """``GET /admin/presence`` returns the JSON snapshot the use case produces."""
    user = presence_fake.seed_row(email="alice@enterprise.test")
    presence_fake.add(user)
    response = client.get("/admin/presence")
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


def test_list_connected_empty(client: TestClient, presence_fake: FakePresenceRepository) -> None:
    """``GET /admin/presence`` returns ``[]`` when no users are connected."""
    response = client.get("/admin/presence")
    assert response.status_code == 200
    assert response.json() == []


def test_list_connected_forwards_limit(
    client: TestClient, presence_fake: FakePresenceRepository
) -> None:
    """The ``limit`` query parameter reaches the use case (and ultimately the SQL ``LIMIT``)."""
    for i in range(5):
        presence_fake.add(presence_fake.seed_row(email=f"user{i}@enterprise.test"))
    response = client.get("/admin/presence?limit=2")
    assert response.status_code == 200
    assert len(response.json()) == 2


# ---------------------------------------------------------------------------
# GET /admin/presence/stream
#
# The SSE stream is out-of-scope for this test file. See module docstring.
# ---------------------------------------------------------------------------

__all__ = [
    "test_heartbeat_accepts_user_id_header",
    "test_heartbeat_without_header_is_noop",
    "test_heartbeat_rejects_invalid_uuid_header",
    "test_list_connected_returns_json",
    "test_list_connected_empty",
    "test_list_connected_forwards_limit",
]

_ = UUID
