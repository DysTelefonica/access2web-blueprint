# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W61
# W61 (#524) — integration tests for the app CRUD HTTP routes.
"""Integration tests for the W61 (#524) ``/admin/apps/...`` JSON routes.

The routes go through the full chain ``HTTP route → LanzaderaContainer →
use case → repository``. This module wires the container with the
in-memory fakes from ``tests/lanzadera/_fakes.py`` and exercises the
routes end-to-end via FastAPI's ``TestClient``.

The tests cover the W61 slice:

- ``POST /admin/apps`` — create a catalog row, return 201 + serialised body.
- ``POST /admin/apps`` — reject empty ``name`` / ``short_code`` (400).
- ``POST /admin/apps`` — reject unknown ``deployment_topology`` (400).
- ``GET  /admin/apps/{id}`` — return the row (200) and 404 for unknown.
- ``PATCH /admin/apps/{id}`` — apply a partial patch and return 200.
- ``DELETE /admin/apps/{id}`` — flip to ``retired`` and return 204.

The ``X-Admin-User-ID`` header is accepted by the destructive routes
(a missing header is a no-op until W62 lands the real auth middleware).
"""

from __future__ import annotations

from collections.abc import Iterator
from datetime import UTC, datetime
from typing import TYPE_CHECKING
from uuid import uuid4

import pytest
from fastapi import APIRouter, FastAPI
from fastapi.testclient import TestClient

from app.src.modules.lanzadera.delivery.http.admin_routes_apps import (
    router as apps_router,
)
from app.src.modules.lanzadera.di.container import LanzaderaContainer
from app.src.modules.lanzadera.domain.app import (
    App,
    AppRegistrationStatus,
    AppTopology,
)
from tests.lanzadera._fakes import FakeAppRepository

if TYPE_CHECKING:
    from tests.lanzadera.conftest import FakeFixtures


def _build_app(container: LanzaderaContainer) -> FastAPI:
    """Minimal FastAPI app with the apps router wired to ``container``.

    The lifespan and the production ``app.state.container`` dance are
    unnecessary here — the routes resolve the container via
    ``request.state.container``, so a small middleware mirrors the
    app-level container onto every request. The router's prefix is
    ``/admin`` so the test paths match the production surface.
    """
    app = FastAPI()
    app.state.container = container  # type: ignore[attr-defined]

    @app.middleware("http")
    async def _mirror_container(request, call_next):  # type: ignore[no-untyped-def]
        request.state.container = app.state.container
        return await call_next(request)

    router = APIRouter(prefix="/admin")
    router.include_router(apps_router)
    app.include_router(router)
    return app


@pytest.fixture
def client(container: LanzaderaContainer) -> Iterator[TestClient]:
    """Yield a ``TestClient`` whose ``/admin/apps`` routes use the fakes-backed container."""
    app = _build_app(container)
    with TestClient(app) as test_client:
        yield test_client


@pytest.fixture
def apps_fake(fake_fixtures: FakeFixtures) -> FakeAppRepository:
    """Expose the app fake directly for assertions."""
    return fake_fixtures.apps


def _seed_active_app(fakes: FakeFixtures, *, app_id: int = 1) -> App:
    """Seed an ``ACTIVE`` app the routes can target."""
    app = App(
        id=app_id,
        name="Expedientes",
        short_code="EXP",
        deployment_topology=AppTopology.CENTRAL,
        requires_office_presence=False,
        registration_status=AppRegistrationStatus.ACTIVE,
        created_at=datetime.now(UTC),
        updated_at=datetime.now(UTC),
    )
    fakes.apps.add(app)
    return app


# ---------------------------------------------------------------------------
# POST /admin/apps
# ---------------------------------------------------------------------------


def test_post_apps_creates_row_with_pending_status(
    client: TestClient, apps_fake: FakeAppRepository
) -> None:
    """``POST /admin/apps`` creates the row and returns 201 with the serialised body."""
    response = client.post(
        "/admin/apps",
        json={
            "name": "Lanzadera",
            "short_code": "lanza",
            "deployment_topology": "central",
            "requires_office_presence": False,
        },
    )

    assert response.status_code == 201
    body = response.json()
    assert body["name"] == "Lanzadera"
    assert body["short_code"] == "lanza"
    assert body["deployment_topology"] == "central"
    assert body["requires_office_presence"] is False
    assert body["registration_status"] == "pending"
    assert isinstance(body["id"], int)
    # The fake recorded exactly one create call.
    assert len(apps_fake.create_calls) == 1
    assert apps_fake.create_calls[0].short_code == "lanza"


def test_post_apps_rejects_empty_name(client: TestClient, apps_fake: FakeAppRepository) -> None:
    """``POST /admin/apps`` returns 400 when ``name`` is empty / whitespace."""
    response = client.post(
        "/admin/apps",
        json={
            "name": "   ",
            "short_code": "EXP",
            "deployment_topology": "central",
        },
    )

    assert response.status_code == 400
    assert "name" in response.json()["detail"]
    assert apps_fake.create_calls == []


def test_post_apps_rejects_empty_short_code(client: TestClient) -> None:
    """``POST /admin/apps`` returns 400 when ``short_code`` is empty."""
    response = client.post(
        "/admin/apps",
        json={
            "name": "Lanzadera",
            "short_code": "",
            "deployment_topology": "central",
        },
    )

    assert response.status_code == 400
    assert "short_code" in response.json()["detail"]


def test_post_apps_rejects_unknown_topology(client: TestClient) -> None:
    """``POST /admin/apps`` returns 400 when ``deployment_topology`` is invalid."""
    response = client.post(
        "/admin/apps",
        json={
            "name": "Lanzadera",
            "short_code": "lanza",
            "deployment_topology": "edge-cloud",  # not in AppTopology
        },
    )

    assert response.status_code == 400
    assert "deployment_topology" in response.json()["detail"]


# ---------------------------------------------------------------------------
# GET /admin/apps/{app_id}
# ---------------------------------------------------------------------------


def test_get_apps_returns_row(client: TestClient, fake_fixtures: FakeFixtures) -> None:
    """``GET /admin/apps/{id}`` returns the row when it exists."""
    _seed_active_app(fake_fixtures, app_id=42)

    response = client.get("/admin/apps/42")

    assert response.status_code == 200
    body = response.json()
    assert body["id"] == 42
    assert body["name"] == "Expedientes"
    assert body["short_code"] == "EXP"
    assert body["deployment_topology"] == "central"
    assert body["registration_status"] == "active"


def test_get_apps_returns_404_for_unknown_id(client: TestClient) -> None:
    """``GET /admin/apps/{id}`` returns 404 when the row does not exist."""
    response = client.get("/admin/apps/999")

    assert response.status_code == 404
    assert "999" in response.json()["detail"]


# ---------------------------------------------------------------------------
# PATCH /admin/apps/{app_id}
# ---------------------------------------------------------------------------


def test_patch_apps_updates_partial_fields(
    client: TestClient, fake_fixtures: FakeFixtures, apps_fake: FakeAppRepository
) -> None:
    """``PATCH /admin/apps/{id}`` applies the partial patch and returns the new row."""
    _seed_active_app(fake_fixtures, app_id=1)

    response = client.patch(
        "/admin/apps/1",
        json={"name": "Expedientes v2", "requires_office_presence": True},
    )

    assert response.status_code == 200
    body = response.json()
    assert body["name"] == "Expedientes v2"
    assert body["requires_office_presence"] is True
    # The untouched fields keep their original values.
    assert body["deployment_topology"] == "central"
    assert body["registration_status"] == "active"
    assert len(apps_fake.update_calls) == 1
    assert apps_fake.update_calls[0][0] == 1
    assert apps_fake.update_calls[0][1] == {
        "name": "Expedientes v2",
        "requires_office_presence": True,
    }


def test_patch_apps_rejects_invalid_topology(client: TestClient) -> None:
    """``PATCH /admin/apps/{id}`` returns 400 on an invalid ``deployment_topology``."""
    response = client.patch(
        "/admin/apps/1",
        json={"deployment_topology": "edge-cloud"},
    )

    assert response.status_code == 400
    assert "deployment_topology" in response.json()["detail"]


# ---------------------------------------------------------------------------
# DELETE /admin/apps/{app_id}
# ---------------------------------------------------------------------------


def test_delete_apps_flips_to_retired(
    client: TestClient, fake_fixtures: FakeFixtures, apps_fake: FakeAppRepository
) -> None:
    """``DELETE /admin/apps/{id}`` flips the row to ``retired`` and returns 204."""
    _seed_active_app(fake_fixtures, app_id=5)

    response = client.delete("/admin/apps/5")

    assert response.status_code == 204
    assert response.content == b""
    assert apps_fake.disable_calls == [5]
    # The fake's by_id carries the post-disable row.
    assert apps_fake.by_id[5].registration_status is AppRegistrationStatus.RETIRED


# ---------------------------------------------------------------------------
# X-Admin-User-ID header — W62 stub
# ---------------------------------------------------------------------------


def test_post_apps_accepts_admin_user_header(client: TestClient) -> None:
    """``POST /admin/apps`` accepts a valid ``X-Admin-User-ID`` header without error."""
    response = client.post(
        "/admin/apps",
        json={
            "name": "Lanzadera",
            "short_code": "lanza",
            "deployment_topology": "central",
        },
        headers={"X-Admin-User-ID": str(uuid4())},
    )

    assert response.status_code == 201


def test_post_apps_rejects_invalid_admin_user_header(client: TestClient) -> None:
    """``POST /admin/apps`` returns 400 when ``X-Admin-User-ID`` is not a valid UUID."""
    response = client.post(
        "/admin/apps",
        json={
            "name": "Lanzadera",
            "short_code": "lanza",
            "deployment_topology": "central",
        },
        headers={"X-Admin-User-ID": "not-a-uuid"},
    )

    assert response.status_code == 400
    assert "X-Admin-User-ID" in response.json()["detail"]


__all__ = [
    "test_delete_apps_flips_to_retired",
    "test_get_apps_returns_404_for_unknown_id",
    "test_get_apps_returns_row",
    "test_patch_apps_rejects_invalid_topology",
    "test_patch_apps_updates_partial_fields",
    "test_post_apps_accepts_admin_user_header",
    "test_post_apps_creates_row_with_pending_status",
    "test_post_apps_rejects_empty_name",
    "test_post_apps_rejects_empty_short_code",
    "test_post_apps_rejects_invalid_admin_user_header",
    "test_post_apps_rejects_unknown_topology",
]
