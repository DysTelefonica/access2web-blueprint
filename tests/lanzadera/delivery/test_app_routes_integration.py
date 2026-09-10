# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W62 PR-7
# W61 (#524) + W62 PR-7 (#544) — integration tests for the app CRUD HTTP routes.
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

W62 PR-7 (#544): the ``X-Admin-User-ID`` header stub that the W61 routes
parsed manually has been replaced by ``request.state.user_id``
populated by ``AuthMiddleware`` (PR-5). The test app wires a
``_TestAuthMiddleware`` that reads the ``X-Test-Session-Id`` header and
looks up the seeded session in the container's ``session_repo`` — the
same seam the production ``AuthMiddleware`` populates after a successful
JWT verify, but without requiring the test path to mint + verify a
token (the JWT coverage lives in
``tests/lanzadera/delivery/test_auth_routes_integration.py``).

The W61 contract continues to accept a missing actor on the app routes
(see ``admin_routes_apps.py`` docstring): ``_actor_id(request)`` returns
``None`` when the header is absent and the use case accepts that value
without raising.

Issue #576: all four routes (``create_app``, ``get_app``, ``update_app``,
``disable_app``) now call ``await _admin.require_global_admin(request)``
as their first statement, mirroring the gate already used by
``admin_routes_users.py`` / ``admin_routes_misc.py``. Every test below
therefore seeds ``auth_session`` and forwards the ``X-Test-Session-Id``
header; the ``*_requires_admin_without_session`` tests pin the 401 the
gate returns for an unauthenticated caller.
"""

from __future__ import annotations

from collections.abc import Awaitable, Callable, Iterator
from datetime import UTC, datetime
from typing import TYPE_CHECKING
from uuid import UUID

import pytest
from fastapi import APIRouter, FastAPI
from fastapi.testclient import TestClient
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request as _StarletteRequest
from starlette.responses import Response

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


# ---------------------------------------------------------------------------
# Test app factory — mounts the apps router with the fakes-backed container.
# ---------------------------------------------------------------------------


class _TestAuthMiddleware(BaseHTTPMiddleware):
    """Test-only middleware that populates ``request.state.user_id`` from ``X-Test-Session-Id``.

    W62 PR-7 (#544) replaces the W61 ``X-Admin-User-ID`` header stub
    (parsed by ``_actor_id`` in ``admin_routes_apps.py``) with
    ``request.state.user_id`` populated by the production
    ``AuthMiddleware`` (PR-5). The production middleware verifies a
    Bearer JWT against the container's ``jwt_signer``; the test path
    exercises the same end-to-end coverage without minting + verifying a
    token per test by reading the test-only ``X-Test-Session-Id``
    header and looking up the seeded ``Session`` row in the
    container's ``session_repo``. If the header is missing or the
    session is unknown, the middleware leaves ``request.state.user_id``
    ``None`` — mirroring the AD-W62-2 silent-failure contract — and
    ``_actor_id(request)`` returns ``None`` so the audit-actor field on
    the use case stays explicit.
    """

    async def dispatch(
        self,
        request: _StarletteRequest,
        call_next: Callable[[_StarletteRequest], Awaitable[Response]],
    ) -> Response:
        # Default state: no auth. Mirrors ``AuthMiddleware.dispatch``
        # before the Bearer branch — the production silent-failure
        # contract (AD-W62-2) leaves ``request.state.user_id`` empty.
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


def _build_app(container: LanzaderaContainer) -> FastAPI:
    """Minimal FastAPI app with the apps router wired to ``container``.

    The lifespan and the production ``app.state.container`` dance are
    unnecessary here — the routes resolve the container via
    ``request.state.container``, so a small middleware mirrors the
    app-level container onto every request. The router's prefix is
    ``/admin`` so the test paths match the production surface.

    W62 PR-7 (#544): the app also mounts ``_TestAuthMiddleware`` so the
    W62 auth path (Bearer → ``request.state.user_id`` → ``_actor_id``)
    is exercised end-to-end against the seeded session. The mirror
    middleware stays in place because the W61 routes' ``_container``
    helper still reads ``request.state.container`` (the production
    seam).
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
    app.add_middleware(_TestAuthMiddleware)
    return app


def _auth_headers(session_id: UUID) -> dict[str, str]:
    """Return the ``X-Test-Session-Id`` header that ``_TestAuthMiddleware`` reads.

    Helper for the W62 PR-7 auth-aware tests. The middleware looks up
    the seeded session in ``container.sessions`` and copies the
    session's ``user_id`` onto ``request.state.user_id`` — the same
    seam the production ``AuthMiddleware`` populates after a successful
    JWT verify.
    """
    return {"X-Test-Session-Id": str(session_id)}


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


@pytest.mark.usefixtures("auth_session")
def test_post_apps_creates_row_with_pending_status(
    client: TestClient, apps_fake: FakeAppRepository, auth_session: dict[str, UUID]
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
        headers=_auth_headers(auth_session["session_id"]),
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


@pytest.mark.usefixtures("auth_session")
def test_post_apps_rejects_empty_name(
    client: TestClient, apps_fake: FakeAppRepository, auth_session: dict[str, UUID]
) -> None:
    """``POST /admin/apps`` returns 400 when ``name`` is empty / whitespace."""
    response = client.post(
        "/admin/apps",
        json={
            "name": "   ",
            "short_code": "EXP",
            "deployment_topology": "central",
        },
        headers=_auth_headers(auth_session["session_id"]),
    )

    assert response.status_code == 400
    assert "name" in response.json()["detail"]
    assert apps_fake.create_calls == []


@pytest.mark.usefixtures("auth_session")
def test_post_apps_rejects_empty_short_code(
    client: TestClient, auth_session: dict[str, UUID]
) -> None:
    """``POST /admin/apps`` returns 400 when ``short_code`` is empty."""
    response = client.post(
        "/admin/apps",
        json={
            "name": "Lanzadera",
            "short_code": "",
            "deployment_topology": "central",
        },
        headers=_auth_headers(auth_session["session_id"]),
    )

    assert response.status_code == 400
    assert "short_code" in response.json()["detail"]


@pytest.mark.usefixtures("auth_session")
def test_post_apps_rejects_unknown_topology(
    client: TestClient, auth_session: dict[str, UUID]
) -> None:
    """``POST /admin/apps`` returns 400 when ``deployment_topology`` is invalid."""
    response = client.post(
        "/admin/apps",
        json={
            "name": "Lanzadera",
            "short_code": "lanza",
            "deployment_topology": "edge-cloud",  # not in AppTopology
        },
        headers=_auth_headers(auth_session["session_id"]),
    )

    assert response.status_code == 400
    assert "deployment_topology" in response.json()["detail"]


def test_post_apps_requires_admin_without_session(client: TestClient) -> None:
    """Issue #576: ``POST /admin/apps`` returns 401 without an auth session.

    Reproduces the vulnerability report's first ``curl`` call — no
    ``Authorization`` header, no ``X-Test-Session-Id`` header. Before the
    fix this call reached the use case and returned 201.
    """
    response = client.post(
        "/admin/apps",
        json={
            "name": "Lanzadera",
            "short_code": "lanza",
            "deployment_topology": "central",
        },
    )

    assert response.status_code == 401


# ---------------------------------------------------------------------------
# GET /admin/apps/{app_id}
# ---------------------------------------------------------------------------


@pytest.mark.usefixtures("auth_session")
def test_get_apps_returns_row(
    client: TestClient, fake_fixtures: FakeFixtures, auth_session: dict[str, UUID]
) -> None:
    """``GET /admin/apps/{id}`` returns the row when it exists."""
    _seed_active_app(fake_fixtures, app_id=42)

    response = client.get("/admin/apps/42", headers=_auth_headers(auth_session["session_id"]))

    assert response.status_code == 200
    body = response.json()
    assert body["id"] == 42
    assert body["name"] == "Expedientes"
    assert body["short_code"] == "EXP"
    assert body["deployment_topology"] == "central"
    assert body["registration_status"] == "active"


@pytest.mark.usefixtures("auth_session")
def test_get_apps_returns_404_for_unknown_id(
    client: TestClient, auth_session: dict[str, UUID]
) -> None:
    """``GET /admin/apps/{id}`` returns 404 when the row does not exist."""
    response = client.get("/admin/apps/999", headers=_auth_headers(auth_session["session_id"]))

    assert response.status_code == 404
    assert "999" in response.json()["detail"]


def test_get_apps_requires_admin_without_session(client: TestClient) -> None:
    """Issue #576: ``GET /admin/apps/{id}`` returns 401 without an auth session.

    ``get_app`` exposes internal catalog metadata and is treated as an
    admin-only read, matching every other ``/admin`` read-route.
    """
    response = client.get("/admin/apps/42")

    assert response.status_code == 401


# ---------------------------------------------------------------------------
# PATCH /admin/apps/{app_id}
# ---------------------------------------------------------------------------


@pytest.mark.usefixtures("auth_session")
def test_patch_apps_updates_partial_fields(
    client: TestClient,
    fake_fixtures: FakeFixtures,
    apps_fake: FakeAppRepository,
    auth_session: dict[str, UUID],
) -> None:
    """``PATCH /admin/apps/{id}`` applies the partial patch and returns the new row."""
    _seed_active_app(fake_fixtures, app_id=1)

    response = client.patch(
        "/admin/apps/1",
        json={"name": "Expedientes v2", "requires_office_presence": True},
        headers=_auth_headers(auth_session["session_id"]),
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


@pytest.mark.usefixtures("auth_session")
def test_patch_apps_rejects_invalid_topology(
    client: TestClient, auth_session: dict[str, UUID]
) -> None:
    """``PATCH /admin/apps/{id}`` returns 400 on an invalid ``deployment_topology``."""
    response = client.patch(
        "/admin/apps/1",
        json={"deployment_topology": "edge-cloud"},
        headers=_auth_headers(auth_session["session_id"]),
    )

    assert response.status_code == 400
    assert "deployment_topology" in response.json()["detail"]


def test_patch_apps_requires_admin_without_session(
    client: TestClient, fake_fixtures: FakeFixtures
) -> None:
    """Issue #576: ``PATCH /admin/apps/{id}`` returns 401 without an auth session."""
    _seed_active_app(fake_fixtures, app_id=1)

    response = client.patch(
        "/admin/apps/1",
        json={"deployment_topology": "office-nas"},
    )

    assert response.status_code == 401


# ---------------------------------------------------------------------------
# DELETE /admin/apps/{app_id}
# ---------------------------------------------------------------------------


@pytest.mark.usefixtures("auth_session")
def test_delete_apps_flips_to_retired(
    client: TestClient,
    fake_fixtures: FakeFixtures,
    apps_fake: FakeAppRepository,
    auth_session: dict[str, UUID],
) -> None:
    """``DELETE /admin/apps/{id}`` flips the row to ``retired`` and returns 204."""
    _seed_active_app(fake_fixtures, app_id=5)

    response = client.delete("/admin/apps/5", headers=_auth_headers(auth_session["session_id"]))

    assert response.status_code == 204
    assert response.content == b""
    assert apps_fake.disable_calls == [5]
    # The fake's by_id carries the post-disable row.
    assert apps_fake.by_id[5].registration_status is AppRegistrationStatus.RETIRED


def test_delete_apps_requires_admin_without_session(
    client: TestClient, fake_fixtures: FakeFixtures
) -> None:
    """Issue #576: ``DELETE /admin/apps/{id}`` returns 401 without an auth session."""
    _seed_active_app(fake_fixtures, app_id=5)

    response = client.delete("/admin/apps/5")

    assert response.status_code == 401


# ---------------------------------------------------------------------------
# W62 PR-7 auth path — X-Test-Session-Id → request.state.user_id
# ---------------------------------------------------------------------------


@pytest.mark.usefixtures("auth_session")
def test_post_apps_requires_auth_session(client: TestClient, auth_session: dict[str, UUID]) -> None:
    """``POST /admin/apps`` accepts the ``X-Test-Session-Id`` header and returns 201.

    W62 PR-7 (#544): the W61 ``X-Admin-User-ID`` header stub is
    retired; the new auth path is ``AuthMiddleware`` →
    ``request.state.user_id`` → ``_actor_id(request)``. The test app
    wires ``_TestAuthMiddleware`` (mirroring the production
    ``AuthMiddleware`` for the JWT verify step) that reads
    ``X-Test-Session-Id`` and looks up the seeded ``Session`` row in
    ``container.sessions``. This test confirms the seam works: the
    seeded session id produces a populated ``request.state.user_id``
    and the destructive route still returns 201.

    The ``auth_session`` fixture (see ``tests/lanzadera/conftest.py``)
    seeds a global admin + a 24 h ``Session`` row via
    ``FakeSessionRepository``. The session id is forwarded in the
    ``X-Test-Session-Id`` header so the test middleware can resolve
    it end-to-end against the same ``session_repo`` the production
    middleware hits after a successful JWT verify.
    """
    response = client.post(
        "/admin/apps",
        json={
            "name": "Lanzadera",
            "short_code": "lanza",
            "deployment_topology": "central",
        },
        headers=_auth_headers(auth_session["session_id"]),
    )

    assert response.status_code == 201


__all__ = [
    "test_delete_apps_flips_to_retired",
    "test_delete_apps_requires_admin_without_session",
    "test_get_apps_requires_admin_without_session",
    "test_get_apps_returns_404_for_unknown_id",
    "test_get_apps_returns_row",
    "test_patch_apps_rejects_invalid_topology",
    "test_patch_apps_requires_admin_without_session",
    "test_patch_apps_updates_partial_fields",
    "test_post_apps_creates_row_with_pending_status",
    "test_post_apps_requires_admin_without_session",
    "test_post_apps_requires_auth_session",
    "test_post_apps_rejects_empty_name",
    "test_post_apps_rejects_empty_short_code",
    "test_post_apps_rejects_unknown_topology",
]
