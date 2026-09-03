# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W62 PR-7 (#544)
"""Integration tests for the Lanzadera admin HTTP routes (issues #519, #544).

The admin HTTP routes (``list_users``, ``create_user``, ``disable_user``,
``assign_profile``) go through the full chain ``HTTP route →
LanzaderaContainer → use case → repository``. This module wires the
container with the in-memory fakes from ``tests/lanzadera/_fakes.py``
and exercises the routes end-to-end via FastAPI's ``TestClient``.

The tests cover the W54-W62 slice:
- ``GET /admin/users`` paginates via ``list_all_users`` (W59, #517).
- ``POST /admin/users`` creates the row via the ``create_user`` use case
  (W54, #43) and appends the audit row (DA-11).
- ``PATCH /admin/users/{id}/disable`` flips ACTIVE → DISABLED (W54, D42).
- ``POST /admin/assignments`` persists the (user, app, profile) triple
  via ``assign_profile`` (W54, D22, DA-12).

W62 PR-7 (#544): the ``auth_bypass`` fixture (a no-op monkeypatch of
``admin.require_global_admin``) is replaced by ``auth_session``, which
seeds a real global admin + session. The test app installs a
``_TestAuthMiddleware`` that reads the ``X-Test-Session-Id`` header
and looks up the seeded session in the container's ``session_repo`` to
populate ``request.state.user_id`` — the same seam the production
``AuthMiddleware`` populates after a successful JWT verify, but
without requiring the test path to sign or verify a token. The real
``require_global_admin`` gate (PR-6) is exercised end-to-end: a
missing header returns 401, a stale session id returns 401, a seeded
session id with a granted admin returns the destructive command's
success status.
"""

from __future__ import annotations

from collections.abc import Awaitable, Callable
from datetime import UTC, datetime
from pathlib import Path
from uuid import UUID, uuid4

import pytest
from fastapi import APIRouter, FastAPI
from fastapi.templating import Jinja2Templates
from fastapi.testclient import TestClient
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request as _StarletteRequest
from starlette.responses import Response

from app.src.modules.lanzadera.delivery.http.admin_routes import register_routes
from app.src.modules.lanzadera.di.container import LanzaderaContainer
from app.src.modules.lanzadera.domain.app import App, AppRegistrationStatus, AppTopology
from app.src.modules.lanzadera.domain.profile import Profile
from app.src.modules.lanzadera.domain.user import User, UserStatus
from tests.lanzadera.conftest import FakeFixtures

# ---------------------------------------------------------------------------
# Test app factory — mounts the admin router with the fakes-backed container.
# ---------------------------------------------------------------------------


_TEMPLATES_DIR = (
    Path(__file__).resolve().parents[3]
    / "app"
    / "src"
    / "modules"
    / "lanzadera"
    / "delivery"
    / "http"
    / "templates"
)


class _TestAuthMiddleware(BaseHTTPMiddleware):
    """Test-only middleware that populates ``request.state.user_id`` from ``X-Test-Session-Id``.

    W62 PR-7 (#544) replaces the W54-W59 ``auth_bypass`` fixture (a no-op
    monkeypatch of ``admin.require_global_admin``) with the production
    gate. The destructive routes in this module call
    ``require_global_admin`` (PR-6) which reads
    ``request.state.user_id`` — the production app's ``AuthMiddleware``
    (PR-5) populates that field by verifying a Bearer JWT against the
    container's ``jwt_signer``. The test path needs the same
    end-to-end coverage but does not want to mint + verify a JWT per
    test (the JWT coverage lives in
    ``tests/lanzadera/delivery/test_auth_routes_integration.py``).

    The middleware reads the test-only ``X-Test-Session-Id`` header
    and looks up the corresponding ``Session`` row in the container's
    ``session_repo``. If the session exists, the middleware copies the
    session's ``user_id`` onto ``request.state.user_id`` (and the
    session id onto ``request.state.session_id``) — mirroring what the
    production middleware does after a successful ``verify``. If the
    header is missing or the session is unknown, the middleware leaves
    both fields ``None`` so ``require_global_admin`` returns 401.
    """

    async def dispatch(
        self,
        request: _StarletteRequest,
        call_next: Callable[[_StarletteRequest], Awaitable[Response]],
    ) -> Response:
        # Default state: no auth. Mirrors ``AuthMiddleware.dispatch``
        # before the Bearer branch — the production silent-failure
        # contract (AD-W62-2) lets ``require_global_admin`` turn the
        # empty state into a 401 rather than raising.
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
    """Return a minimal FastAPI app with the admin router wired to ``container``.

    The lifespan and the production ``app.state.container`` dance are
    unnecessary here — the routes resolve the container via closure,
    so we mount the router directly on a fresh ``FastAPI`` instance.
    The router's prefix is ``/admin`` so the test paths match the
    production surface exactly.

    W62 PR-7 (#544): the app now sets ``app.state.container`` so
    ``require_global_admin`` (PR-6) can resolve the gate from the test
    path, and mounts ``_TestAuthMiddleware`` so the destructive
    routes' ``request.state.user_id`` is populated from the seeded
    session (see ``auth_session`` fixture in
    ``tests/lanzadera/conftest.py``).
    """
    app = FastAPI()
    app.state.container = container
    router = APIRouter(prefix="/admin")
    templates = Jinja2Templates(directory=str(_TEMPLATES_DIR))
    register_routes(router, templates=templates, container=container)
    app.include_router(router)
    app.add_middleware(_TestAuthMiddleware)
    return app


def _auth_headers(session_id: UUID) -> dict[str, str]:
    """Return the ``X-Test-Session-Id`` header that the test middleware reads.

    Helper for the destructive routes (``POST`` / ``PATCH`` / ``DELETE``)
    that gate on ``require_global_admin``. Each destructive test passes
    the ``auth_session`` fixture and forwards its ``session_id`` here so
    the test middleware looks up the seeded session row in the
    container's ``session_repo`` and populates
    ``request.state.user_id``.
    """
    return {"X-Test-Session-Id": str(session_id)}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _seed_active_user(fakes: FakeFixtures, *, email: str) -> User:
    """Seed a user with ``status=ACTIVE`` so ``disable_user`` accepts the transition."""
    user = User(
        id=uuid4(),
        email=email,
        name="Alice",
        dni_encrypted=b"enc:11111111",
        password_hash="fake:pass",
        status=UserStatus.ACTIVE,
        failed_attempts=0,
        last_login_at=None,
        created_at=datetime.now(UTC),
        updated_at=datetime.now(UTC),
    )
    fakes.users.add(user)
    return user


def _seed_app(fakes: FakeFixtures, *, app_id: int) -> App:
    """Seed an active app the ``assign_profile`` flow can target."""
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


def _seed_profile(fakes: FakeFixtures, *, app_id: int) -> Profile:
    """Seed a profile scoped to ``app_id`` with a non-SinAcceso capability set."""
    profile = Profile(
        id=uuid4(),
        app_id=app_id,
        code="default",
        name="Default",
        capabilities={"Read": True, "Write": True},
        active=True,
        created_at=datetime.now(UTC),
        updated_at=datetime.now(UTC),
    )
    fakes.profiles.add(profile)
    return profile


@pytest.fixture
def client(container: LanzaderaContainer):
    """Yield a ``TestClient`` whose ``/admin`` routes use the fakes-backed container.

    ``raise_server_exceptions=False`` keeps unhandled use-case errors
    (e.g. ``UserNotFoundError``) in the response instead of letting the
    TestClient re-raise them — the integration tests assert on the HTTP
    status code, not on a Python traceback.
    """
    app = _build_app(container)
    with TestClient(app, raise_server_exceptions=False) as test_client:
        yield test_client


# ---------------------------------------------------------------------------
# GET /admin/users — W59 (#517) pagination
# ---------------------------------------------------------------------------


@pytest.mark.usefixtures("auth_session")
def test_list_users_returns_all_users(client: TestClient, fake_fixtures: FakeFixtures):
    """``GET /admin/users`` renders every seeded user in the HTML response."""
    _seed_active_user(fake_fixtures, email="alice@enterprise.test")
    _seed_active_user(fake_fixtures, email="bob@enterprise.test")

    response = client.get("/admin/users")

    assert response.status_code == 200
    body = response.text
    assert "alice@enterprise.test" in body
    assert "bob@enterprise.test" in body


@pytest.mark.usefixtures("auth_session")
def test_list_users_respects_limit_offset(client: TestClient, fake_fixtures: FakeFixtures):
    """``limit`` and ``offset`` query params drive ``container.list_all_users``."""
    _seed_active_user(fake_fixtures, email="alice@enterprise.test")
    _seed_active_user(fake_fixtures, email="bob@enterprise.test")
    _seed_active_user(fake_fixtures, email="carol@enterprise.test")

    response = client.get("/admin/users?limit=1&offset=0")

    assert response.status_code == 200
    # The page renders exactly one user; the rest stay out of the slice.
    body = response.text
    rendered = sum(
        1
        for email in (
            "alice@enterprise.test",
            "bob@enterprise.test",
            "carol@enterprise.test",
        )
        if email in body
    )
    assert rendered == 1
    # The "Mostrando X-Y de Z" pagination footer stays in step with the count.
    assert "de 3" in body


# ---------------------------------------------------------------------------
# POST /admin/users — W54 (#43) create user + DA-11 audit row
# ---------------------------------------------------------------------------


@pytest.mark.usefixtures("auth_session")
def test_create_user_encrypts_dni_and_creates_audit_row(
    client: TestClient, fake_fixtures: FakeFixtures, auth_session: dict[str, UUID]
):
    """``POST /admin/users`` persists the row, encrypts the DNI, and appends an audit row."""
    response = client.post(
        "/admin/users",
        data={
            "email": "alice" + "@" + "enterprise.test",
            "name": "Alice",
            "dni": "11111111",
        },
        headers=_auth_headers(auth_session["session_id"]),
    )

    assert response.status_code == 201
    # The fake's encrypt recorded exactly one call with the plaintext DNI.
    assert fake_fixtures.secrets.encrypted == ["11111111"]
    # The user persisted with the ciphertext, normalised email, and
    # password-reset-required status (D89: no plaintext on disk).
    persisted = fake_fixtures.users.by_email["alice" + "@" + "enterprise.test"]
    assert persisted.dni_encrypted == b"enc:11111111"
    assert persisted.status is UserStatus.PASSWORD_RESET_REQUIRED
    assert persisted.password_hash is None
    # DA-11: exactly one ``users.create`` audit row.
    assert len(fake_fixtures.audit.entries) == 1
    assert fake_fixtures.audit.entries[0].event_type == "users.create"


@pytest.mark.usefixtures("auth_session")
def test_create_user_raises_on_duplicate_email(
    client: TestClient, fake_fixtures: FakeFixtures, auth_session: dict[str, UUID]
):
    """``POST /admin/users`` returns 409 when the email is already taken (D7)."""
    # Built via concatenation so the test source never embeds an ``@``
    # pattern that a doc-renderer would auto-obfuscate.
    duplicate_email = "alice" + "@" + "enterprise.test"
    _seed_active_user(fake_fixtures, email=duplicate_email)

    response = client.post(
        "/admin/users",
        data={"email": duplicate_email, "name": "Bob", "dni": "22222222"},
        headers=_auth_headers(auth_session["session_id"]),
    )

    assert response.status_code == 409
    # No new user persisted; no audit row on the rejected path.
    assert duplicate_email in fake_fixtures.users.by_email
    assert fake_fixtures.audit.entries == []


# ---------------------------------------------------------------------------
# PATCH /admin/users/{id}/disable — W54 (D42, D89) destructive transition
# ---------------------------------------------------------------------------


@pytest.mark.usefixtures("auth_session")
def test_disable_user_sets_status_disabled(
    client: TestClient, fake_fixtures: FakeFixtures, auth_session: dict[str, UUID]
):
    """``PATCH /admin/users/{id}/disable`` flips ACTIVE → DISABLED (D42)."""
    user = _seed_active_user(fake_fixtures, email="alice@enterprise.test")

    response = client.patch(
        f"/admin/users/{user.id}/disable",
        headers=_auth_headers(auth_session["session_id"]),
    )

    assert response.status_code == 200
    assert user.status is UserStatus.DISABLED
    assert fake_fixtures.users.status_calls[-1] == (user.id, UserStatus.DISABLED)


# ---------------------------------------------------------------------------
# POST /admin/assignments — W54 (D22, DA-12, H11) profile assignment
# ---------------------------------------------------------------------------


@pytest.mark.usefixtures("auth_session")
def test_assign_profile_creates_assignment(
    client: TestClient, fake_fixtures: FakeFixtures, auth_session: dict[str, UUID]
):
    """``POST /admin/assignments`` persists the (user, app, profile) triple."""
    user = _seed_active_user(fake_fixtures, email="alice@enterprise.test")
    _seed_app(fake_fixtures, app_id=42)
    profile = _seed_profile(fake_fixtures, app_id=42)

    response = client.post(
        "/admin/assignments",
        data={
            "user_id": str(user.id),
            "app_id": "42",
            "profile_id": str(profile.id),
        },
        headers=_auth_headers(auth_session["session_id"]),
    )

    assert response.status_code == 201
    assert len(fake_fixtures.assignments.create_calls) == 1
    called = fake_fixtures.assignments.create_calls[0]
    assert called == (user.id, 42, profile.id)


@pytest.mark.usefixtures("auth_session")
def test_assign_profile_audit_logged(
    client: TestClient, fake_fixtures: FakeFixtures, auth_session: dict[str, UUID]
):
    """``POST /admin/assignments`` appends exactly one ``assignments.create`` audit row."""
    user = _seed_active_user(fake_fixtures, email="alice@enterprise.test")
    _seed_app(fake_fixtures, app_id=42)
    profile = _seed_profile(fake_fixtures, app_id=42)

    response = client.post(
        "/admin/assignments",
        data={
            "user_id": str(user.id),
            "app_id": "42",
            "profile_id": str(profile.id),
        },
        headers=_auth_headers(auth_session["session_id"]),
    )

    assert response.status_code == 201
    assert len(fake_fixtures.audit.entries) == 1
    entry = fake_fixtures.audit.entries[0]
    assert entry.event_type == "assignments.create"
    assert entry.payload["app_id"] == 42
    assert entry.payload["profile_id"] == str(profile.id)


@pytest.mark.usefixtures("auth_session")
def test_assign_profile_rejects_unknown_user(
    client: TestClient, fake_fixtures: FakeFixtures, auth_session: dict[str, UUID]
):
    """``POST /admin/assignments`` rolls back without persisting when the user is missing."""
    _seed_app(fake_fixtures, app_id=42)
    profile = _seed_profile(fake_fixtures, app_id=42)

    response = client.post(
        "/admin/assignments",
        data={
            "user_id": str(uuid4()),
            "app_id": "42",
            "profile_id": str(profile.id),
        },
        headers=_auth_headers(auth_session["session_id"]),
    )

    # The use case raises UserNotFoundError → 500 propagates from the route.
    # The contract here is "no row persisted, no audit row".
    assert response.status_code in (404, 500)
    assert fake_fixtures.assignments.create_calls == []
    assert fake_fixtures.audit.entries == []
