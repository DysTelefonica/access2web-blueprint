# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W-TEST (#519)
"""Integration tests for the Lanzadera admin HTTP routes (issue #519).

The admin HTTP routes (``list_users``, ``create_user``, ``disable_user``,
``assign_profile``) go through the full chain ``HTTP route →
LanzaderaContainer → use case → repository``. This module wires the
container with the in-memory fakes from ``tests/lanzadera/_fakes.py``
and exercises the routes end-to-end via FastAPI's ``TestClient``.

The tests cover the W54-W59 slice:
- ``GET /admin/users`` paginates via ``list_all_users`` (W59, #517).
- ``POST /admin/users`` creates the row via the ``create_user`` use case
  (W54, #43) and appends the audit row (DA-11).
- ``PATCH /admin/users/{id}/disable`` flips ACTIVE → DISABLED (W54, D42).
- ``POST /admin/assignments`` persists the (user, app, profile) triple
  via ``assign_profile`` (W54, D22, DA-12).

The ``auth_bypass`` fixture monkeypatches ``admin.require_global_admin``
so the destructive commands propagate without an auth session — the
real gate ships with M02 (A01..A03) and is out of scope for W58/W59.
"""

from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path
from uuid import uuid4

import pytest
from fastapi import APIRouter, FastAPI
from fastapi.templating import Jinja2Templates
from fastapi.testclient import TestClient

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


def _build_app(container: LanzaderaContainer) -> FastAPI:
    """Return a minimal FastAPI app with the admin router wired to ``container``.

    The lifespan and the production ``app.state.container`` dance are
    unnecessary here — the routes resolve the container via closure,
    so we mount the router directly on a fresh ``FastAPI`` instance.
    The router's prefix is ``/admin`` so the test paths match the
    production surface exactly.
    """
    app = FastAPI()
    router = APIRouter(prefix="/admin")
    templates = Jinja2Templates(directory=str(_TEMPLATES_DIR))
    register_routes(router, templates=templates, container=container)
    app.include_router(router)
    return app


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


@pytest.mark.usefixtures("auth_bypass")
def test_list_users_returns_all_users(client: TestClient, fake_fixtures: FakeFixtures):
    """``GET /admin/users`` renders every seeded user in the HTML response."""
    _seed_active_user(fake_fixtures, email="alice@enterprise.test")
    _seed_active_user(fake_fixtures, email="bob@enterprise.test")

    response = client.get("/admin/users")

    assert response.status_code == 200
    body = response.text
    assert "alice@enterprise.test" in body
    assert "bob@enterprise.test" in body


@pytest.mark.usefixtures("auth_bypass")
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
        for email in ("alice@enterprise.test", "bob@enterprise.test", "carol@enterprise.test")
        if email in body
    )
    assert rendered == 1
    # The "Mostrando X-Y de Z" pagination footer stays in step with the count.
    assert "de 3" in body


# ---------------------------------------------------------------------------
# POST /admin/users — W54 (#43) create user + DA-11 audit row
# ---------------------------------------------------------------------------


@pytest.mark.usefixtures("auth_bypass")
def test_create_user_encrypts_dni_and_creates_audit_row(
    client: TestClient, fake_fixtures: FakeFixtures
):
    """``POST /admin/users`` persists the row, encrypts the DNI, and appends an audit row."""
    response = client.post(
        "/admin/users",
        data={"email": "alice" + "@" + "enterprise.test", "name": "Alice", "dni": "11111111"},
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


@pytest.mark.usefixtures("auth_bypass")
def test_create_user_raises_on_duplicate_email(client: TestClient, fake_fixtures: FakeFixtures):
    """``POST /admin/users`` returns 409 when the email is already taken (D7)."""
    # Built via concatenation so the test source never embeds an ``@``
    # pattern that a doc-renderer would auto-obfuscate.
    duplicate_email = "alice" + "@" + "enterprise.test"
    _seed_active_user(fake_fixtures, email=duplicate_email)

    response = client.post(
        "/admin/users",
        data={"email": duplicate_email, "name": "Bob", "dni": "22222222"},
    )

    assert response.status_code == 409
    # No new user persisted; no audit row on the rejected path.
    assert duplicate_email in fake_fixtures.users.by_email
    assert fake_fixtures.audit.entries == []


# ---------------------------------------------------------------------------
# PATCH /admin/users/{id}/disable — W54 (D42, D89) destructive transition
# ---------------------------------------------------------------------------


@pytest.mark.usefixtures("auth_bypass")
def test_disable_user_sets_status_disabled(client: TestClient, fake_fixtures: FakeFixtures):
    """``PATCH /admin/users/{id}/disable`` flips ACTIVE → DISABLED (D42)."""
    user = _seed_active_user(fake_fixtures, email="alice@enterprise.test")

    response = client.patch(f"/admin/users/{user.id}/disable")

    assert response.status_code == 200
    assert user.status is UserStatus.DISABLED
    assert fake_fixtures.users.status_calls[-1] == (user.id, UserStatus.DISABLED)


# ---------------------------------------------------------------------------
# POST /admin/assignments — W54 (D22, DA-12, H11) profile assignment
# ---------------------------------------------------------------------------


@pytest.mark.usefixtures("auth_bypass")
def test_assign_profile_creates_assignment(client: TestClient, fake_fixtures: FakeFixtures):
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
    )

    assert response.status_code == 201
    assert len(fake_fixtures.assignments.create_calls) == 1
    called = fake_fixtures.assignments.create_calls[0]
    assert called == (user.id, 42, profile.id)


@pytest.mark.usefixtures("auth_bypass")
def test_assign_profile_audit_logged(client: TestClient, fake_fixtures: FakeFixtures):
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
    )

    assert response.status_code == 201
    assert len(fake_fixtures.audit.entries) == 1
    entry = fake_fixtures.audit.entries[0]
    assert entry.event_type == "assignments.create"
    assert entry.payload["app_id"] == 42
    assert entry.payload["profile_id"] == str(profile.id)


@pytest.mark.usefixtures("auth_bypass")
def test_assign_profile_rejects_unknown_user(client: TestClient, fake_fixtures: FakeFixtures):
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
    )

    # The use case raises UserNotFoundError → 500 propagates from the route.
    # The contract here is "no row persisted, no audit row".
    assert response.status_code in (404, 500)
    assert fake_fixtures.assignments.create_calls == []
    assert fake_fixtures.audit.entries == []
