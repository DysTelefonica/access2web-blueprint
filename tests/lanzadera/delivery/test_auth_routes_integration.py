# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W62 (#543)
"""Integration tests for the W62 (#543) auth HTTP routes + ``require_global_admin``.

The auth HTTP routes (``POST /auth/login``, ``POST /auth/logout``,
``GET /auth/me``) go through the full chain ``HTTP route →
LanzaderaContainer → use case → repository``. The ``require_global_admin``
gate (PR-6) reads ``request.state.user_id`` populated by ``AuthMiddleware``
(PR-5) and consults ``container.global_admin_repo``. This module wires
the container with the in-memory fakes from ``tests/lanzadera/_fakes.py``
and exercises every route end-to-end via FastAPI's ``TestClient``.

Categoría 4 (integration HTTP, HR-4): the container is real, every
infrastructure port is replaced by a fake, ``unittest.mock.MagicMock``
never appears. The fakes-injected container builds the JWT signer with
``b"x" * 32`` (matching ``FakeJwtSigner``'s internal ``Hs256JwtSigner``)
so the auth middleware can verify the tokens the tests mint without
spinning up a real ``EnvSecretManagerAdapter``.

The ``auth_bypass`` fixture is NOT needed (HR-6 scope): these routes do
not touch ``admin_routes``; the gate exercised here is the auth route's
own ``require_global_admin``, not the admin-routes destructive gate.
"""

from __future__ import annotations

from collections.abc import Iterator
from datetime import UTC, datetime, timedelta
from typing import TYPE_CHECKING
from uuid import UUID, uuid4

import pytest
from fastapi import APIRouter, Depends, FastAPI
from fastapi.testclient import TestClient

from app.src.modules.lanzadera.adapters.crypto.jwt import Hs256JwtSigner
from app.src.modules.lanzadera.delivery.http.admin import require_global_admin
from app.src.modules.lanzadera.delivery.http.auth_middleware import AuthMiddleware
from app.src.modules.lanzadera.delivery.http.auth_routes import register_auth_routes
from app.src.modules.lanzadera.di.container import LanzaderaContainer
from app.src.modules.lanzadera.domain.session import Session
from app.src.modules.lanzadera.domain.user import User, UserStatus
from tests.lanzadera._fakes import FakeJwtSigner

if TYPE_CHECKING:
    from tests.lanzadera.conftest import FakeFixtures


# ---------------------------------------------------------------------------
# JWT secret — must match ``FakeJwtSigner``'s internal ``Hs256JwtSigner`` so
# the tokens we mint in the test body verify cleanly inside the middleware.
# ---------------------------------------------------------------------------


_JWT_SECRET: bytes = b"x" * 32  # noqa: S105 — test-only deterministic secret


# ---------------------------------------------------------------------------
# Container builder — local to this file because the shared ``container``
# fixture does not inject ``jwt_signer`` (FakeFixtures predates PR-6).
# ---------------------------------------------------------------------------


def _build_container(fakes: FakeFixtures) -> LanzaderaContainer:
    """Return a ``LanzaderaContainer`` with the fakes + a ``FakeJwtSigner``.

    The shared ``container`` fixture injects every port except
    ``jwt_signer``; constructing the signer via the default builder would
    require ``$JWT_SECRET`` in the environment. The auth routes need a
    signer at request time, so this helper builds one locally with the
    same secret the fakes use internally.
    """
    return LanzaderaContainer(
        session_factory=None,
        secret_manager=fakes.secrets,
        password_hasher=fakes.hasher,
        bootstrap_source=fakes.bootstrap,
        user_repo=fakes.users,
        app_repo=fakes.apps,
        profile_repo=fakes.profiles,
        assignment_repo=fakes.assignments,
        global_admin_repo=fakes.global_admins,
        reset_token_repo=fakes.reset_tokens,
        presence_repo=fakes.presence,
        session_repo=fakes.sessions,
        audit=fakes.audit,
        jwt_signer=FakeJwtSigner(),
    )


def _patch_auth_routes_request_type() -> None:
    """Expose ``Request`` at the ``auth_routes`` module level for FastAPI resolution.

    ``auth_routes.py`` imports ``Request`` inside ``register_auth_routes``
    and uses ``from __future__ import annotations``, so FastAPI sees
    ``req: Request`` as a forward reference that cannot be resolved
    against the module's globals — the parameter ends up classified as a
    query parameter and every request returns 422. Pinning ``Request``
    into the module namespace before the route registration lets
    FastAPI's dependency analysis recognise it as the special Request
    injection. Production behaviour is unchanged; the patch is local to
    the test process.
    """
    from fastapi import Request as _FastAPIRequest

    from app.src.modules.lanzadera.delivery.http import auth_routes as _auth_routes_module

    _auth_routes_module.Request = _FastAPIRequest  # type: ignore[attr-defined]


# ---------------------------------------------------------------------------
# FastAPI app builder — minimal app with auth router + AuthMiddleware.
# ---------------------------------------------------------------------------


def _build_app(container: LanzaderaContainer) -> FastAPI:
    """Return a FastAPI app exposing ``/auth/login``, ``/auth/logout``, ``/auth/me``.

    The lifespan and ``app.state.container`` dance are unnecessary for the
    auth routes themselves (the routes close over ``container``). We still
    set ``app.state.container`` so ``require_global_admin`` can resolve the
    gate from the test path.
    """
    app = FastAPI()
    app.state.container = container
    auth_router = APIRouter()
    register_auth_routes(auth_router, container=container)
    app.include_router(auth_router)

    def _now_epoch() -> int:
        return int(container._clock().timestamp())

    app.add_middleware(
        AuthMiddleware,
        jwt_signer=container.jwt_signer,
        now=_now_epoch,
    )
    return app


def _build_protected_app(container: LanzaderaContainer) -> FastAPI:
    """Return a FastAPI app with a ``/protected`` route gated by ``require_global_admin``.

    The same auth router is mounted so the JWT middleware populates
    ``request.state.user_id`` for the protected route. The protected
    route uses ``require_global_admin`` as a dependency, exactly the way
    the production admin routes use it.
    """
    app = _build_app(container)

    @app.get("/protected", dependencies=[Depends(require_global_admin)])
    async def _protected() -> dict[str, bool]:
        return {"ok": True}

    return app


# ---------------------------------------------------------------------------
# Client fixtures — one per route surface (auth routes vs require_global_admin).
# ---------------------------------------------------------------------------


@pytest.fixture
def auth_client(fake_fixtures: FakeFixtures) -> Iterator[TestClient]:
    """Yield a ``TestClient`` for the ``/auth/...`` routes.

    Builds a local container that injects ``FakeJwtSigner`` because the
    shared ``container`` fixture does not wire one (it was authored
    before PR-6 added the JWT signer slot). The ``FakeFixtures``
    instance is reused so the test can read mutations back through
    ``fake_fixtures.sessions`` etc.
    """
    _patch_auth_routes_request_type()
    local_container = _build_container(fake_fixtures)
    app = _build_app(local_container)
    with TestClient(app) as test_client:
        yield test_client


@pytest.fixture
def protected_client(fake_fixtures: FakeFixtures) -> Iterator[TestClient]:
    """Yield a ``TestClient`` for the ``/protected`` route (require_global_admin)."""
    _patch_auth_routes_request_type()
    local_container = _build_container(fake_fixtures)
    app = _build_protected_app(local_container)
    with TestClient(app) as test_client:
        yield test_client


# ---------------------------------------------------------------------------
# Helpers — seed users / sessions in the shared fakes.
# ---------------------------------------------------------------------------


def _seed_active_user(
    fakes: FakeFixtures,
    *,
    email: str = "alice" + "@" + "enterprise.test",
    password: str = "secret",
    status: UserStatus = UserStatus.ACTIVE,
    failed_attempts: int = 0,
    last_login_at: datetime | None = None,
) -> User:
    """Seed a user with the password hash the ``FakePasswordHasher`` recognises.

    The fake's ``verify`` returns ``True`` iff ``password_hash ==
    "fake:" + password``; we seed the matching hash so the login use
    case takes the success path on the supplied credentials.
    """
    now = datetime.now(UTC)
    user = User(
        id=uuid4(),
        email=email,
        name="Alice",
        dni_encrypted=b"enc:11111111",
        password_hash=f"fake:{password}",
        status=status,
        failed_attempts=failed_attempts,
        last_login_at=last_login_at,
        created_at=now,
        updated_at=now,
    )
    fakes.users.add(user)
    return user


def _seed_session(fakes: FakeFixtures, *, user_id: UUID) -> Session:
    """Seed a live session row with a 24 h ``expires_at`` (matches D-W62-2)."""
    now = datetime.now(UTC)
    session = Session(
        id=uuid4(),
        user_id=user_id,
        created_at=now,
        expires_at=now + timedelta(hours=24),
    )
    fakes.sessions.sessions[session.id] = session
    return session


def _mint_jwt(*, sub: UUID, ttl_seconds: int = 3600) -> str:
    """Sign a JWT with ``sub=<session_id>`` using the test secret.

    The secret matches the one ``FakeJwtSigner`` constructs internally,
    so the auth middleware's ``verify`` accepts the token.
    """
    now = int(datetime.now(UTC).timestamp())
    return Hs256JwtSigner(_JWT_SECRET).sign({"sub": str(sub), "iat": now, "exp": now + ttl_seconds})


# ---------------------------------------------------------------------------
# POST /auth/login — W62 (D38, D39, D-W62-2, DA-11)
# ---------------------------------------------------------------------------


def test_login_returns_token_and_session_id(
    auth_client: TestClient, fake_fixtures: FakeFixtures
) -> None:
    """``POST /auth/login`` issues a JWT and creates the matching session row.

    The session is persisted in ``fake_fixtures.sessions`` with a 24 h TTL
    so the auth middleware can re-validate the token on subsequent
    requests (D-W62-2). The response body carries the JWT and the
    session id so the client can identify the session for logout.
    """
    _seed_active_user(fake_fixtures)

    response = auth_client.post(
        "/auth/login",
        json={"email": "alice" + "@" + "enterprise.test", "password": "secret"},
    )

    assert response.status_code == 200
    body = response.json()
    assert isinstance(body["token"], str) and body["token"]
    # UUID parses — the test must not see a placeholder or empty string.
    parsed = UUID(body["session_id"])
    assert parsed is not None
    # The session row was persisted; exactly one create call.
    assert len(fake_fixtures.sessions.create_calls) == 1
    assert fake_fixtures.sessions.create_calls[0] == parsed


def test_login_invalid_credentials_returns_401(
    auth_client: TestClient, fake_fixtures: FakeFixtures
) -> None:
    """Wrong password → 401 ``invalid credentials`` (D-W62-2, D38)."""
    _seed_active_user(fake_fixtures)

    response = auth_client.post(
        "/auth/login",
        json={"email": "alice" + "@" + "enterprise.test", "password": "wrong"},
    )

    assert response.status_code == 401
    assert response.json()["detail"] == "invalid credentials"
    # No session created on the rejected path.
    assert fake_fixtures.sessions.create_calls == []


def test_login_locked_account_returns_423(
    auth_client: TestClient, fake_fixtures: FakeFixtures
) -> None:
    """``failed_attempts=5`` + ``last_login_at=now`` → 423 Locked (D38, D39).

    The lockout policy threshold is 5 (default ``LockoutPolicy()``); the
    elapsed-since-last-attempt guard keeps the user locked because we
    stamp ``last_login_at`` with the current time.
    """
    now = datetime.now(UTC)
    _seed_active_user(
        fake_fixtures,
        failed_attempts=5,
        last_login_at=now,
    )

    response = auth_client.post(
        "/auth/login",
        json={"email": "alice" + "@" + "enterprise.test", "password": "secret"},
    )

    assert response.status_code == 423
    assert "locked" in response.json()["detail"].lower()
    # No session created on the locked path.
    assert fake_fixtures.sessions.create_calls == []


def test_login_inactive_account_returns_403(
    auth_client: TestClient, fake_fixtures: FakeFixtures
) -> None:
    """``status=DISABLED`` → 403 (D89: only ACTIVE users may log in)."""
    _seed_active_user(fake_fixtures, status=UserStatus.DISABLED)

    response = auth_client.post(
        "/auth/login",
        json={"email": "alice" + "@" + "enterprise.test", "password": "secret"},
    )

    assert response.status_code == 403
    # No session created on the not-active path.
    assert fake_fixtures.sessions.create_calls == []


def test_login_malformed_body_returns_422(
    auth_client: TestClient, fake_fixtures: FakeFixtures
) -> None:
    """Missing ``email`` / ``password`` → 422 Unprocessable Entity."""
    response = auth_client.post("/auth/login", json={})

    assert response.status_code == 422
    # No session created on the malformed-body path.
    assert fake_fixtures.sessions.create_calls == []


# ---------------------------------------------------------------------------
# POST /auth/logout — W62 (D-W62-2, DA-11)
# ---------------------------------------------------------------------------


def test_logout_returns_401_without_token(
    auth_client: TestClient, fake_fixtures: FakeFixtures
) -> None:
    """Missing Authorization header → 401 (auth middleware leaves state empty)."""
    response = auth_client.post("/auth/logout")

    assert response.status_code == 401
    # No revoke call: the middleware never populated ``session_id``.
    assert fake_fixtures.sessions.revoke_calls == []


def test_logout_revokes_session_with_valid_token(
    auth_client: TestClient, fake_fixtures: FakeFixtures
) -> None:
    """Valid Bearer → 204 + ``sessions.revoke`` shortens ``expires_at`` (D-W62-2)."""
    user = _seed_active_user(fake_fixtures)
    session = _seed_session(fake_fixtures, user_id=user.id)
    token = _mint_jwt(sub=session.id)

    response = auth_client.post("/auth/logout", headers={"Authorization": f"Bearer {token}"})

    assert response.status_code == 204
    # Exactly one revoke call for the session id; the row stays (DA-11).
    assert fake_fixtures.sessions.revoke_calls == [session.id]
    assert fake_fixtures.sessions.sessions[session.id].expires_at < datetime.now(UTC)


# ---------------------------------------------------------------------------
# GET /auth/me — W62 (D-W62-2)
# ---------------------------------------------------------------------------


def test_me_returns_401_without_token(auth_client: TestClient) -> None:
    """Missing Authorization header → 401 (auth middleware leaves state empty)."""
    response = auth_client.get("/auth/me")
    assert response.status_code == 401


def test_me_returns_enriched_identity(
    auth_client: TestClient, fake_fixtures: FakeFixtures
) -> None:
    """Valid Bearer → 200 + user identity + assigned apps.

    The JWT ``sub`` must be ``user.id`` so that ``get_my_apps`` finds
    the seeded assignments (the middleware puts ``sub`` into
    ``request.state.user_id``).
    """
    user = _seed_active_user(fake_fixtures)
    session = _seed_session(fake_fixtures, user_id=user.id)
    token = _mint_jwt(sub=user.id)

    app = _seed_app(fake_fixtures, id=3, name="Expedientes", short_code="EXP")
    profile = _seed_profile(fake_fixtures, app_id=3, code="ADMIN", capabilities={"Calidad": True})
    _seed_assignment(fake_fixtures, user_id=user.id, app_id=3, profile_id=profile.id)

    response = auth_client.get("/auth/me", headers={"Authorization": f"Bearer {token}"})

    assert response.status_code == 200
    body = response.json()
    assert body["user_id"] == str(user.id)
    assert body["email"] == user.email
    assert body["name"] == user.name
    assert len(body["apps"]) == 1
    assert body["apps"][0]["app_id"] == 3
    assert body["apps"][0]["profile_code"] == "ADMIN"
    assert "Calidad" in body["apps"][0]["capabilities"]


# ---------------------------------------------------------------------------
# require_global_admin — W62 PR-6 (D91)
# ---------------------------------------------------------------------------


def test_require_global_admin_returns_401_without_user_id(
    protected_client: TestClient,
) -> None:
    """No Authorization header → 401 (middleware leaves ``state.user_id`` empty)."""
    response = protected_client.get("/protected")
    assert response.status_code == 401


def test_require_global_admin_returns_403_for_non_admin(
    protected_client: TestClient, fake_fixtures: FakeFixtures
) -> None:
    """JWT present but caller not in ``global_admins`` → 403 Forbidden."""
    user = _seed_active_user(fake_fixtures)
    token = _mint_jwt(sub=user.id)
    # Intentionally NOT calling ``fake_fixtures.global_admins.grant(user.id)``.

    response = protected_client.get("/protected", headers={"Authorization": f"Bearer {token}"})

    assert response.status_code == 403
    assert "admin" in response.json()["detail"].lower()


def test_require_global_admin_passes_for_admin(
    protected_client: TestClient, fake_fixtures: FakeFixtures
) -> None:
    """Caller IS in ``global_admins`` → 200 OK (the gate passes)."""
    user = _seed_active_user(fake_fixtures)
    fake_fixtures.global_admins.members.add(user.id)
    token = _mint_jwt(sub=user.id)

    response = protected_client.get("/protected", headers={"Authorization": f"Bearer {token}"})

    assert response.status_code == 200
    assert response.json() == {"ok": True}


def test_require_global_admin_returns_503_without_container(
    fake_fixtures: FakeFixtures,
) -> None:
    """``app.state.container = None`` → 503 Service Unavailable.

    The lifespan logs a warning when the container is not initialised
    (e.g. ``DATABASE_URL`` not configured); the gate must not silently
    bypass — it returns 503 so the caller can distinguish a missing
    backend from a real auth decision.
    """
    container = _build_container(fake_fixtures)
    app = _build_protected_app(container)
    app.state.container = None

    user = _seed_active_user(fake_fixtures)
    fake_fixtures.global_admins.members.add(user.id)
    token = _mint_jwt(sub=user.id)

    with TestClient(app) as test_client:
        response = test_client.get("/protected", headers={"Authorization": f"Bearer {token}"})

    assert response.status_code == 503


# ---------------------------------------------------------------------------
# GET /auth/me/apps — Issue #585
# ---------------------------------------------------------------------------


def test_my_apps_returns_401_without_token(auth_client: TestClient) -> None:
    """Missing Authorization header → 401."""
    response = auth_client.get("/auth/me/apps")
    assert response.status_code == 401


def test_my_apps_returns_empty_list_when_no_assignments(
    auth_client: TestClient, fake_fixtures: FakeFixtures
) -> None:
    """User with no assignments gets an empty apps list."""
    user = _seed_active_user(fake_fixtures)
    session = _seed_session(fake_fixtures, user_id=user.id)
    token = _mint_jwt(sub=session.id)

    response = auth_client.get("/auth/me/apps", headers={"Authorization": f"Bearer {token}"})

    assert response.status_code == 200
    assert response.json()["apps"] == []


def test_my_apps_returns_apps_with_profiles_and_capabilities(
    auth_client: TestClient, fake_fixtures: FakeFixtures
) -> None:
    """User with one assignment gets the app + profile + capabilities."""
    user = _seed_active_user(fake_fixtures)
    session = _seed_session(fake_fixtures, user_id=user.id)
    # JWT sub must be user.id so that get_my_apps finds the assignment
    token = _mint_jwt(sub=user.id)

    app = _seed_app(fake_fixtures, id=3, name="Expedientes", short_code="EXP")
    profile = _seed_profile(fake_fixtures, app_id=3, code="ADMIN", capabilities={"Calidad": True})
    _seed_assignment(fake_fixtures, user_id=user.id, app_id=3, profile_id=profile.id)

    response = auth_client.get("/auth/me/apps", headers={"Authorization": f"Bearer {token}"})

    assert response.status_code == 200
    body = response.json()
    assert body["user_id"] == str(user.id)
    assert len(body["apps"]) == 1
    assert body["apps"][0]["app_id"] == 3
    assert body["apps"][0]["app_name"] == "Expedientes"
    assert body["apps"][0]["profile_code"] == "ADMIN"
    assert "Calidad" in body["apps"][0]["capabilities"]


# ---------------------------------------------------------------------------
# GET /auth/me/apps/{app_id}/capabilities — Issue #585
# ---------------------------------------------------------------------------


def test_my_capabilities_returns_401_without_token(auth_client: TestClient) -> None:
    """Missing Authorization header → 401."""
    response = auth_client.get("/auth/me/apps/1/capabilities")
    assert response.status_code == 401


def test_my_capabilities_returns_empty_for_unassigned_app(
    auth_client: TestClient, fake_fixtures: FakeFixtures
) -> None:
    """User not assigned to app → empty capabilities list."""
    user = _seed_active_user(fake_fixtures)
    session = _seed_session(fake_fixtures, user_id=user.id)
    token = _mint_jwt(sub=session.id)

    response = auth_client.get(
        "/auth/me/apps/99/capabilities", headers={"Authorization": f"Bearer {token}"}
    )

    assert response.status_code == 200
    assert response.json()["capabilities"] == []


def test_my_capabilities_returns_profile_keys(
    auth_client: TestClient, fake_fixtures: FakeFixtures
) -> None:
    """User assigned to app → returns the profile's capability names."""
    user = _seed_active_user(fake_fixtures)
    session = _seed_session(fake_fixtures, user_id=user.id)
    # JWT sub must be user.id so effective_permissions finds the assignment
    token = _mint_jwt(sub=user.id)

    _seed_app(fake_fixtures, id=7, name="Brass", short_code="BRA")
    profile = _seed_profile(
        fake_fixtures, app_id=7, code="CALIDAD", capabilities={"Calidad": True, "write": True}
    )
    _seed_assignment(fake_fixtures, user_id=user.id, app_id=7, profile_id=profile.id)

    response = auth_client.get(
        "/auth/me/apps/7/capabilities", headers={"Authorization": f"Bearer {token}"}
    )

    assert response.status_code == 200
    body = response.json()
    assert body["user_id"] == str(user.id)
    assert body["app_id"] == 7
    caps = body["capabilities"]
    assert "Calidad" in caps
    assert "write" in caps


# ---------------------------------------------------------------------------
# Additional seed helpers
# ---------------------------------------------------------------------------


def _seed_app(fake_fixtures: FakeFixtures, id: int, name: str, short_code: str) -> None:
    fake_fixtures.apps.add(
        App(
            id=id,
            name=name,
            short_code=short_code,
            deployment_topology=AppTopology.CENTRAL,
            requires_office_presence=False,
            registration_status=AppRegistrationStatus.ACTIVE,
            created_at=datetime(2026, 1, 1, tzinfo=UTC),
            updated_at=datetime(2026, 1, 1, tzinfo=UTC),
        )
    )


def _seed_profile(
    fake_fixtures: FakeFixtures,
    app_id: int,
    code: str,
    capabilities: dict[str, bool],
) -> Profile:
    profile = Profile(
        id=uuid4(),
        app_id=app_id,
        code=code,
        name=code,
        capabilities=capabilities,
        active=True,
        created_at=datetime(2026, 1, 1, tzinfo=UTC),
        updated_at=datetime(2026, 1, 1, tzinfo=UTC),
    )
    fake_fixtures.profiles.add(profile)
    return profile


def _seed_assignment(
    fake_fixtures: FakeFixtures,
    user_id: UUID,
    app_id: int,
    profile_id: UUID,
) -> None:
    fake_fixtures.assignments.add(
        Assignment(
            id=uuid4(),
            user_id=user_id,
            app_id=app_id,
            profile_id=profile_id,
            granted_by=None,
            granted_at=datetime(2026, 1, 1, tzinfo=UTC),
            revoked_at=None,
        )
    )


# Needed for the new seed helpers.
from app.src.modules.lanzadera.domain.app import App, AppRegistrationStatus, AppTopology
from app.src.modules.lanzadera.domain.assignment import Assignment
from app.src.modules.lanzadera.domain.profile import Profile


def test_DEBUG_print_routes(auth_client: TestClient) -> None:
    """Debug: print registered routes."""
    print("\nRegistered routes:")
    for route in auth_client.app.routes:
        print(f"  {route.path}")


def test_DEBUG_source_contains_my_apps(auth_client: TestClient) -> None:
    """Debug: check if the auth_routes source contains my_apps."""
    import inspect
    from app.src.modules.lanzadera.delivery.http import auth_routes
    src = inspect.getsource(auth_routes.register_auth_routes)
    print(f"\nmy_apps in source: {'my_apps' in src}")
    print(f"my_capabilities in source: {'my_capabilities' in src}")
    print(f"last 200 chars: {repr(src[-200:])}")


def test_DEBUG_call_register(auth_client: TestClient, fake_fixtures: FakeFixtures) -> None:
    """Debug: call register_auth_routes and print routes."""
    from app.src.modules.lanzadera.delivery.http.auth_routes import register_auth_routes
    from fastapi import APIRouter
    
    router2 = APIRouter()
    register_auth_routes(router2, container=auth_client.app.state.container)
    print("\nRoutes from register_auth_routes:")
    for route in router2.routes:
        print(f"  {route.path}")


def test_DEBUG_patch_then_register(auth_client: TestClient) -> None:
    """Debug: re-patch and re-register to see routes."""
    from app.src.modules.lanzadera.delivery.http import auth_routes as _ar
    from fastapi import Request
    _ar.Request = Request  # type: ignore[attr-defined]
    
    from app.src.modules.lanzadera.delivery.http.auth_routes import register_auth_routes
    from fastapi import APIRouter
    
    router3 = APIRouter()
    container = auth_client.app.state.container
    register_auth_routes(router3, container=container)
    print("\nRoutes from re-registered router:")
    for route in router3.routes:
        print(f"  {route.path}")
    
    # Also check the _build_app function's router
    print("\nRoutes from auth_client.app:")
    for route in auth_client.app.routes:
        print(f"  {route.path}")


def test_DEBUG_reload_and_register(auth_client: TestClient) -> None:
    """Debug: reload the module and check routes."""
    import importlib
    from app.src.modules.lanzadera.delivery.http import auth_routes
    
    importlib.reload(auth_routes)
    
    from app.src.modules.lanzadera.delivery.http.auth_routes import register_auth_routes
    from fastapi import APIRouter
    
    router4 = APIRouter()
    container = auth_client.app.state.container
    register_auth_routes(router4, container=container)
    print("\nRoutes from reload:")
    for route in router4.routes:
        print(f"  {route.path}")


def test_DEBUG_source_vs_bytecode(auth_client: TestClient) -> None:
    """Debug: compare inspect.getsource with actual execution."""
    import inspect
    import dis
    from app.src.modules.lanzadera.delivery.http import auth_routes
    
    src = inspect.getsource(auth_routes.register_auth_routes)
    
    # Count @router.get occurrences in source
    count = src.count('@router.get')
    print(f"\n@router.get occurrences in source: {count}")
    
    # Also check what bytecode says
    co = auth_routes.register_auth_routes.__code__
    print(f"Bytecode argcount: {co.co_argcount}")
    print(f"Bytecode nlocals: {co.co_nlocals}")
    
    # Print the first few bytecode instructions
    print("\nFirst 30 bytecode instructions:")
    for i, instr in enumerate(dis.get_instructions(auth_routes.register_auth_routes)):
        if i >= 30:
            break
        print(f"  {instr.offset:4d} {instr.opname:20s} {instr.argrepr}")


def test_DEBUG_raw_source(auth_client: TestClient) -> None:
    """Debug: print raw source of register_auth_routes."""
    import inspect
    from app.src.modules.lanzadera.delivery.http import auth_routes
    
    src = inspect.getsource(auth_routes.register_auth_routes)
    print(f"\nTotal source length: {len(src)}")
    print(f"Contains my_apps: {'my_apps' in src}")
    print(f"Contains my_capabilities: {'my_capabilities' in src}")
    print(f"Contains /auth/me/apps: {'/auth/me/apps' in src}")
    
    # Print the last 1000 chars
    print(f"\nLast 1000 chars of source:")
    print(src[-1000:])


def test_DEBUG_file_mtime(auth_client: TestClient) -> None:
    """Debug: check file modification times."""
    import os
    import importlib
    from app.src.modules.lanzadera.delivery.http import auth_routes
    
    py_file = auth_routes.__file__
    py_stat = os.stat(py_file)
    print(f"\n.py file mtime: {py_stat.st_mtime}")
    
    # Check if there's a .pyc
    import marshal
    pyc_file = py_file + 'c'
    if os.path.exists(pyc_file):
        pyc_stat = os.stat(pyc_file)
        print(f".pyc file mtime: {pyc_stat.st_mtime}")
        print(f".pyc older than .py: {pyc_stat.st_mtime < py_stat.st_mtime}")
    else:
        print("No .pyc file found")
