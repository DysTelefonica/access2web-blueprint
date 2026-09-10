# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W62 PR-7 (#544)
"""Pytest fixtures for the W54-W62 admin integration tests (issue #519, #544).

The ``fake_fixtures`` fixture returns the bag of in-memory port fakes
the admin chain needs. The ``container`` fixture wires those fakes into
a real ``LanzaderaContainer`` so the use cases and the HTTP routes run
end-to-end against the fakes.

The ``auth_session`` fixture (W62 PR-7) seeds a global admin and a live
session row so the destructive routes that gate on
``require_global_admin`` (PR-6) can be exercised end-to-end against the
real gate rather than the W54-W59 ``auth_bypass`` no-op. Tests that need
auth pass ``auth_session`` as a fixture argument and read the
``session_id`` to set the test-only ``X-Test-Session-Id`` header that
the admin routes' test middleware (see
``tests/lanzadera/delivery/test_admin_routes_integration.py``) reads
to populate ``request.state.user_id``.

The ``auth_bypass`` fixture is kept as a deprecated alias that still
monkeypatches ``admin.require_global_admin`` — it lets the
``test_check_test_classification`` meta-test (which cites the fixture
by source string) keep passing while the destructive routes migrate
to the real gate. New tests must use ``auth_session``.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import UTC, datetime, timedelta
from typing import TYPE_CHECKING
from uuid import UUID, uuid4

import pytest

from app.src.modules.lanzadera.di.container import LanzaderaContainer
from tests.lanzadera._fakes import (
    FakeAppRepository,
    FakeAssignmentRepository,
    FakeAuditLog,
    FakeBootstrapAdminSource,
    FakeGlobalAdminRepository,
    FakeJwtSigner,
    FakePasswordHasher,
    FakePresenceRepository,
    FakeProfileRepository,
    FakeResetTokenRepository,
    FakeSecretManager,
    FakeSessionRepository,
    FakeUserRepository,
)

if TYPE_CHECKING:
    from collections.abc import Iterator


# ---------------------------------------------------------------------------
# Fakes bag — a single namespace the test files can destructure.
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class FakeFixtures:
    """Bundle of every fake the admin chain needs."""

    users: FakeUserRepository
    apps: FakeAppRepository
    assignments: FakeAssignmentRepository
    global_admins: FakeGlobalAdminRepository
    reset_tokens: FakeResetTokenRepository
    profiles: FakeProfileRepository
    audit: FakeAuditLog
    secrets: FakeSecretManager
    hasher: FakePasswordHasher
    bootstrap: FakeBootstrapAdminSource
    # W60 (#522): presence fake wired in alongside the rest of the
    # in-memory fakes so the container fixture can inject it for
    # the SSE/heartbeat use cases.
    presence: FakePresenceRepository
    # W62 (#539): session fake mirrors the rest — one per test,
    # isolated. The login use case calls ``sessions.create`` after a
    # successful password verification (PR-2); the logout use case
    # (PR-3) calls ``sessions.revoke``.
    sessions: FakeSessionRepository


@pytest.fixture
def fake_fixtures() -> FakeFixtures:
    """Return a fresh bag of fakes for every test.

    Returning a fresh bag per test keeps the fakes isolated — one
    test's mutations never leak into the next. The container fixture
    below reads from this bag, so any mutation the test performs on a
    fake is what the use case / HTTP route sees.

    ``assignments`` is wired with ``profiles`` so ``effective_permissions``
    (used by ``GET /auth/me/apps/{app_id}/capabilities``) can join
    assignments - profiles - capabilities in tests.
    """
    profiles = FakeProfileRepository()
    return FakeFixtures(
        users=FakeUserRepository(),
        apps=FakeAppRepository(),
        assignments=FakeAssignmentRepository().with_profiles(profiles),
        global_admins=FakeGlobalAdminRepository(),
        reset_tokens=FakeResetTokenRepository(),
        profiles=profiles,
        audit=FakeAuditLog(),
        secrets=FakeSecretManager(),
        hasher=FakePasswordHasher(),
        bootstrap=FakeBootstrapAdminSource(),
        # W60 (#522): presence fake — one per test, mirroring the
        # isolation the rest of the fakes already pin.
        presence=FakePresenceRepository(),
        # W62 (#539): session fake, one per test.
        sessions=FakeSessionRepository(),
    )


# ---------------------------------------------------------------------------
# Container — the composition root wired with the fakes.
# ---------------------------------------------------------------------------


@pytest.fixture
def container(fake_fixtures: FakeFixtures) -> LanzaderaContainer:
    """Return a ``LanzaderaContainer`` whose every port is the matching fake.

    ``session_factory`` stays ``None`` because no Postgres adapter is
    constructed — every port slot is occupied by the injected fake.
    The container's use-case partials are bound to the injected fakes,
    so calls to ``container.create_user(...)`` etc. dispatch into the
    in-memory fakes without ever touching SQLAlchemy.
    """
    return LanzaderaContainer(
        session_factory=None,
        secret_manager=fake_fixtures.secrets,
        password_hasher=fake_fixtures.hasher,
        bootstrap_source=fake_fixtures.bootstrap,
        user_repo=fake_fixtures.users,
        app_repo=fake_fixtures.apps,
        profile_repo=fake_fixtures.profiles,
        assignment_repo=fake_fixtures.assignments,
        global_admin_repo=fake_fixtures.global_admins,
        reset_token_repo=fake_fixtures.reset_tokens,
        # W60 (#522): presence fake injected so the SSE slice
        # never reaches a real Postgres adapter in the test path.
        presence_repo=fake_fixtures.presence,
        # W62 (#539): session fake injected so the login use case
        # never reaches a real Postgres adapter in the test path.
        session_repo=fake_fixtures.sessions,
        audit=fake_fixtures.audit,
        # W62 PR-6: jwt_signer fake so the auth routes (and the
        # AuthMiddleware that PR-6 mounts on the production app) exercise
        # a real HS256 implementation in tests rather than the env-driven
        # default builder.
        jwt_signer=FakeJwtSigner(),
    )


# ---------------------------------------------------------------------------
# Auth session — W62 PR-7 (#544) replacement for ``auth_bypass``.
# ---------------------------------------------------------------------------


@pytest.fixture(autouse=False)
def auth_session(fake_fixtures: FakeFixtures) -> Iterator[dict[str, UUID]]:
    """Seed a global admin + session so ``require_global_admin`` accepts destructive routes.

    W62 PR-7 replaces the W54-W59 ``auth_bypass`` fixture (a no-op
    monkeypatch of ``admin.require_global_admin``) with a real seeding
    fixture that mirrors the W62 auth flow:

    1. Mint a fresh ``admin_id`` and call
       ``fake_fixtures.global_admins.grant(admin_id)`` so
       ``require_global_admin``'s ``is_global_admin`` check returns
       ``True``.
    2. Mint a fresh ``session_id`` and persist a ``Session`` row in
       ``fake_fixtures.sessions`` with a 24 h ``expires_at`` (matching
       D-W62-2). The ``Session`` is inserted via the fake's own
       ``create`` coroutine so the audit trail (DA-11) and the call
       introspection both work.

    The fixture yields ``{"admin_id": <UUID>, "session_id": <UUID>}``
    so the test can send the session id in the test-only
    ``X-Test-Session-Id`` header that the admin routes' test middleware
    reads to populate ``request.state.user_id``. Tests that do not need
    auth (e.g. ``GET /admin/users`` pagination) can still opt in via
    ``@pytest.mark.usefixtures("auth_session")``; the seeded row is
    simply unused.

    Marked ``autouse=False`` so individual test files opt in via
    ``@pytest.mark.usefixtures("auth_session")`` — only the route
    tests need the seeding.
    """
    import asyncio

    from app.src.modules.lanzadera.domain.session import Session

    admin_id: UUID = uuid4()
    session_id: UUID = uuid4()
    now = datetime.now(UTC)

    async def _seed() -> None:
        await fake_fixtures.global_admins.grant(admin_id)
        session = Session(
            id=session_id,
            user_id=admin_id,
            created_at=now,
            expires_at=now + timedelta(hours=24),
        )
        await fake_fixtures.sessions.create(session)

    asyncio.run(_seed())

    yield {"admin_id": admin_id, "session_id": session_id}


# ---------------------------------------------------------------------------
# Auth bypass — DEPRECATED, retained only for the meta-test that cites it.
# ---------------------------------------------------------------------------


@pytest.fixture(autouse=False)
def auth_bypass(monkeypatch: pytest.MonkeyPatch) -> Iterator[None]:
    """DEPRECATED — use ``auth_session`` instead.

    Stub ``admin.require_global_admin`` so destructive commands propagate
    without an auth session.

    W62 PR-7 (#544) replaces this fixture with ``auth_session`` which
    seeds a real global admin + session and lets the production gate
    (``require_global_admin``) take its real decision. This no-op
    fixture is retained only so the
    ``test_check_test_classification`` meta-test (which validates the
    HR-6 string match by source) keeps passing without churn. New
    destructive-route tests must use ``@pytest.mark.usefixtures("auth_session")``
    and pass the session id in the ``X-Test-Session-Id`` header.
    """
    from app.src.modules.lanzadera.delivery.http import admin as _admin

    async def _noop_admin_gate(_request: object) -> None:
        return None

    monkeypatch.setattr(_admin, "require_global_admin", _noop_admin_gate)
    yield


__all__ = [
    "FakeFixtures",
    "auth_bypass",
    "auth_session",
    "container",
    "fake_fixtures",
]
