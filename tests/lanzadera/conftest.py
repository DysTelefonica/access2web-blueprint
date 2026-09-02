# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W-TEST (#519)
"""Pytest fixtures for the W54-W59 admin integration tests (issue #519).

The ``fake_fixtures`` fixture returns the bag of in-memory port fakes
the admin chain needs. The ``container`` fixture wires those fakes into
a real ``LanzaderaContainer`` so the use cases and the HTTP routes run
end-to-end against the fakes. The ``auth_bypass`` fixture monkeypatches
``admin.require_global_admin`` so the destructive commands propagate
without an auth session in the test process.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

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
    """
    return FakeFixtures(
        users=FakeUserRepository(),
        apps=FakeAppRepository(),
        assignments=FakeAssignmentRepository(),
        global_admins=FakeGlobalAdminRepository(),
        reset_tokens=FakeResetTokenRepository(),
        profiles=FakeProfileRepository(),
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
# Auth bypass — disable the D91 global-admin gate for the destructive routes.
# ---------------------------------------------------------------------------


@pytest.fixture(autouse=False)
def auth_bypass(monkeypatch: pytest.MonkeyPatch) -> Iterator[None]:
    """Stub ``admin.require_global_admin`` so destructive commands propagate.

    The HTTP routes call ``_admin.require_global_admin()`` at the top of
    the destructive handlers. The placeholder raises nothing in the
    test process (no auth wiring exists in the unit-test path), but
    stubbing it explicitly documents the intent and lets the destructive
    tests run without depending on the default no-op behaviour.

    Marked ``autouse=False`` so individual test files opt in via
    ``@pytest.mark.usefixtures("auth_bypass")`` — only the route tests
    need the bypass.
    """
    from app.src.modules.lanzadera.delivery.http import admin as _admin

    async def _noop_admin_gate(_request: object) -> None:
        return None

    monkeypatch.setattr(_admin, "require_global_admin", _noop_admin_gate)
    yield


__all__ = [
    "FakeFixtures",
    "auth_bypass",
    "container",
    "fake_fixtures",
]
