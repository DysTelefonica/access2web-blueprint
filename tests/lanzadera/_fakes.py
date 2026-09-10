# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W-TEST (#519)
"""In-memory fakes for the W54-W59 admin use cases (issue #519).

The admin HTTP routes (``list_users``, ``create_user``, ``disable_user``,
``assign_profile``) go through the full chain ``HTTP route →
LanzaderaContainer → use case → repository``. This module stands in for
the Postgres adapters at the repository seam so the integration tests
exercise the use-case logic end-to-end without a real database.

Why a top-level ``tests/lanzadera/_fakes.py`` (not ``delivery/_fakes.py``
or ``auth/_fakes.py``): the admin chain uses every port the container
owns, and routing imports from three of them. Splitting the fakes across
``auth/`` and ``adapters/`` would force the admin tests to reach across
two modules to build the container — which is exactly the kind of test
brittleness the W-TEST slice set out to fix. One module, one chain,
one fixture.

Design choices:

- Every fake is a plain dataclass with ``field(default_factory=...)``
  so two tests can build two independent fakes without sharing state.
- The fakes record every mutation in ``*_calls`` / ``entries`` lists so
  the assertions can pin the exact mutation count (DA-11: one audit
  row per mutation, no more, no less).
- ``FakeSecretManager`` implements the ``encrypt`` method the use cases
  call via ``functools.partial(secret_manager.encrypt)`` — that is the
  contract the production wiring pins, so the fake mirrors it.
- ``FakeUserRepository.list_all_paginated`` returns ``(rows, total)`` to
  match the W59 (#517) ``UserRepositoryPg.list_all_paginated`` contract
  the admin ``list_users`` route goes through.
"""

from __future__ import annotations

import dataclasses
from collections.abc import Sequence
from dataclasses import dataclass, field
from datetime import UTC, datetime
from typing import TYPE_CHECKING, Any
from uuid import UUID, uuid4

# Imported at module level (not under ``TYPE_CHECKING``) so ``ruff``'s
# ``F821`` rule stays happy — the dataclass field annotations reference
# these names as strings at runtime (via ``from __future__ import
# annotations``), but the static checker cannot see forward references
# defined only inside ``TYPE_CHECKING``.
from app.src.modules.lanzadera.domain.app import App, AppTopology
from app.src.modules.lanzadera.domain.assignment import Assignment
from app.src.modules.lanzadera.domain.audit_event import AuditEvent
from app.src.modules.lanzadera.domain.global_admin import GlobalAdmin
from app.src.modules.lanzadera.domain.profile import Profile
from app.src.modules.lanzadera.domain.reset_token import ResetToken
from app.src.modules.lanzadera.domain.session import Session
from app.src.modules.lanzadera.domain.user import User, UserStatus

if TYPE_CHECKING:
    from app.src.modules.lanzadera.adapters.crypto.jwt import Hs256JwtSigner

    pass  # the runtime imports above already cover the static checker

# W60 (#522): presence fake imported at module level so the rest of the
# fakes (and the conftest) can reach it without a TYPE_CHECKING dance.
from tests.lanzadera._presence_fakes import FakePresenceRepository

# ---------------------------------------------------------------------------
# User repository — DA-1, D89, W59
# ---------------------------------------------------------------------------


@dataclass
class FakeUserRepository:
    """In-memory ``UserRepository`` covering every method the admin chain hits.

    The fake mirrors the production surface: ``create`` / ``get_by_email`` /
    ``get_by_id`` / ``update_status`` for the use cases, plus
    ``list_all_paginated`` for the W59 admin ``list_users`` route. The
    full ``UserRepository`` Protocol is implemented so any test in the
    file can use the same fake.
    """

    by_id: dict[UUID, User] = field(default_factory=dict)
    by_email: dict[str, User] = field(default_factory=dict)
    create_calls: list[User] = field(default_factory=list)
    status_calls: list[tuple[UUID, UserStatus]] = field(default_factory=list)
    failed_attempts_calls: list[tuple[UUID, int]] = field(default_factory=list)
    last_login_calls: list[tuple[UUID, datetime]] = field(default_factory=list)

    def add(self, user: User) -> None:
        """Seed the fake synchronously (test-side helper, not in the Protocol)."""
        self.by_id[user.id] = user
        self.by_email[user.email] = user

    async def get_by_email(self, email: str) -> User | None:
        return self.by_email.get(email)

    async def get_by_id(self, user_id: UUID) -> User | None:
        return self.by_id.get(user_id)

    async def list_all(self) -> Sequence[User]:
        return list(self.by_id.values())

    async def list_all_paginated(
        self,
        limit: int = 50,
        offset: int = 0,
    ) -> tuple[Sequence[User], int]:
        """Mirror the W59 (#517) Postgres adapter contract.

        ``limit`` is hard-capped at 200 by the production adapter; the
        fake honours the same cap so a test that pushes 250 rows still
        gets the documented behaviour. ``offset`` is clamped to a
        non-negative value.
        """
        cap = min(limit, 200)
        rows = sorted(self.by_id.values(), key=lambda u: u.email)
        return rows[offset : offset + cap], len(rows)

    async def create(self, user: User) -> None:
        self.create_calls.append(user)
        self.add(user)

    async def update_status(self, user_id: UUID, status: UserStatus) -> None:
        self.status_calls.append((user_id, status))
        user = self.by_id[user_id]
        user.status = status

    async def update_password_and_activate(self, user_id: UUID, password_hash: str) -> None:
        user = self.by_id[user_id]
        user.password_hash = password_hash
        user.status = self._active()  # type: ignore[attr-defined]

    async def update_failed_attempts(self, user_id: UUID, failed_attempts: int) -> None:
        self.failed_attempts_calls.append((user_id, failed_attempts))
        user = self.by_id[user_id]
        user.failed_attempts = failed_attempts

    async def record_login_attempt(self, user_id: UUID, *, at: datetime) -> None:
        self.last_login_calls.append((user_id, at))
        user = self.by_id[user_id]
        user.last_login_at = at

    async def reset_failed_attempts(self, user_id: UUID) -> None:
        user = self.by_id[user_id]
        user.failed_attempts = 0

    def _active(self) -> UserStatus:
        # Imported lazily to keep the dataclass free of circular-import risk.
        from app.src.modules.lanzadera.domain.user import UserStatus

        return UserStatus.ACTIVE


# ---------------------------------------------------------------------------
# App repository — DA-7, D52, D85
# ---------------------------------------------------------------------------


@dataclass
class FakeAppRepository:
    """In-memory ``AppRepositoryPort`` covering every method the admin chain hits.

    The admin ``list_apps`` route reads via ``list_active``. The
    ``assign_profile`` use case reads via ``get_by_id`` to validate
    that the target app exists.
    """

    by_id: dict[int, App] = field(default_factory=dict)
    list_active_calls: int = 0
    get_by_id_calls: int = 0
    list_visible_to_calls: list[UUID] = field(default_factory=list)
    # W61 (#524): mutation recorders — one entry per call. ``create``
    # appends the new ``App``; ``update`` appends
    # ``(app_id, kwargs)``; ``disable`` appends ``app_id``.
    create_calls: list[App] = field(default_factory=list)
    update_calls: list[tuple[int, dict[str, object]]] = field(default_factory=list)
    disable_calls: list[int] = field(default_factory=list)
    # W61 (#524): counter the create / update / disable fake uses to
    # mint new ids. Starts at 100 to avoid colliding with the
    # seeded ``add()`` rows (the tests usually seed ``id=1``).
    next_id: int = 100

    def add(self, app: App) -> None:
        self.by_id[app.id] = app

    async def get_by_id(self, app_id: int) -> App | None:
        self.get_by_id_calls += 1
        return self.by_id.get(app_id)

    async def list_active(self) -> Sequence[App]:
        """Return every app with ``registration_status='active'``."""
        self.list_active_calls += 1
        from app.src.modules.lanzadera.domain.app import AppRegistrationStatus

        return [
            a for a in self.by_id.values() if a.registration_status is AppRegistrationStatus.ACTIVE
        ]

    async def list_visible_to(self, user_id: UUID) -> Sequence[App]:
        self.list_visible_to_calls.append(user_id)
        return list(self.by_id.values())

    async def create(
        self,
        name: str,
        short_code: str,
        deployment_topology: AppTopology,
        requires_office_presence: bool,
    ) -> App:
        """Insert a fake row; the row's ``id`` is minted off ``next_id``.

        Mirrors the production ``AppRepositoryPg.create`` contract:
        the returned ``App`` has ``registration_status='pending'``
        (the production adapter stamps it on INSERT) and the
        authoritative ``id`` / ``created_at`` / ``updated_at`` the
        route renders into the response payload.
        """
        from datetime import UTC, datetime

        from app.src.modules.lanzadera.domain.app import AppRegistrationStatus

        self.next_id += 1
        now = datetime.now(UTC)
        app = App(
            id=self.next_id,
            name=name,
            short_code=short_code,
            deployment_topology=deployment_topology,
            requires_office_presence=requires_office_presence,
            registration_status=AppRegistrationStatus.PENDING,
            created_at=now,
            updated_at=now,
        )
        self.create_calls.append(app)
        self.by_id[app.id] = app
        return app

    async def update(
        self,
        app_id: int,
        *,
        name: str | None = None,
        deployment_topology: AppTopology | None = None,
        requires_office_presence: bool | None = None,
    ) -> App:
        """Patch the matching row in-place; ``None`` fields are skipped.

        Mirrors the production ``AppRepositoryPg.update`` contract:
        the returned ``App`` carries the post-update values and a
        fresh ``updated_at`` stamp. ``RuntimeError`` if the id is
        unknown — matches the production adapter's empty-RETURNING
        signal.
        """
        from datetime import UTC, datetime

        kwargs: dict[str, object] = {}
        if name is not None:
            kwargs["name"] = name
        if deployment_topology is not None:
            kwargs["deployment_topology"] = deployment_topology
        if requires_office_presence is not None:
            kwargs["requires_office_presence"] = requires_office_presence
        self.update_calls.append((app_id, kwargs))
        existing = self.by_id.get(app_id)
        if existing is None:
            raise RuntimeError(f"app update() returned no row for id={app_id!r}")
        patched = App(
            id=existing.id,
            name=kwargs.get("name", existing.name),
            short_code=existing.short_code,
            deployment_topology=kwargs.get("deployment_topology", existing.deployment_topology),
            requires_office_presence=kwargs.get(
                "requires_office_presence", existing.requires_office_presence
            ),
            registration_status=existing.registration_status,
            created_at=existing.created_at,
            updated_at=datetime.now(UTC),
        )
        self.by_id[app_id] = patched
        return patched

    async def disable(self, app_id: int) -> App:
        """Flip the matching row to ``registration_status='retired'``.

        Mirrors the production ``AppRepositoryPg.disable`` contract.
        ``RuntimeError`` if the id is unknown.
        """
        from datetime import UTC, datetime

        from app.src.modules.lanzadera.domain.app import AppRegistrationStatus

        self.disable_calls.append(app_id)
        existing = self.by_id.get(app_id)
        if existing is None:
            raise RuntimeError(f"app disable() returned no row for id={app_id!r}")
        disabled = App(
            id=existing.id,
            name=existing.name,
            short_code=existing.short_code,
            deployment_topology=existing.deployment_topology,
            requires_office_presence=existing.requires_office_presence,
            registration_status=AppRegistrationStatus.RETIRED,
            created_at=existing.created_at,
            updated_at=datetime.now(UTC),
        )
        self.by_id[app_id] = disabled
        return disabled


# ---------------------------------------------------------------------------
# Assignment repository — DA-12, D22, D42, H11
# ---------------------------------------------------------------------------


@dataclass
class FakeAssignmentRepository:
    """In-memory ``AssignmentRepositoryPort``.

    Tracks every ``create`` call so the tests can assert the use case
    issued exactly one INSERT (DA-12 + H11: ``SinAcceso`` exclusivity
    is enforced at the application layer; the fake matches the
    lenient path used by the admin route tests).

    For ``effective_permissions`` the fake uses an injected
    ``FakeProfileRepository`` (set via ``with_profiles``) so tests
    can verify the join through assignments - profiles - capabilities.
    When no profile repo is injected, ``effective_permissions`` returns
    an empty list (the ``get_my_apps`` tests inject one).
    """

    by_id: dict[UUID, Assignment] = field(default_factory=dict)
    create_calls: list[tuple[UUID, int, UUID]] = field(default_factory=list)
    effective_permissions_calls: list[tuple[UUID, int]] = field(default_factory=list)
    _profile_repo: typing.Any = None

    def with_profiles(self, profile_repo: FakeProfileRepository) -> FakeAssignmentRepository:
        """Inject a ``FakeProfileRepository`` so ``effective_permissions`` can join."""
        self._profile_repo = profile_repo
        return self

    def add(self, assignment: Assignment) -> None:
        self.by_id[assignment.id] = assignment

    async def create(self, user_id: UUID, app_id: int, profile_id: UUID) -> Assignment:
        from app.src.modules.lanzadera.domain.assignment import Assignment

        self.create_calls.append((user_id, app_id, profile_id))
        assignment = Assignment(
            id=uuid4(),
            user_id=user_id,
            app_id=app_id,
            profile_id=profile_id,
            granted_by=None,
            granted_at=datetime.now(UTC),
            revoked_at=None,
        )
        self.by_id[assignment.id] = assignment
        return assignment

    async def list_for_user(self, user_id: UUID) -> Sequence[Assignment]:
        return [a for a in self.by_id.values() if a.user_id == user_id and a.revoked_at is None]

    async def list_for_app(self, app_id: int) -> Sequence[Assignment]:
        return [a for a in self.by_id.values() if a.app_id == app_id and a.revoked_at is None]

    async def effective_permissions(self, user_id: UUID, app_id: int) -> Sequence[str]:
        self.effective_permissions_calls.append((user_id, app_id))
        if self._profile_repo is None:
            return []
        caps: set[str] = set()
        for a in self.by_id.values():
            if a.user_id == user_id and a.app_id == app_id and a.revoked_at is None:
                profile = await self._profile_repo.get_by_id(a.profile_id)
                if profile is not None and profile.active:
                    caps.update(profile.capabilities.keys())
        return sorted(caps)


# ---------------------------------------------------------------------------
# Global admin repository — D21, D42
# ---------------------------------------------------------------------------


@dataclass
class FakeGlobalAdminRepository:
    """In-memory ``GlobalAdminRepositoryPort`` with the W54 surface.

    The fake honours the D42 last-admin invariant in ``revoke``: the
    use-case tests that exercise that path can rely on the
    ``ValueError`` the production Postgres adapter raises.
    """

    members: set[UUID] = field(default_factory=set)
    grant_calls: list[UUID] = field(default_factory=list)
    revoke_calls: list[UUID] = field(default_factory=list)

    async def there_is_any(self) -> bool:
        return bool(self.members)

    async def is_global_admin(self, user_id: UUID) -> bool:
        return user_id in self.members

    async def list_all(self) -> Sequence[GlobalAdmin]:
        from app.src.modules.lanzadera.domain.global_admin import GlobalAdmin

        return [GlobalAdmin(user_id=u) for u in self.members]

    async def grant(self, user_id: UUID) -> None:
        self.grant_calls.append(user_id)
        self.members.add(user_id)

    async def revoke(self, user_id: UUID) -> None:
        self.revoke_calls.append(user_id)
        if len(self.members) <= 1:
            raise ValueError("cannot revoke the last global admin")
        self.members.discard(user_id)


# ---------------------------------------------------------------------------
# Reset token repository — DA-4, D90
# ---------------------------------------------------------------------------


@dataclass
class FakeResetTokenRepository:
    """In-memory ``ResetTokenRepository`` (the auth-flow port).

    The admin chain does not exercise the reset-flow directly, but the
    container instantiates a reset-token repo at boot — passing a fake
    keeps the test path free of any SQLAlchemy session. ``insert`` /
    ``find_unused`` mirror the legacy ``Protocol`` shape so the fake
    stays compatible with the auth-flow tests in
    ``tests/lanzadera/auth/_fakes.py`` that import from here.
    """

    by_hash: dict[str, ResetToken] = field(default_factory=dict)
    insert_calls: list[UUID] = field(default_factory=list)

    async def insert(self, token: ResetToken) -> None:
        self.insert_calls.append(token.user_id)
        self.by_hash[token.token_hash] = token

    async def find_unused(self, token_hash: str, now: datetime) -> ResetToken | None:
        token = self.by_hash.get(token_hash)
        if token is None or token.consumed_at or token.superseded_at or token.expires_at <= now:
            return None
        return token

    async def mark_consumed(self, token_hash: str, at: datetime) -> None:
        token = self.by_hash.get(token_hash)
        if token is not None:
            self.by_hash[token_hash] = dataclasses.replace(token, consumed_at=at)

    async def mark_superseded(self, user_id: UUID, at: datetime) -> None:
        for h, row in list(self.by_hash.items()):
            if row.user_id == user_id and not row.consumed_at and not row.superseded_at:
                self.by_hash[h] = dataclasses.replace(row, superseded_at=at)


# ---------------------------------------------------------------------------
# Profile repository — DA-12, D22, D45
# ---------------------------------------------------------------------------


@dataclass
class FakeProfileRepository:
    """In-memory ``ProfileRepositoryPort`` for the assign_profile use case.

    ``assign_profile`` calls ``get_by_id`` to validate the profile
    exists and that its ``app_id`` matches the target app. Tests that
    exercise the SinAcceso exclusivity branch seed the fake with two
    profiles: a non-SinAcceso one and a SinAcceso one, and assert that
    the use case rejects the swap.
    """

    by_id: dict[UUID, Profile] = field(default_factory=dict)
    get_by_id_calls: list[UUID] = field(default_factory=list)

    def add(self, profile: Profile) -> None:
        self.by_id[profile.id] = profile

    async def get_by_id(self, profile_id: UUID) -> Profile | None:
        self.get_by_id_calls.append(profile_id)
        return self.by_id.get(profile_id)

    async def get_by_code(self, app_id: int, code: str) -> Profile | None:
        for p in self.by_id.values():
            if p.app_id == app_id and p.code == code:
                return p
        return None

    async def list_for_app(self, app_id: int) -> Sequence[Profile]:
        return [p for p in self.by_id.values() if p.app_id == app_id]

    async def create(self, profile: Profile) -> None:
        self.by_id[profile.id] = profile

    async def set_active(self, app_id: int, code: str, *, active: bool) -> None:
        for p in self.by_id.values():
            if p.app_id == app_id and p.code == code:
                p.active = active


# ---------------------------------------------------------------------------
# Audit log — DA-11, D27, D55
# ---------------------------------------------------------------------------


@dataclass
class FakeAuditLog:
    """In-memory ``AuditLog`` covering ``append`` and ``list_recent``.

    The admin route ``GET /admin/audit`` reads via ``list_recent`` on
    the Postgres adapter (``AuditLogPg.list_recent``); the fake
    implements the same method so the route works end-to-end against
    the injected fake. ``append`` records the raw ``AuditLogEntry``
    dataclass — DA-11: the production adapter stores an ``AuditEvent``
    that carries the row's UUID, but the use case emits an
    ``AuditLogEntry`` which the Postgres adapter translates under the
    hood. The fake accepts the entry shape the use case actually
    passes, which is what the contract test pins.
    """

    entries: list = field(default_factory=list)
    list_recent_calls: int = 0
    next_raises: BaseException | None = None

    async def append(self, event) -> None:
        if self.next_raises is not None:
            exc, self.next_raises = self.next_raises, None
            raise exc
        self.entries.append(event)

    async def list_recent(self, limit: int = 200) -> Sequence[AuditEvent]:
        """Mirror the production ``list_recent`` ordering (newest first)."""
        self.list_recent_calls += 1
        return list(reversed(self.entries))[:limit]

    async def list_for_actor(self, actor_id: UUID, since: datetime) -> Sequence[AuditEvent]:
        return [e for e in self.entries if e.actor_id == actor_id and e.created_at >= since]


# ---------------------------------------------------------------------------
# Secret manager — D25, D73, D11
# ---------------------------------------------------------------------------


@dataclass
class FakeSecretManager:
    """In-memory ``SecretManager`` with the ``encrypt`` method the use cases call.

    The container's ``build_use_case_factories`` passes
    ``secret_manager.encrypt`` as a ``functools.partial`` argument to
    ``create_user`` and ``bootstrap_global_admins``. The fake mirrors
    that exact attribute — the call site is ``secrets.encrypt(plaintext)
    -> bytes`` and the fake returns ``b"enc:" + plaintext.encode()``.

    The ``get`` method is included so the fake also satisfies the
    ``SecretManagerPort`` Protocol if a test asserts it.
    """

    encrypted: list[str] = field(default_factory=list)
    fail_next: bool = False

    def get(self, key: str) -> str:
        return f"fake-secret-for-{key}"

    def encrypt(self, plaintext: str) -> bytes:
        self.encrypted.append(plaintext)
        if self.fail_next:
            self.fail_next = False
            raise RuntimeError("simulated encrypt failure")
        return b"enc:" + plaintext.encode("utf-8")


# ---------------------------------------------------------------------------
# Password hasher — DA-2, D88
# ---------------------------------------------------------------------------


@dataclass
class FakePasswordHasher:
    """Identity ``PasswordHasher`` for the admin chain.

    The admin HTTP routes never invoke the hasher (the reset-flow
    chain is what hashes), but the container requires a
    ``PasswordHasher`` at construction time. The fake is included so
    the container fixture does not have to reach for
    ``CredentialHasherArgon2id`` (which spins up a real Argon2
    derivation per call — overkill for the admin tests).
    """

    hash_calls: list[str] = field(default_factory=list)
    verify_calls: list[tuple[str, str]] = field(default_factory=list)

    async def hash(self, password: str) -> str:
        self.hash_calls.append(password)
        return f"fake:{password}"

    async def verify(self, password: str, password_hash: str) -> bool:
        self.verify_calls.append((password, password_hash))
        return password_hash == f"fake:{password}"


# ---------------------------------------------------------------------------
# Bootstrap admin source — D91
# ---------------------------------------------------------------------------


@dataclass
class FakeBootstrapAdminSource:
    """In-memory ``BootstrapAdminSource`` returning a fixed email set.

    The bootstrap path is exercised by ``container.bootstrap_global_admins``;
    the admin HTTP tests do not call it, but the container is built
    once and the source is queried only at startup. Returning an empty
    list keeps the bootstrap path a no-op.
    """

    emails: list[str] = field(default_factory=list)

    def list_emails(self) -> list[str]:
        return list(self.emails)


# ---------------------------------------------------------------------------
# Session repository — W62 (D-W62-2, DA-1)
# ---------------------------------------------------------------------------


@dataclass
class FakeSessionRepository:
    """In-memory SessionRepositoryPort for the W62 auth flow.

    Mirrors the production SessionRepositoryPg contract: create
    stores the row verbatim, get_by_id returns the row regardless
    of expiry (the caller is responsible for the expiry check), and
    revoke shortens expires_at to a past timestamp without
    deleting the row (DA-11 audit-consumer visibility).
    """

    sessions: dict[UUID, Session] = field(default_factory=dict)
    create_calls: list[UUID] = field(default_factory=list)
    revoke_calls: list[UUID] = field(default_factory=list)

    async def create(self, session: Session) -> Session:
        self.sessions[session.id] = session
        self.create_calls.append(session.id)
        return session

    async def get_by_id(self, session_id: UUID) -> Session | None:
        return self.sessions.get(session_id)

    async def revoke(self, session_id: UUID, at: datetime) -> bool:
        if session_id not in self.sessions:
            return False
        existing = self.sessions[session_id]
        new_expires = at if at < existing.expires_at else existing.expires_at
        self.sessions[session_id] = dataclasses.replace(existing, expires_at=new_expires)
        self.revoke_calls.append(session_id)
        return True


# ---------------------------------------------------------------------------
# JWT signer — D-W62-1 (PR-4)
# ---------------------------------------------------------------------------


@dataclass
# ---------------------------------------------------------------------------
# JWT signer — D-W62-1 (PR-4)
# ---------------------------------------------------------------------------


@dataclass
class FakeJwtSigner:
    """In-memory ``JwtSignerPort`` for the W62 auth flow.

    Wraps the production ``Hs256JwtSigner`` (with a fixed 32-byte
    secret) so the test path behaves identically to production
    while still recording every ``sign`` and ``verify`` call for
    introspection (e.g. count the rounds of a token-rotation test,
    or assert the middleware never re-signs).
    """

    _signer: Hs256JwtSigner
    sign_calls: list[Any] = field(default_factory=list)
    verify_calls: list[tuple[str, int]] = field(default_factory=list)

    def __init__(self) -> None:
        # Constructed manually (not via @dataclass __init__) so the
        # ``Hs256JwtSigner`` instance is built lazily without needing
        # ``__future__.annotations`` to forward-reference the field.
        # The runtime import lives inside the method to avoid a circular
        # dependency at module load time (``crypto.jwt`` does not import
        # from ``_fakes``; this is one-directional).
        from app.src.modules.lanzadera.adapters.crypto.jwt import (
            Hs256JwtSigner as _Hs256JwtSigner,
        )

        self._signer = _Hs256JwtSigner(b"x" * 32)
        self.sign_calls = []
        self.verify_calls = []

    def sign(self, payload: dict[str, Any]) -> str:
        self.sign_calls.append(dict(payload))
        return self._signer.sign(payload)

    def verify(self, token: str, *, now: int) -> dict[str, Any]:
        self.verify_calls.append((token, now))
        return self._signer.verify(token, now=now)
        return self._signer.verify(token, now=now)


__all__ = [
    "FakeAppRepository",
    "FakeAssignmentRepository",
    "FakeAuditLog",
    "FakeBootstrapAdminSource",
    "FakeGlobalAdminRepository",
    "FakeJwtSigner",
    "FakePasswordHasher",
    "FakePresenceRepository",
    "FakeProfileRepository",
    "FakeResetTokenRepository",
    "FakeSecretManager",
    "FakeSessionRepository",
    "FakeUserRepository",
]
