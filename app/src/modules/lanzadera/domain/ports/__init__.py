# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Protocol definitions for driven ports.

Purity rule: ports may import `domain` and other ports only. Frameworks are
forbidden here (DA-1).

W54 (#508) extended the `UserRepository`, `ProfileRepository`,
`AppRepositoryPort`, `AssignmentRepositoryPort`, and
`GlobalAdminRepository` Protocols to surface the admin-side operations
the new use cases call. W55 (#509, the composition root WU) also
needs the `PasswordHasher` (Argon2id, D88) and `SecretManager` (D25)
ports; those were already declared by the auth-reset WU but were
consolidated here in the squash-merge of W54 — re-declared for
clarity in this single source of truth.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass, field
from datetime import datetime
from typing import TYPE_CHECKING, Protocol
from uuid import UUID

if TYPE_CHECKING:
    from app.src.modules.lanzadera.domain.app import App
    from app.src.modules.lanzadera.domain.assignment import Assignment
    from app.src.modules.lanzadera.domain.global_admin import GlobalAdmin
    from app.src.modules.lanzadera.domain.profile import Profile
    from app.src.modules.lanzadera.domain.reset_token import ResetToken
    from app.src.modules.lanzadera.domain.user import User, UserStatus


class PasswordHasher(Protocol):
    """Argon2id port (DA-2, D88). Real adapter in PR 46.

    Methods are async to match the design (DA-1): the Postgres adapter
    uses ``argon2-cffi``'s async bindings under the hood, and the
    HTTP delivery that calls them already runs inside an event loop.
    The in-memory fakes in ``tests/lanzadera/auth/_fakes.py`` carry the
    matching ``async def`` shape so the contract is symmetric end to
    end.

    D89: the hash is the only persisted credential. ``update_password_and_activate``
    stores the new hash and flips ``status`` to ``ACTIVE`` in the same
    transaction.
    """

    async def hash(self, password: str) -> str: ...
    async def verify(self, password: str, password_hash: str) -> bool: ...


class SecretManager(Protocol):
    """Secret-retrieval port (CA-S2, D25, D88).

    The MVP reads from process environment (``EnvSecretManagerAdapter``);
    production swaps in Vault / AWS Secrets Manager without touching
    the domain (D73, D11). The secret value never appears in logs
    or CLI arguments (CA-S4): consumers receive opaque bytes via
    :meth:`get` and pass them to KDFs/ciphers that hold them
    in-memory only.
    """

    def get(self, key: str) -> str:
        """Return the secret material bound to ``key``.

        Raises:
            KeyError: when ``key`` is not configured in the underlying
                backend (e.g. env-var missing in dev). Callers MUST
                translate this to a user-facing 4xx/5xx, never log the
                key name alongside any other secret.
        """
        ...


class NotificationDelivery(Protocol):
    """Email delivery port (DA-10). Real adapter in W05 (#47).

    Async: the SMTP/SES adapter opens a network connection; a sync
    call would block the FastAPI event loop. The contract is a single
    ``send(...)`` matching the DA-10 message shape.
    """

    async def send(self, to: str, subject: str, body: str) -> None: ...


class Cache(Protocol):
    """TTL-bounded cache port (D70, D71).

    The launchable MVP uses ``cachetools.TTLCache`` (the
    ``AssumeInOfficeAdapter`` and the ``AssignmentRepository.effective_permissions``
    hot path both go through this seam). Production swaps in Redis or
    a LRU; the call site does not change.
    """

    def get(self, key: str) -> object | None: ...
    def set(self, key: str, value: object, ttl_seconds: int) -> None: ...


class Location(Protocol):
    """Office-presence port (H12, D52). MVP adapter is the always-on
    ``AssumeInOfficeAdapter`` (returns ``True`` for the corporate
    network). Production swaps in a real VPN / IP check without
    touching the domain.
    """

    def is_in_office(self, user_id: UUID) -> bool: ...


class BootstrapAdminSource(Protocol):
    """Bootstrap-admin source port (D21, D42, D48, D91, DA-6).

    Driven port that lists the e-mail addresses of the platform's first
    global admins. The composition root invokes this at process startup;
    the result feeds the bootstrap seed path (D21 + D42 + D48) that
    creates the initial ``users`` and ``global_admins`` rows. Idempotent:
    if the variable is unset, the source returns an empty sequence and
    the bootstrap is a no-op.

    The port is intentionally tiny: it returns addresses, not users.
    The seed-mutation lives in the application layer (composition root
    + the seed step of the Alembic 0001 migration) so the source stays
    substitutable (Vault, secret manager, LDAP) without touching domain.
    """

    def list_initial_emails(self) -> Sequence[str]:
        """Return every e-mail address that should bootstrap a global admin.

        An empty sequence means the operator has not configured the
        bootstrap yet — the caller MUST treat that as a no-op and not
        raise. The list is unordered; the caller applies its own
        ordering and dedup policy.
        """
        ...


class UserRepository(Protocol):
    """User storage port (DA-1, D89).

    Async to match the design (DA-1). The protocol was declared sync by
    PR #42 (#288) before the design ratified async for every repository;
    this PR restates the contract to match the AsyncSession-backed
    adapters (Postgres + the in-memory Fakes) and the async delivery
    adapters that already call ``await users.<method>(...)``.

    W54 (#508) extended the protocol to surface the admin-side
    operations (create, update_status, list_all, the failed_attempts
    family). The Postgres adapter and the in-memory fakes already
    implement all of these — this PR just brings the contract in
    line with the implementations.
    """

    async def get_by_email(self, email: str) -> User | None: ...
    async def get_by_id(self, user_id: UUID) -> User | None: ...
    async def list_all(self) -> Sequence[User]: ...
    async def create(self, user: User) -> None: ...
    async def update_status(self, user_id: UUID, status: UserStatus) -> None: ...
    async def update_password_and_activate(self, user_id: UUID, password_hash: str) -> None: ...
    async def update_failed_attempts(self, user_id: UUID, failed_attempts: int) -> None: ...
    async def record_login_attempt(self, user_id: UUID, *, at: datetime) -> None: ...
    async def reset_failed_attempts(self, user_id: UUID) -> None: ...


class ResetTokenRepository(Protocol):
    """Reset-token storage port (DA-4, D90). W02 (#45) ships the Postgres adapter.

    Async to match the design (DA-1). ``insert`` takes the
    domain ``ResetToken`` value object (the in-memory fake and the
    Postgres adapter both persist it as-is; the Postgres adapter
    decomposes it into columns under the hood).
    """

    async def insert(self, token: ResetToken) -> None: ...
    async def find_unused(self, token_hash: str, now: datetime) -> ResetToken | None: ...
    async def mark_consumed(self, token_hash: str, at: datetime) -> None: ...
    async def mark_superseded(self, user_id: UUID, at: datetime) -> None: ...
    async def purge_expired(self, now: datetime) -> int: ...


class GlobalAdminRepository(Protocol):
    """Global-admin storage port (D21, D42, DA-11). Async.

    W54 (#508) extended the surface with the operations the admin use
    cases need: is_global_admin (idempotency check), grant and revoke
    (the membership mutation), and list_all (audit / bootstrap).
    """

    async def there_is_any(self) -> bool: ...
    async def is_global_admin(self, user_id: UUID) -> bool: ...
    async def list_all(self) -> Sequence[GlobalAdmin]: ...
    async def grant(self, user_id: UUID) -> None: ...
    async def revoke(self, user_id: UUID) -> None: ...


class AssignmentRepositoryPort(Protocol):
    """Persistence boundary for user-app-profile assignments.

    D22, D42, DA-12, H11. The single source of truth for effective
    permissions in the platform: a (user, app, profile) triple plus
    soft-deletion via ``revoked_at``. The application layer enforces
    ``SinAcceso`` exclusivity on top of this port — see DA-12 + H11.

    The fourth method, ``effective_permissions``, is the read path the
    delivery layer hits on every request and that is why the adapter is
    required to cache it per ``(user_id, app_id)`` with TTL 60 s (DA-8);
    the cache is invalidated on every mutation.
    """

    async def create(self, user_id: UUID, app_id: int, profile_id: UUID) -> Assignment: ...
    async def list_for_user(self, user_id: UUID) -> Sequence[Assignment]: ...
    async def list_for_app(self, app_id: int) -> Sequence[Assignment]: ...
    async def effective_permissions(self, user_id: UUID, app_id: int) -> Sequence[str]: ...


class AppRepositoryPort(Protocol):
    """Persistence boundary for the app catalog.

    D7, D52, D85, DA-7. The single source of truth for the platform's
    app catalog: identity, deployment topology (central vs office-nas),
    office-presence requirement, and registration status. The
    ``list_active`` read path is the hot path for the delivery layer
    (``GET /admin/apps`` and the ``list_visible_to`` permission check);
    DA-8 caches it with a 5-minute TTL.
    """

    async def get_by_id(self, app_id: int) -> App | None:
        """Return the app with ``app_id`` or ``None`` if it does not exist."""
        ...

    async def list_active(self) -> Sequence[App]:
        """Return every app with ``registration_status='active'``.

        Ordered by ``id`` (the catalog is read-mostly; a stable
        ordering keeps the admin UI deterministic).
        """
        ...

    async def list_visible_to(self, user_id: UUID) -> Sequence[App]:
        """Every ``active`` app the user is assigned to (assignment EXISTS + NOT revoked)."""
        ...


class ProfileRepositoryPort(Protocol):
    """Persistence boundary for app-scoped profiles.

    D22, D45, DA-12. Profiles are scoped under an app: a profile
    belongs to exactly one app and is identified by ``(app_id, code)``.
    The Postgres adapter lands in W01 (#44); the in-memory fake lives in
    ``tests/lanzadera/auth/_fakes.py`` for the reset-flow and in
    ``tests/lanzadera/delivery/_fakes.py`` for the admin router.

    W54 (#508) added ``get_by_id`` for the ``assign_profile`` use case
    to look up the target profile by UUID.
    """

    async def list_for_app(self, app_id: int) -> Sequence[Profile]: ...
    async def get_by_code(self, app_id: int, code: str) -> Profile | None: ...
    async def get_by_id(self, profile_id: UUID) -> Profile | None: ...
    async def create(self, profile: Profile) -> None: ...
    async def set_active(self, app_id: int, code: str, *, active: bool) -> None: ...


@dataclass(frozen=True)
class AuditLogEntry:
    """Audit-row value object (DA-11). Stays narrow on the auth-reset surface.

    Mutable on purpose: ``payload`` may evolve after creation when
    additional context arrives in the same transaction (e.g. the lockout
    event records the final ``failed_attempts`` count). Other attributes
    are set-once.

    W54 (#508) widened the surface with ``module`` / ``correlation_id``
    / ``payload`` so the admin use cases can stamp canonical events
    with fresh UUIDs and structured context.
    """

    event_type: str
    actor_id: UUID | None
    target_id: str
    result: str
    created_at: datetime
    module: str = "lanzadera"
    correlation_id: UUID | None = None
    payload: dict[str, object] = field(default_factory=dict)


class AuditLog(Protocol):
    """Audit emission port (DA-11). Real adapter in PR 44. Async."""

    async def append(self, event: AuditLogEntry) -> None: ...
    async def list_for_actor(self, actor_id: UUID, since: datetime) -> Sequence: ...


__all__ = [
    "AppRepositoryPort",
    "AssignmentRepositoryPort",
    "AuditLog",
    "AuditLogEntry",
    "BootstrapAdminSource",
    "Cache",
    "GlobalAdminRepository",
    "Location",
    "NotificationDelivery",
    "PasswordHasher",
    "ProfileRepositoryPort",
    "ResetTokenRepository",
    "SecretManager",
    "UserRepository",
]
