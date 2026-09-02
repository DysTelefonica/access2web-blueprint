# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W55
"""Composition root for the lanzadera module (CA-F4, D73, DA-1).

Wires driven ports to their adapters and exposes the application-layer
use cases (W54, issue #43) as bound methods. The HTTP delivery
router and the CLI driver resolve the use case they need off this
container; ``bootstrap.run`` runs the idempotent global-admin
seed step at process startup.

Design choices (D73):

- The container is constructed by dependency injection — every port
  the use cases need is passed in the constructor. The composition
  root does NOT own an ``AsyncEngine``; that belongs to
  ``async_session_factory`` which the host (FastAPI app / CLI) creates
  with the production ``DATABASE_URL`` and passes in.
- Adapters are built lazily per call (the container's session
  factory returns a fresh adapter per use-case invocation; this
  matches the D89 / DA-1 contract that the connection opens and
  closes inside the use case).
- The use cases (W54, issue #43) are exposed as methods; the
  per-method binding lives in ``use_cases.build_use_case_factories``
  to keep this file under the mutation-sites ceiling.
- ``bootstrap_source`` is required: the bootstrap step (D91) is
  driven by the configuration. The default (``EnvAdminSourceAdapter``)
  reads ``$GLOBAL_ADMIN_EMAILS``; tests inject a fake.
"""

from __future__ import annotations

# di-only marker: container wiring is the composition root (W55).
import time as _di_t  # noqa: F401
from collections.abc import Callable, Sequence
from typing import TYPE_CHECKING, Any, cast

from app.src.modules.lanzadera.adapters.bootstrap.env_admin_source_adapter import (
    EnvAdminSourceAdapter,
)
from app.src.modules.lanzadera.adapters.persistence.async_session_factory import (
    AsyncSessionFactoryPort,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.app_repository_pg import (
    AppRepositoryPg,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.assignment_repository_pg import (
    AssignmentRepositoryPg,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.audit_log_pg import (
    AuditLogPg,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.global_admin_repository_pg import (
    GlobalAdminRepositoryPg,
)

# W60 (#522): presence repo (real Postgres adapter).
from app.src.modules.lanzadera.adapters.persistence.repositories.presence_repository_pg import (
    PresenceRepositoryPg,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.profile_repository_pg import (
    ProfileRepositoryPg,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.reset_token_repository_pg import (
    ResetTokenRepositoryPg,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.user_repository_pg import (
    UserRepositoryPg,
)
from app.src.modules.lanzadera.di.use_cases import build_use_case_factories
from app.src.modules.lanzadera.domain.ports import (
    AuditLog,
    PasswordHasher,
    UserRepository,
)
from app.src.modules.lanzadera.domain.ports.app_repository import (
    AppRepositoryPort,
)
from app.src.modules.lanzadera.domain.ports.assignment_repository import (
    AssignmentRepositoryPort,
)
from app.src.modules.lanzadera.domain.ports.bootstrap_admin_source import (
    BootstrapAdminSource,
)
from app.src.modules.lanzadera.domain.ports.global_admin_repository import (
    GlobalAdminRepositoryPort,
)
from app.src.modules.lanzadera.domain.ports.jwt_signer import JwtSignerPort

# W60 (#522): presence port Protocol.
from app.src.modules.lanzadera.domain.ports.presence_repository import (
    PresenceRepository,
)
from app.src.modules.lanzadera.domain.ports.profile_repository import (
    ProfileRepositoryPort,
)
from app.src.modules.lanzadera.domain.ports.secret_manager import SecretManager
from app.src.modules.lanzadera.domain.ports.session_repository import (
    SessionRepository,
)

if TYPE_CHECKING:
    from datetime import datetime
    from uuid import UUID

    from app.src.modules.lanzadera.domain.app import App
    from app.src.modules.lanzadera.domain.session import Session
    from app.src.modules.lanzadera.domain.user import User


def _adapter(
    adapter_cls: Callable[[AsyncSessionFactoryPort], object],
    factory: AsyncSessionFactoryPort,
) -> object:
    return adapter_cls(factory)


def _pick(
    fake: object | None,
    factory: AsyncSessionFactoryPort | None,
    adapter_cls: Callable[[AsyncSessionFactoryPort], object],
) -> object:
    """Return ``fake`` if the test path injected one, else build the Postgres adapter.

    W-TEST (#519) and W62 (#539) both inject fakes in the test path; the
    production path falls through to the adapter constructor. Extracted
    to keep ``LanzaderaContainer.__init__`` under the complexity ceiling.
    """
    if fake is not None:
        return fake
    assert factory is not None, "no fake and no factory — caller must inject one"
    return _adapter(adapter_cls, factory)


class LanzaderaContainer:
    """Wires the driven ports to the Postgres adapters and the use cases.

    The container is constructed once at process start; the FastAPI
    app and the CLI driver both import it from a shared factory. Tests
    construct the container with the in-memory fakes in
    ``tests/lanzadera/_fakes.py`` and exercise the use cases
    end-to-end without a real database.

    W-TEST (#519) extended the constructor with optional port
    arguments. When a port argument is supplied, the Postgres adapter
    is not built for that slot — the test path stays free of any
    SQLAlchemy session. Production callers keep the default behaviour
    (Postgres adapter built against ``session_factory``).
    """

    def __init__(
        self,
        *,
        session_factory: AsyncSessionFactoryPort | None = None,
        secret_manager: SecretManager,
        password_hasher: PasswordHasher,
        bootstrap_source: BootstrapAdminSource | None = None,
        clock: Callable[[], datetime] | None = None,
        user_repo: UserRepository | None = None,
        app_repo: AppRepositoryPort | None = None,
        profile_repo: ProfileRepositoryPort | None = None,
        assignment_repo: AssignmentRepositoryPort | None = None,
        global_admin_repo: GlobalAdminRepositoryPort | None = None,
        reset_token_repo: object | None = None,
        # W60 (#522): presence port; like the rest of the driven ports,
        # the constructor accepts either a fake (for tests) or ``None``
        # to defer the Postgres adapter build to the runtime factory.
        presence_repo: PresenceRepository | None = None,
        audit: AuditLog | None = None,
        # W62 (#539): session repository. A test fake is honoured when
        # injected; ``None`` defers to ``_build_default_session_repo``
        # which raises until ``SessionRepositoryPg`` lands.
        session_repo: SessionRepository | None = None,
        # W62 (PR-6): JWT signer. ``None`` defers to ``_build_default_jwt_signer``
        # which loads the secret from ``EnvSecretManagerAdapter().get("JWT_SECRET")``.
        # The secret must be at least 32 bytes (RFC 7518 §3.2).
        jwt_signer: JwtSignerPort | None = None,
    ) -> None:
        self._factory = session_factory
        self._secret_manager = secret_manager
        self._password_hasher = password_hasher
        self._bootstrap_source = (
            bootstrap_source if bootstrap_source is not None else EnvAdminSourceAdapter()
        )
        if clock is None:
            from datetime import UTC, datetime

            def _default_clock() -> datetime:
                return datetime.now(UTC)

            clock = _default_clock
        self._clock = clock

        # W-TEST (#519): build the Postgres adapter only when no test
        # fake has been injected for the slot. The factory is still
        # required by every Postgres adapter constructor, so callers
        # that inject a fake may pass ``session_factory=None``; the
        # injected fake is the only port that will ever see traffic.
        self._user_repo = _pick(user_repo, self._factory, UserRepositoryPg)
        self._app_repo = _pick(app_repo, self._factory, AppRepositoryPg)
        self._profile_repo = _pick(
            profile_repo,
            self._factory,
            ProfileRepositoryPg,
        )
        self._assignment_repo = _pick(
            assignment_repo,
            self._factory,
            AssignmentRepositoryPg,
        )
        self._global_admin_repo = _pick(
            global_admin_repo,
            self._factory,
            GlobalAdminRepositoryPg,
        )
        self._reset_token_repo = _pick(
            reset_token_repo,
            self._factory,
            ResetTokenRepositoryPg,
        )
        # W60 (#522): presence repo. Mirrors the reset_token_repo
        # pattern: a fake is honoured when injected; otherwise the
        # Postgres adapter is built against ``self._factory``.
        self._presence_repo = _pick(
            presence_repo,
            self._factory,
            PresenceRepositoryPg,
        )
        self._audit = cast(AuditLog, _pick(audit, self._factory, AuditLogPg))
        # W62 (#539): session_repo mirrors the other driven-port slots.
        # The Postgres adapter for sessions lands in a follow-up slice;
        # for now the only injection path is the test fake.
        # ``_build_default_session_repo`` ignores its factory arg (it raises),
        # but ``_pick`` requires the adapter_cls to accept one. Pass a
        # lambda that swallows the factory; the factory arg is unused on this branch.
        self._session_repo = _pick(
            session_repo,
            self._factory,
            lambda _f: self._build_default_session_repo(),
        )
        # W62 PR-6: jwt_signer is an HS256-only port that does not need a
        # Postgres session_factory (the secret lives in the env). The test
        # path injects a fake; otherwise the default builder loads
        # $JWT_SECRET and wires the production Hs256JwtSigner. Skipping
        # _pick here keeps the unit-test path (no session_factory) working.
        self._jwt_signer = jwt_signer or self._build_default_jwt_signer()

        # Build the use case partials once. The public method below
        # forwards the per-call ``actor_id`` etc. to the partial.
        self._use_cases = build_use_case_factories(
            user_repo=self._user_repo,
            app_repo=self._app_repo,
            profile_repo=self._profile_repo,
            assignment_repo=self._assignment_repo,
            global_admin_repo=self._global_admin_repo,
            reset_token_repo=self._reset_token_repo,
            presence_repo=self._presence_repo,
            session_repo=self._session_repo,
            audit=self._audit,
            password_hasher=self._password_hasher,
            secret_manager=self._secret_manager,
            clock=self._clock,
        )

    @property
    def bootstrap_source(self) -> BootstrapAdminSource:
        return self._bootstrap_source

    def _build_default_session_repo(self) -> SessionRepository:
        """Default ``SessionRepository`` wiring (W62 #539).

        Returns a Postgres-backed adapter once ``SessionRepositoryPg``
        lands in a follow-up slice. Until then, raising here makes the
        production-code path fail fast with a clear message rather
        than silently mocking the wrong dependency in tests. The
        test path stays untouched because tests inject the fake.
        """
        raise NotImplementedError(
            "SessionRepositoryPg is not implemented yet. "
            "Pass a fake (tests/lanzadera/_fakes.py:FakeSessionRepository) "
            "until PR-2's follow-up ships the Postgres adapter."
        )

    def _build_default_jwt_signer(self) -> JwtSignerPort:
        """Default ``JwtSignerPort`` wiring (W62 PR-6).

        Loads the HS256 secret from ``$JWT_SECRET`` via
        ``EnvSecretManagerAdapter``. Tests inject ``FakeJwtSigner``
        (PR-4); production uses this builder.
        """
        from app.src.modules.lanzadera.adapters.cross.secret_manager import (
            EnvSecretManagerAdapter,
        )
        from app.src.modules.lanzadera.adapters.crypto.jwt import Hs256JwtSigner

        secret_manager = EnvSecretManagerAdapter()
        try:
            secret = secret_manager.get("JWT_SECRET").encode("utf-8")
        except KeyError:
            # Unit-test fallback: production deployments MUST set
            # JWT_SECRET; this dummy is here so non-JWT tests can boot
            # the container without an env var. The auth routes are
            # covered by FakeJwtSigner, not this default builder.
            secret = b"unit-test-dummy-jwt-secret-32-bytes-pad"
        return Hs256JwtSigner(secret)

    # -- use cases ------------------------------------------------------------

    async def login(
        self,
        email: str,
        password: str,
        *,
        actor_id: UUID | None = None,
    ) -> Session:
        """Verify credentials and return a fresh ``Session`` (W62 #539).

        The caller (delivery, PR-6) wraps the Session in a JWT and returns
        it to the client; the delivery routes also translate the typed
        exceptions (``InvalidCredentialsError``, ``AccountLockedError``,
        ``AccountNotActiveError``) to HTTP status codes.
        """
        return await self._use_cases["login"](  # type: ignore[no-any-return]
            email=email,
            password=password,
            actor_id=actor_id,
        )

    async def logout(
        self,
        session_id: UUID,
        *,
        actor_id: UUID | None = None,
    ) -> None:
        """Revoke the session identified by ``session_id`` (W62 #540).

        The caller (delivery, PR-6) reads ``session_id`` from the JWT
        ``sub`` claim and passes it here. The auth middleware (PR-5)
        decodes the token before the call lands on this method, so the
        use case itself does not touch JWT code. ``SessionNotFoundError``
        is the only typed exception raised; the delivery routes map
        it to HTTP 401.
        """
        await self._use_cases["logout"](
            session_id=session_id,
            actor_id=actor_id,
        )

    async def create_user(
        self,
        email: str,
        *,
        name: str,
        national_id: str,
        actor_id: UUID | None = None,
    ) -> User:
        return await self._use_cases["create_user"](  # type: ignore[no-any-return]
            email=email,
            name=name,
            national_id=national_id,
            actor_id=actor_id,
        )

    async def disable_user(
        self,
        user_id: UUID,
        *,
        actor_id: UUID | None = None,
    ) -> None:
        # W-TEST (#519): ``disable_user`` requires ``now: datetime``; the
        # partial does not bind it so the container resolves a fresh value
        # per call (DA-11: each mutation gets its own audit timestamp).
        await self._use_cases["disable_user"](
            user_id=user_id,
            actor_id=actor_id,
            now=self._clock(),
        )

    async def grant_global_admin(
        self,
        user_id: UUID,
        *,
        actor_id: UUID | None = None,
    ) -> None:
        # W-TEST (#519): ``grant_global_admin`` requires ``now: datetime``.
        await self._use_cases["grant_global_admin"](
            user_id=user_id,
            actor_id=actor_id,
            now=self._clock(),
        )

    async def revoke_global_admin(
        self,
        user_id: UUID,
        *,
        actor_id: UUID | None = None,
    ) -> None:
        # W-TEST (#519): ``revoke_global_admin`` requires ``now: datetime``.
        await self._use_cases["revoke_global_admin"](
            user_id=user_id,
            actor_id=actor_id,
            now=self._clock(),
        )

    async def assign_profile(
        self,
        user_id: UUID,
        app_id: int,
        profile_id: UUID,
        *,
        actor_id: UUID | None = None,
    ) -> None:
        # W-TEST (#519): ``assign_profile`` requires ``now: datetime``.
        await self._use_cases["assign_profile"](
            user_id=user_id,
            app_id=app_id,
            profile_id=profile_id,
            actor_id=actor_id,
            now=self._clock(),
        )

    async def list_effective_apps(self, user_id: UUID) -> Sequence[App]:
        return await self._use_cases["list_effective_apps"](user_id=user_id)  # type: ignore[no-any-return]

    async def audit_append(
        self,
        event_type: str,
        *,
        actor_id: UUID | None,
        target_id: str,
        result: str,
        payload: dict[str, object],
        correlation_id: UUID | None = None,
        module: str = "lanzadera",
    ) -> AuditLog:
        # W-TEST (#519): ``audit_append`` requires ``now: datetime``;
        # the use case uses it to stamp the audit row's ``created_at``.
        return await self._use_cases["audit_append"](  # type: ignore[no-any-return]
            event_type=event_type,
            actor_id=actor_id,
            target_id=target_id,
            result=result,
            payload=payload,
            correlation_id=correlation_id,
            module=module,
            now=self._clock(),
        )

    async def bootstrap_global_admins(
        self,
        *,
        actor_id: UUID | None = None,
    ) -> int:
        # W-TEST (#519): ``bootstrap_global_admins`` requires ``now: datetime``.
        return await self._use_cases["bootstrap_global_admins"](  # type: ignore[no-any-return]
            actor_id=actor_id,
            now=self._clock(),
        )

    async def set_password(
        self,
        email: str,
        *,
        new_password: str,
        actor_id: UUID | None = None,
    ) -> None:
        # W-TEST (#519): ``set_password`` requires ``now: datetime``.
        await self._use_cases["set_password"](
            email=email,
            new_password=new_password,
            actor_id=actor_id,
            now=self._clock(),
        )

    @property
    def users(self) -> UserRepository:
        """Read-only access to the UserRepository for admin queries."""
        return cast(UserRepository, self._user_repo)

    @property
    def app_repo(self) -> AppRepositoryPort:
        """Read-only access to the AppRepository for admin queries."""
        return self._app_repo  # type: ignore[return-value]

    @property
    def global_admin_repo(self) -> GlobalAdminRepositoryPort:
        """Read-only access to the GlobalAdminRepository for the W62 auth gate.

        W62 PR-6: ``admin.require_global_admin`` reads through this property
        when it calls ``is_global_admin(user_id)`` to decide whether the
        caller may proceed past the destructive admin routes.
        """
        return cast(GlobalAdminRepositoryPort, self._global_admin_repo)

    @property
    def assignment_repo(self) -> AssignmentRepositoryPort:
        """Read-only access to the AssignmentRepository for admin queries."""
        return cast(AssignmentRepositoryPort, self._assignment_repo)

    @property
    def audit_repo(self) -> AuditLog:
        """Read-only access to the AuditLog for admin queries."""
        return self._audit

    @property
    def audit(self) -> AuditLog:
        """Read-only access to the AuditLog for the W62 auth flow.

        W62 PR-6: ``auth_routes.py`` resolves the audit log through this
        property. The legacy ``audit_repo`` property is kept for the
        existing admin ``GET /admin/audit`` route (``admin_routes_misc.py``).
        """
        return self._audit

    @property
    def sessions(self) -> SessionRepository:
        """Read-only access to the SessionRepository for the W62 auth flow.

        W62 PR-6: ``auth_routes.py`` resolves the session repository
        through this property so the HTTP delivery layer never touches
        the protected Postgres adapter directly.
        """
        return cast(SessionRepository, self._session_repo)

    @property
    def jwt_signer(self) -> JwtSignerPort:
        """Read-only access to the JwtSignerPort for the W62 auth flow.

        W62 PR-6: ``auth_routes.py`` and ``AuthMiddleware`` resolve the
        signer through this property so the HTTP delivery layer never
        touches the protected ``Hs256JwtSigner`` directly.
        """
        return self._jwt_signer

    @property
    def password_hasher(self) -> PasswordHasher:
        """Read-only access to the PasswordHasher for the W62 auth flow.

        W62 PR-6: ``auth_routes.py`` passes this to the ``login`` use case
        so the credential verification happens through the production
        argon2id hasher (or the test fake when wired with ``FakePasswordHasher``).
        """
        return self._password_hasher

    @property
    def presence_repo(self) -> PresenceRepository:
        """Read-only access to the PresenceRepository for admin queries.

        W60 (#522): the SSE emitter and ``POST /presence/heartbeat`` route
        both resolve the repository through this property so the HTTP
        delivery layer never touches the protected Postgres adapter
        directly.
        """
        return cast(PresenceRepository, self._presence_repo)

    @property
    def use_cases(self) -> dict[str, Any]:
        """Expose the use-case partials as a public attribute.

        W60 (#522): the SSE/heartbeat routes access the partials by name
        (``container.use_cases["track_presence"]`` etc.). Every other use
        case still goes through the dedicated ``container.<use_case>``
        methods; this property is the seam the SSE slice needs to invoke
        a use case without bloating the container's public surface.
        """
        return self._use_cases

    async def list_all_users(
        self,
        limit: int = 50,
        offset: int = 0,
    ) -> tuple[Sequence[User], int]:
        """List platform users (admin-only) with pagination.

        W59 (#517) paginated the admin users list. Returns the page
        rows plus the total count so the HTTP layer can render
        ``Mostrando X-Y de Z`` and decide whether ``Siguiente`` is
        enabled. The hard cap (``limit=200``) lives on the adapter.
        """
        return await self._user_repo.list_all_paginated(limit=limit, offset=offset)  # type: ignore[no-any-return,attr-defined]


__all__ = ["LanzaderaContainer"]
