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
from typing import TYPE_CHECKING

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
from app.src.modules.lanzadera.ports import (
    AuditLog,
    BootstrapAdminSource,
    PasswordHasher,
    SecretManager,
)

if TYPE_CHECKING:
    from datetime import datetime
    from uuid import UUID

    from app.src.modules.lanzadera.domain.app import App
    from app.src.modules.lanzadera.domain.user import User


def _adapter(adapter_cls, factory: AsyncSessionFactoryPort) -> object:
    return adapter_cls(factory)


class LanzaderaContainer:
    """Wires the driven ports to the Postgres adapters and the use cases.

    The container is constructed once at process start; the FastAPI
    app and the CLI driver both import it from a shared factory. Tests
    construct the container with the in-memory fakes in
    ``tests/lanzadera/_fakes.py`` and exercise the use cases
    end-to-end without a real database.
    """

    def __init__(
        self,
        *,
        session_factory: AsyncSessionFactoryPort,
        secret_manager: SecretManager,
        password_hasher: PasswordHasher,
        bootstrap_source: BootstrapAdminSource | None = None,
        clock: Callable[[], datetime] | None = None,
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

        # Build adapter instances once and re-use them per call. The
        # adapter is the Postgres-backed implementation; tests inject
        # fakes via this very constructor (D89 / DA-1: the connection
        # opens and closes per call inside the adapter).
        self._user_repo = UserRepositoryPg(self._factory)
        self._app_repo = AppRepositoryPg(self._factory)
        self._profile_repo = ProfileRepositoryPg(self._factory)
        self._assignment_repo = AssignmentRepositoryPg(self._factory)
        self._global_admin_repo = GlobalAdminRepositoryPg(self._factory)
        self._reset_token_repo = ResetTokenRepositoryPg(self._factory)
        self._audit = AuditLogPg(self._factory)

        # Build the use case partials once. The public method below
        # forwards the per-call ``actor_id`` etc. to the partial.
        self._use_cases = build_use_case_factories(
            user_repo=self._user_repo,
            app_repo=self._app_repo,
            profile_repo=self._profile_repo,
            assignment_repo=self._assignment_repo,
            global_admins=self._global_admin_repo,
            reset_tokens=self._reset_token_repo,
            audit=self._audit,
            password_hasher=self._password_hasher,
            secret_manager=self._secret_manager,
            clock=self._clock,
        )

    @property
    def bootstrap_source(self) -> BootstrapAdminSource:
        return self._bootstrap_source

    # -- use cases ------------------------------------------------------------

    async def create_user(
        self,
        email: str,
        *,
        name: str,
        national_id: str,
        actor_id: UUID | None = None,
    ) -> User:
        return await self._use_cases["create_user"](
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
        await self._use_cases["disable_user"](user_id=user_id, actor_id=actor_id)

    async def grant_global_admin(
        self,
        user_id: UUID,
        *,
        actor_id: UUID | None = None,
    ) -> None:
        await self._use_cases["grant_global_admin"](user_id=user_id, actor_id=actor_id)

    async def revoke_global_admin(
        self,
        user_id: UUID,
        *,
        actor_id: UUID | None = None,
    ) -> None:
        await self._use_cases["revoke_global_admin"](user_id=user_id, actor_id=actor_id)

    async def assign_profile(
        self,
        user_id: UUID,
        app_id: int,
        profile_id: UUID,
        *,
        actor_id: UUID | None = None,
    ) -> None:
        await self._use_cases["assign_profile"](
            user_id=user_id,
            app_id=app_id,
            profile_id=profile_id,
            actor_id=actor_id,
        )

    async def list_effective_apps(self, user_id: UUID) -> Sequence[App]:
        return await self._use_cases["list_effective_apps"](user_id=user_id)

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
        return await self._use_cases["audit_append"](
            event_type=event_type,
            actor_id=actor_id,
            target_id=target_id,
            result=result,
            payload=payload,
            correlation_id=correlation_id,
            module=module,
        )

    async def bootstrap_global_admins(
        self,
        *,
        actor_id: UUID | None = None,
    ) -> int:
        return await self._use_cases["bootstrap_global_admins"](actor_id=actor_id)

    async def set_password(
        self,
        email: str,
        *,
        new_password: str,
        actor_id: UUID | None = None,
    ) -> None:
        await self._use_cases["set_password"](
            email=email,
            new_password=new_password,
            actor_id=actor_id,
        )

        @property
        def users(self) -> UserRepositoryPg:
            """Read-only access to the UserRepository for admin queries."""
            return self._user_repo

        @property
        def app_repo(self) -> AppRepositoryPg:
            """Read-only access to the AppRepository for admin queries."""
            return self._app_repo

        @property
        def assignment_repo(self) -> AssignmentRepositoryPg:
            """Read-only access to the AssignmentRepository for admin queries."""
            return self._assignment_repo

        @property
        def audit_repo(self) -> AuditLogPg:
            """Read-only access to the AuditLog for admin queries."""
            return self._audit

        async def list_all_users(self) -> Sequence[User]:
            """List all platform users (admin-only)."""
            return await self._user_repo.list_all()

    __all__ = ["LanzaderaContainer"]
