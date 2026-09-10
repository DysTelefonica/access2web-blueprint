# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp W55
# W65 (#596): port wiring split to _container_ports + _container_auth;
# use case methods split to _container_facade_auth/users/apps.
"""Composition root for the lanzadera module (CA-F4, D73, DA-1).

W65 (#596): ``container.py`` es un thin wrapper (~80 sitios) que
instancia los tres facade classes, almacena ports directas, y delega
via ``__getattr__``:
  ``LanzaderaFacadeAuth``   — login, logout, set_password + auth ports
  ``LanzaderaFacadeUsers``   — create_user, disable_user, grant/revoke_admin + user ports
  ``LanzaderaFacadeApps``   — assign_profile, list_effective_apps, bootstrap + app ports

Los properties públicos (``container.users``, ``container.audit``, etc.) se
exponen directamente en ``LanzaderaContainer`` porque los routes los usan
como kwargs de los use cases (no via facade methods).
"""

from __future__ import annotations

# di-only marker
import time as _di_t  # noqa: F401
from typing import TYPE_CHECKING, Any

from app.src.modules.lanzadera.adapters.bootstrap.env_admin_source_adapter import (
    EnvAdminSourceAdapter,
)
from app.src.modules.lanzadera.adapters.persistence.async_session_factory import (
    AsyncSessionFactoryPort,
)
from app.src.modules.lanzadera.di._container_auth import build_auth_ports
from app.src.modules.lanzadera.di._container_facade_apps import LanzaderaFacadeApps
from app.src.modules.lanzadera.di._container_facade_auth import LanzaderaFacadeAuth
from app.src.modules.lanzadera.di._container_facade_users import LanzaderaFacadeUsers
from app.src.modules.lanzadera.di.use_cases import build_use_case_factories
from app.src.modules.lanzadera.domain.ports import AuditLog, PasswordHasher
from app.src.modules.lanzadera.domain.ports.bootstrap_admin_source import (
    BootstrapAdminSource,
)
from app.src.modules.lanzadera.domain.ports.secret_manager import SecretManager

if TYPE_CHECKING:
    from datetime import datetime


def _default_clock() -> datetime:
    from datetime import UTC, datetime

    return datetime.now(UTC)


class LanzaderaContainer:
    """Thin composition root — wires ports and delegates to facade classes.

    W65 (#596): port wiring va a ``_container_ports`` / ``_container_auth``;
    los métodos de use cases van a ``LanzaderaFacadeAuth``,
    ``LanzaderaFacadeUsers``, ``LanzaderaFacadeApps``. Esta clase
    instancia las tres y delega ``__getattr__``. Las ports usadas
    directamente por los routes se almacenan también como properties
    públicos para evitar el overhead de ``__getattr__``.
    """

    def __init__(
        self,
        *,
        session_factory: AsyncSessionFactoryPort | None = None,
        secret_manager: SecretManager,
        password_hasher: PasswordHasher,
        bootstrap_source: BootstrapAdminSource | None = None,
        clock: Any = None,
        user_repo: Any = None,
        app_repo: Any = None,
        profile_repo: Any = None,
        assignment_repo: Any = None,
        global_admin_repo: Any = None,
        reset_token_repo: Any = None,
        presence_repo: Any = None,
        audit: AuditLog | None = None,
        session_repo: Any = None,
        jwt_signer: Any = None,
    ) -> None:
        self._factory = session_factory
        self._bootstrap_source = (
            bootstrap_source if bootstrap_source is not None else EnvAdminSourceAdapter()
        )
        self._clock = clock if clock is not None else _default_clock

        auth_ports = build_auth_ports(
            session_factory=session_factory,
            password_hasher=password_hasher,
            session_repo=session_repo,
            jwt_signer=jwt_signer,
        )
        # Extract auth ports for direct container properties.
        _session_repo = auth_ports.pop("session_repo")
        _jwt_signer = auth_ports.pop("jwt_signer")

        self._use_cases = build_use_case_factories(
            user_repo=user_repo,
            app_repo=app_repo,
            profile_repo=profile_repo,
            assignment_repo=assignment_repo,
            global_admin_repo=global_admin_repo,
            reset_token_repo=reset_token_repo,
            presence_repo=presence_repo,
            session_repo=_session_repo,
            audit=audit,  # type: ignore[arg-type]
            password_hasher=password_hasher,
            secret_manager=secret_manager,
            clock=self._clock,
        )

        # Direct port properties (used by routes as kwargs to use cases).
        self.users = user_repo
        self.sessions = _session_repo
        self.password_hasher = password_hasher
        self.audit = audit
        self.jwt_signer = _jwt_signer
        self.app_repo = app_repo
        self.audit_repo = audit
        self.assignment_repo = assignment_repo
        self.presence_repo = presence_repo

        # Three facade instances — each holds the use-case dict + clock.
        self._facade_auth = LanzaderaFacadeAuth(
            self._use_cases,
            self._clock,
            _session_repo,
            _jwt_signer,
            password_hasher,
        )
        self._facade_users = LanzaderaFacadeUsers(
            self._use_cases, self._clock, user_repo, global_admin_repo, audit
        )
        self._facade_apps = LanzaderaFacadeApps(
            self._use_cases,
            self._clock,
            app_repo,
            assignment_repo,
            audit,
            presence_repo,
        )

    # -- public properties that the facade classes don't own -------------------

    @property
    def bootstrap_source(self) -> BootstrapAdminSource:
        return self._bootstrap_source

    @property
    def use_cases(self) -> dict[str, Any]:
        """Expose the use-case partials (SSE routes use ``container.use_cases``)."""
        return self._use_cases

    # -- __getattr__ delegation: forward unknown names to the facades ---------

    def __getattr__(self, name: str) -> Any:
        """Delegate unknown attributes to the appropriate facade."""
        if name.startswith("_"):
            raise AttributeError(name)
        if hasattr(LanzaderaFacadeAuth, name):
            return getattr(self._facade_auth, name)
        if hasattr(LanzaderaFacadeUsers, name):
            return getattr(self._facade_users, name)
        if hasattr(LanzaderaFacadeApps, name):
            return getattr(self._facade_apps, name)
        raise AttributeError(f"'{type(self).__name__}' object has no attribute '{name}'")


__all__ = ["LanzaderaContainer"]
