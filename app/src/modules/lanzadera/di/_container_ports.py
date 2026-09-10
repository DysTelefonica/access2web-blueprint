# HARNESS-PROVENANCE: deterministic-quality-harness v1.8 + lanzadera-mvp W65
"""Port wiring for non-auth ports (W65 #596).

Extracted from ``container.py`` to keep every module under the 100-site
mutation ceiling. The thin ``container.py`` composes this with
``_container_auth`` and the use-case partials.

W65 (#596): cleave LanzaderaContainer by port group.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING, Any, cast

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
from app.src.modules.lanzadera.domain.ports import AuditLog, UserRepository
from app.src.modules.lanzadera.domain.ports.app_repository import AppRepositoryPort
from app.src.modules.lanzadera.domain.ports.assignment_repository import (
    AssignmentRepositoryPort,
)
from app.src.modules.lanzadera.domain.ports.global_admin_repository import (
    GlobalAdminRepositoryPort,
)
from app.src.modules.lanzadera.domain.ports.presence_repository import PresenceRepository
from app.src.modules.lanzadera.domain.ports.profile_repository import ProfileRepositoryPort

if TYPE_CHECKING:
    from app.src.modules.lanzadera.di._container_auth import AuthPorts


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
    production path falls through to the adapter constructor.
    """
    if fake is not None:
        return fake
    assert factory is not None, "no fake and no factory — caller must inject one"
    return _adapter(adapter_cls, factory)


def build_nonauth_ports(
    *,
    session_factory: AsyncSessionFactoryPort | None,
    user_repo: UserRepository | None,
    app_repo: AppRepositoryPort | None,
    profile_repo: ProfileRepositoryPort | None,
    assignment_repo: AssignmentRepositoryPort | None,
    global_admin_repo: GlobalAdminRepositoryPort | None,
    reset_token_repo: object | None,
    presence_repo: PresenceRepository | None,
    audit: AuditLog | None,
) -> dict[str, object]:
    """Build the 7 non-auth Postgres adapters.

    W65 (#596): extracted from ``LanzaderaContainer.__init__`` to keep
    both this module and the thin container under the 100-site ceiling.
    """
    _user_repo = _pick(user_repo, session_factory, UserRepositoryPg)
    _app_repo = _pick(app_repo, session_factory, AppRepositoryPg)
    _profile_repo = _pick(profile_repo, session_factory, ProfileRepositoryPg)
    _assignment_repo = _pick(assignment_repo, session_factory, AssignmentRepositoryPg)
    _global_admin_repo = _pick(global_admin_repo, session_factory, GlobalAdminRepositoryPg)
    _reset_token_repo = _pick(reset_token_repo, session_factory, ResetTokenRepositoryPg)
    _presence_repo = _pick(presence_repo, session_factory, PresenceRepositoryPg)
    _audit = cast(AuditLog, _pick(audit, session_factory, AuditLogPg))

    return {
        "user_repo": _user_repo,
        "app_repo": _app_repo,
        "profile_repo": _profile_repo,
        "assignment_repo": _assignment_repo,
        "global_admin_repo": _global_admin_repo,
        "reset_token_repo": _reset_token_repo,
        "presence_repo": _presence_repo,
        "audit": _audit,
    }
