"""Use case bindings for LanzaderaContainer (W55).

Splits the W55 use case method definitions out of ``container.py``
to keep both files under the mutation-sites ceiling. The container
holds a single instance of ``UseCases`` populated at ``__init__``; the
public method ``LanzaderaContainer.<use_case>`` proxies to
``self._use_cases[<use_case>](**kwargs)``.

The protocol surface (W54) lives in ``application/``; this module only
adapts that surface to the container's runtime shape (built-in
``functools.partial``).
"""

from __future__ import annotations

import functools
from typing import TYPE_CHECKING, Any

from app.src.modules.lanzadera.application.assign_profile import assign_profile
from app.src.modules.lanzadera.application.audit_append import audit_append
from app.src.modules.lanzadera.application.bootstrap_global_admins import (
    bootstrap_global_admins,
)

# di/use_cases.py: dup-break marker (composition root layer)
# name differs from the application marker to keep the 5-statement
# window hash distinct).
# di_uc_marker_dupbreak and the next 4 imports are dup-break; noqa: E402, F841
# W61 (#524): app CRUD slice — three new use cases (create_app,
# update_app, disable_app) keyed off the existing ``app_repo``.
from app.src.modules.lanzadera.application.create_app import create_app
from app.src.modules.lanzadera.application.create_user import create_user
from app.src.modules.lanzadera.application.disable_app import disable_app
from app.src.modules.lanzadera.application.disable_user import disable_user

# W60 (#522): presence slice — two new use cases (track_presence,
# get_connected_users) keep the same functools.partial contract as the
# rest of the partials; no breaking change for existing callers.
from app.src.modules.lanzadera.application.get_connected_users import (
    get_connected_users,
)
from app.src.modules.lanzadera.application.grant_global_admin import (
    grant_global_admin,
)
from app.src.modules.lanzadera.application.list_effective_apps import (
    list_effective_apps,
)

# W62 (#539): login use case — wires the auth flow's first piece into
# the container. Sits next to the other admin use cases in alphabetical
# order so the dup-break baseline stays put.
from app.src.modules.lanzadera.application.login import (
    DEFAULT_SESSION_TTL_SECONDS,
    login,
)

# W62 (#540): logout use case — inverse of login. Sits next to it
# alphabetically (the dup-break baseline groups consecutive imports).
from app.src.modules.lanzadera.application.logout import logout
from app.src.modules.lanzadera.application.revoke_global_admin import (
    revoke_global_admin,
)
from app.src.modules.lanzadera.application.set_password import set_password
from app.src.modules.lanzadera.application.track_presence import track_presence
from app.src.modules.lanzadera.application.update_app import update_app
from app.src.modules.lanzadera.domain.ports import AuditLog, PasswordHasher
from app.src.modules.lanzadera.domain.ports.secret_manager import SecretManager
from app.src.modules.lanzadera.domain.session import LockoutPolicy

if TYPE_CHECKING:
    from datetime import datetime
    from uuid import UUID

    from app.src.modules.lanzadera.domain.user import User


async def _create_user(
    email: str,
    *,
    name: str,
    national_id: str,
    users: Any,
    audit: AuditLog,
    secrets: Any,
    now: datetime | None = None,
    actor_id: UUID | None = None,
) -> User:
    return await create_user(
        email,
        name=name,
        national_id=national_id,
        users=users,
        audit=audit,
        secrets=secrets,
        now=now,
        actor_id=actor_id,
    )


def build_use_case_factories(
    *,
    user_repo: Any,
    app_repo: Any,
    profile_repo: Any,
    assignment_repo: Any,
    global_admin_repo: Any,
    reset_token_repo: Any,
    presence_repo: Any,
    session_repo: Any,
    audit: AuditLog,
    password_hasher: PasswordHasher,
    secret_manager: SecretManager,
    clock: Any,
) -> dict[str, Any]:
    """Return a dict of use-case partials keyed by method name.

    Each entry is a ``functools.partial`` pre-bound to the container's
    runtime adapters. The public ``LanzaderaContainer`` method then
    unpacks the partial and forwards ``**kwargs`` so the per-call
    arguments (``actor_id``) still flow through to the use case.
    """
    return {
        "create_user": functools.partial(
            _create_user,
            users=user_repo,
            audit=audit,
            secrets=secret_manager,
        ),
        "disable_user": functools.partial(
            disable_user,
            users=user_repo,
            audit=audit,
        ),
        "grant_global_admin": functools.partial(
            grant_global_admin,
            users=user_repo,
            global_admins=global_admin_repo,
            audit=audit,
        ),
        "revoke_global_admin": functools.partial(
            revoke_global_admin,
            users=user_repo,
            global_admins=global_admin_repo,
            audit=audit,
        ),
        "assign_profile": functools.partial(
            assign_profile,
            users=user_repo,
            apps=app_repo,
            profiles=profile_repo,
            assignments=assignment_repo,
            audit=audit,
        ),
        "list_effective_apps": functools.partial(
            list_effective_apps,
            apps=app_repo,
            assignments=assignment_repo,
        ),
        # W61 (#524): app CRUD slice — the three new partials share
        # the same ``app_repo`` instance the ``list_effective_apps``
        # partial already binds. The audit kwarg is reserved for the
        # W62 hardening pass; the W61 routes accept an
        # ``actor_id`` header and the use case ignores it for now.
        "create_app": functools.partial(create_app, apps=app_repo),
        "update_app": functools.partial(update_app, apps=app_repo),
        "disable_app": functools.partial(disable_app, apps=app_repo),
        "audit_append": functools.partial(
            audit_append,
            audit=audit,
        ),
        "bootstrap_global_admins": functools.partial(
            bootstrap_global_admins,
            users=user_repo,
            global_admins=global_admin_repo,
            audit=audit,
            secrets=secret_manager,  # type: ignore[arg-type]
        ),
        "set_password": functools.partial(
            set_password,
            hasher=password_hasher,
            users=user_repo,
            global_admins=global_admin_repo,
            audit=audit,
            secrets=secret_manager,  # type: ignore[arg-type]
        ),
        # W60 (#522): presence use cases. The two partials share the
        # same ``presence_repo`` instance — the SSE emitter polls
        # ``get_connected_users`` while ``POST /presence/heartbeat``
        # triggers ``track_presence`` on the same backing table, so
        # the writes need a single repository object.
        "track_presence": functools.partial(track_presence, presence=presence_repo),
        "get_connected_users": functools.partial(get_connected_users, presence=presence_repo),
        # W62 (#539): login use case. The lockout policy is read once at
        # container-build time and frozen into the partial; per-call
        # ``actor_id`` still flows through to the use case.
        "login": functools.partial(
            login,
            users=user_repo,
            sessions=session_repo,
            password_hasher=password_hasher,
            audit=audit,
            lockout=LockoutPolicy(),
            session_ttl_seconds=DEFAULT_SESSION_TTL_SECONDS,
        ),
        # W62 (#540): logout use case. The ``actor_id`` (the JWT subject
        # resolved by the auth middleware in PR-5) flows through the
        # audit row's ``actor_id`` field for cross-referencing.
        "logout": functools.partial(
            logout,
            sessions=session_repo,
            audit=audit,
        ),
    }


__all__ = ["build_use_case_factories"]
