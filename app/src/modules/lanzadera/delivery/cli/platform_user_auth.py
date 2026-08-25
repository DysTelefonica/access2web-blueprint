# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Auth-flow CLI commands (set-password, grant/revoke global admin).

W46 (#492) split this out of ``platform_user.py`` so each cohesive
slice stays under the mutation-sites ceiling. The auth-flow
commands share the user-resolution + confirmation pattern that
``_resolve_user_with_confirmation`` encapsulates; they all need the
``UserRepository`` dependency and either a hasher (set-password) or a
``GlobalAdminRepositoryPort`` (grant/revoke).

The split keeps the original ``platform_user.py`` as the entry point
that re-exports these commands for backward compatibility (the test
suite imports them as e.g.
``from ...platform_user import cmd_set_password``).
"""

from __future__ import annotations

from app.src.modules.lanzadera.adapters.crypto.credential_hasher_argon2id import (
    CredentialHasherArgon2id,
)
from app.src.modules.lanzadera.delivery.cli import platform_user as _platform_user
from app.src.modules.lanzadera.delivery.cli.platform_user_types import (
    CommandResult,
    _resolve_user_with_confirmation,
)
from app.src.modules.lanzadera.domain.ports import UserRepository
from app.src.modules.lanzadera.domain.ports.global_admin_repository import (
    GlobalAdminRepositoryPort,
)


async def cmd_set_password(
    email: str,
    users: UserRepository,
    hasher: CredentialHasherArgon2id,
    *,
    confirmed: bool = False,
) -> CommandResult:
    """Set the password for a user (D91, bootstrap path).

    Looks up the user by email (lowercased per DA-3); reads the new
    password from stdin (CA-S4); hashes it (Argon2id, profile
    ``RFC_9106_LOW_MEMORY``); and updates the user record to
    ``status='active'``. The destructive operation prompts for
    confirmation unless ``confirmed=True`` (the latter is what the
    test suite uses).
    """
    resolved = await _resolve_user_with_confirmation(
        email,
        users,
        action="set-password",
        description=f"for user {email.strip().lower()}",
        cancellation_message="set-password cancelled by operator",
        confirmed=confirmed,
    )
    if isinstance(resolved, CommandResult):
        return resolved
    user = resolved
    # Lookup via the platform_user module (not the original binding) so
    # tests that monkeypatch ``platform_user._read_password`` propagate.
    password = _platform_user._read_password("New password: ")
    if not password:
        return CommandResult(exit_code=1, message="empty password is not allowed")
    hashed = await hasher.hash(password)
    await users.update_password_and_activate(user.id, hashed)
    return CommandResult(
        exit_code=0,
        message=f"password set for {user.email}; status=active",
    )


async def cmd_grant_global_admin(
    email: str,
    users: UserRepository,
    global_admins: GlobalAdminRepositoryPort,
    *,
    confirmed: bool = False,
) -> CommandResult:
    """Grant global-admin to a user (DA-6, the bootstrap path's sibling)."""
    resolved = await _resolve_user_with_confirmation(
        email,
        users,
        action="grant-global-admin",
        description=f"to {email.strip().lower()}",
        cancellation_message="grant cancelled by operator",
        confirmed=confirmed,
    )
    if isinstance(resolved, CommandResult):
        return resolved
    user = resolved
    try:
        await global_admins.grant(user.id)
    except ValueError as exc:
        return CommandResult(exit_code=1, message=f"cannot grant: {exc}")
    return CommandResult(exit_code=0, message=f"global admin granted to {user.email}")


async def cmd_revoke_global_admin(
    email: str,
    users: UserRepository,
    global_admins: GlobalAdminRepositoryPort,
    *,
    confirmed: bool = False,
) -> CommandResult:
    """Revoke global-admin (D42: must keep at least one admin)."""
    resolved = await _resolve_user_with_confirmation(
        email,
        users,
        action="revoke-global-admin",
        description=f"from {email.strip().lower()}",
        cancellation_message="revoke cancelled by operator",
        confirmed=confirmed,
    )
    if isinstance(resolved, CommandResult):
        return resolved
    user = resolved
    try:
        await global_admins.revoke(user.id)
    except ValueError as exc:
        return CommandResult(exit_code=1, message=f"cannot revoke: {exc}")
    return CommandResult(exit_code=0, message=f"global admin revoked from {user.email}")


__all__ = [
    "cmd_grant_global_admin",
    "cmd_revoke_global_admin",
    "cmd_set_password",
]
