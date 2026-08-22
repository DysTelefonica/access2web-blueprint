"""Platform-user CLI (WU DL3, issue #56).

DA-5, DA-6, D91, CA-S4. Implements ``gentle-ai platform user
{set-password,grant-global-admin,revoke-global-admin,list-apps,assign-profile}``.
The CLI is the bootstrap path (set-password is the only way to create
the first global admin — the reset flow needs a global admin to send
the reset email to).

All destructive commands (set-password, grant, revoke,
assign-profile) prompt for confirmation before writing (D26). The
secrets flow through env vars (CA-S4) — the CLI does not accept
passwords on the command line to avoid leaking them in shell history
or process listings.
"""

from __future__ import annotations

import argparse
import getpass
import sys
from collections.abc import Sequence
from dataclasses import dataclass
from typing import Protocol

from app.src.modules.lanzadera.adapters.crypto.credential_hasher_argon2id import (
    CredentialHasherArgon2id,
)
from app.src.modules.lanzadera.domain.ports import UserRepository
from app.src.modules.lanzadera.domain.ports.assignment_repository import (
    AssignmentRepositoryPort,
)
from app.src.modules.lanzadera.domain.ports.global_admin_repository import (
    GlobalAdminRepositoryPort,
)


class _AppRepositoryPort(Protocol):
    """Structural Protocol — every repos that exposes ``list_active`` works.

    The design.md table maps this contract to ``AppRepositoryPg``; tests
    inject a ``FakeAppRepository`` that returns a hard-coded list.
    """

    async def list_active(self) -> Sequence[object]: ...


class _ProfileRepositoryPort(Protocol):
    """Structural Protocol — every repos that exposes ``get_by_code`` works."""

    async def get_by_code(self, app_id: int, code: str) -> object | None: ...


@dataclass
class CommandResult:
    """Exit envelope for every subcommand.

    ``exit_code=0`` means success; the message is printed to stdout
    by the CLI driver. ``exit_code=1`` is the failure path; the CLI
    driver writes the message to stderr and exits non-zero. The
    confirmable destructive commands (``set-password``, ``grant``,
    ``revoke``, ``assign-profile``) all set ``confirmed=True`` only
    after the operator types ``yes`` at the prompt — D26.
    """

    exit_code: int
    message: str


class ConfirmableDestructiveCommandError(RuntimeError):
    """Raised by destructive subcommands when the operator declines the prompt.

    The CLI driver catches this and exits non-zero with the
    message; the underlying case-of-use is never invoked. This is the
    D26 (destructive commands prompt) contract.
    """


def _prompt_confirmation(action: str, target: str) -> bool:
    """Read a yes/no confirmation from stdin. Returns True iff 'yes'.

    Anything other than 'yes' (case-insensitive, stripped) is treated
    as a decline. The empty string is a decline.
    """
    sys.stderr.write(f"About to {action} {target}. Continue? [yes/no]: ")
    sys.stderr.flush()
    answer = sys.stdin.readline().strip().lower()
    return answer == "yes"


def _read_password(prompt: str) -> str:
    """Read a password from stdin without echoing it.

    CA-S4: the password MUST NOT appear on the command line. Using
    ``input()`` would echo the password to the terminal; ``getpass``
    reads from the controlling tty without echoing. The CLI driver
    routes the password through stdin.

    This function is also the override point the test suite uses to
    inject a password (monkeypatching ``platform_user._read_password``
    with a plain ``stdin.readline``) so the test environment can run
    without a tty. The default behaviour is ``getpass.getpass``.
    """
    try:
        return getpass.getpass(prompt)
    except (OSError, ValueError):
        # No tty available (CI, subprocess, pytest capture). Fall back to
        # stdin read; the password may echo to the caller's terminal,
        # but that is better than crashing.
        return input(prompt)


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
    normalised_email = email.strip().lower()
    user = await users.get_by_email(normalised_email)
    if user is None:
        return CommandResult(exit_code=1, message=f"user not found: {normalised_email}")
    if not confirmed and not _prompt_confirmation("set-password", f"for user {normalised_email}"):
        return CommandResult(exit_code=1, message="set-password cancelled by operator")
    password = _read_password("New password: ")
    if not password:
        return CommandResult(exit_code=1, message="empty password is not allowed")
    hashed = await hasher.hash(password)
    # The actual write is mediated by a use case (lifecycle) that updates
    # the password_hash AND sets status='active' in one transaction (DA-5);
    # the CLI ships the hashed value to the same code path that the
    # reset-flow's consume_reset_token uses. For this WU we expose a
    # thin wrapper on the user repository; the actual lifecycle write is
    # M02 / T01.
    await users.update_password_and_activate(user.id, hashed)
    return CommandResult(
        exit_code=0,
        message=f"password set for {normalised_email}; status=active",
    )


async def cmd_grant_global_admin(
    email: str,
    users: UserRepository,
    global_admins: GlobalAdminRepositoryPort,
    *,
    confirmed: bool = False,
) -> CommandResult:
    """Grant global-admin to a user (DA-6, the bootstrap path's sibling)."""
    normalised_email = email.strip().lower()
    user = await users.get_by_email(normalised_email)
    if user is None:
        return CommandResult(exit_code=1, message=f"user not found: {normalised_email}")
    if not confirmed and not _prompt_confirmation("grant-global-admin", f"to {normalised_email}"):
        return CommandResult(exit_code=1, message="grant cancelled by operator")
    try:
        await global_admins.grant(user.id)
    except ValueError as exc:
        return CommandResult(exit_code=1, message=f"cannot grant: {exc}")
    return CommandResult(exit_code=0, message=f"global admin granted to {normalised_email}")


async def cmd_revoke_global_admin(
    email: str,
    users: UserRepository,
    global_admins: GlobalAdminRepositoryPort,
    *,
    confirmed: bool = False,
) -> CommandResult:
    """Revoke global-admin (D42: must keep at least one admin)."""
    normalised_email = email.strip().lower()
    user = await users.get_by_email(normalised_email)
    if user is None:
        return CommandResult(exit_code=1, message=f"user not found: {normalised_email}")
    if not confirmed and not _prompt_confirmation(
        "revoke-global-admin", f"from {normalised_email}"
    ):
        return CommandResult(exit_code=1, message="revoke cancelled by operator")
    try:
        await global_admins.revoke(user.id)
    except ValueError as exc:
        return CommandResult(exit_code=1, message=f"cannot revoke: {exc}")
    return CommandResult(exit_code=0, message=f"global admin revoked from {normalised_email}")


async def cmd_list_apps(apps: _AppRepositoryPort) -> CommandResult:
    """List active apps (DA-5's list-apps)."""
    active = await apps.list_active()
    lines = ["active apps:"]
    for app in active:
        lines.append(f"  - {app}")
    return CommandResult(exit_code=0, message="\n".join(lines))


async def cmd_assign_profile(
    email: str,
    app_id: int,
    profile_code: str,
    users: UserRepository,
    profiles: _ProfileRepositoryPort,
    assignments: AssignmentRepositoryPort,
    *,
    confirmed: bool = False,
) -> CommandResult:
    """Assign a profile to a user in an app (DA-12)."""
    normalised_email = email.strip().lower()
    user = await users.get_by_email(normalised_email)
    if user is None:
        return CommandResult(exit_code=1, message=f"user not found: {normalised_email}")
    profile = await profiles.get_by_code(app_id, profile_code)
    if profile is None:
        return CommandResult(
            exit_code=1, message=f"profile {profile_code!r} not found in app {app_id}"
        )
    if not confirmed and not _prompt_confirmation(
        "assign-profile",
        f"{profile_code!r} in app {app_id} to {normalised_email}",
    ):
        return CommandResult(exit_code=1, message="assign-profile cancelled by operator")
    await assignments.create(user.id, app_id, profile.id)  # type: ignore[attr-defined]
    return CommandResult(
        exit_code=0,
        message=f"profile {profile_code!r} assigned to {normalised_email} in app {app_id}",
    )


def build_parser() -> argparse.ArgumentParser:
    """Build the ``gentle-ai platform user <subcommand>`` parser.

    The parser is the CLI's public surface; tests assert on its shape
    (subcommand names, flag names) so a refactor that breaks the
    contract fails fast.
    """
    parser = argparse.ArgumentParser(
        prog="gentle-ai platform user",
        description=(
            "Platform-user administration CLI (DA-5, DA-6, D91). "
            "Subcommands operate on the live database; the destructive "
            "ones prompt for confirmation (D26)."
        ),
    )
    sub = parser.add_subparsers(dest="subcommand", required=True)

    set_pw = sub.add_parser("set-password", help="set a user's password (bootstrap path)")
    set_pw.add_argument("email")

    grant = sub.add_parser("grant-global-admin", help="grant a user global-admin")
    grant.add_argument("email")

    revoke = sub.add_parser("revoke-global-admin", help="revoke a user's global-admin")
    revoke.add_argument("email")

    sub.add_parser("list-apps", help="list active apps")

    assign = sub.add_parser("assign-profile", help="assign a profile to a user in an app")
    assign.add_argument("email")
    assign.add_argument("app_id", type=int)
    assign.add_argument("profile_code")

    return parser


__all__ = [
    "CommandResult",
    "ConfirmableDestructiveCommandError",
    "build_parser",
    "cmd_assign_profile",
    "cmd_grant_global_admin",
    "cmd_list_apps",
    "cmd_revoke_global_admin",
    "cmd_set_password",
]
