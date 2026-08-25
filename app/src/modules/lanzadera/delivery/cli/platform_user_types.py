# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""Shared CLI primitives for the platform-user subcommands.

W46 (#492) split the five commands across ``platform_user_auth.py``
and ``platform_user_apps.py``. To avoid a circular import (each split
module needs ``CommandResult`` + the protocols; the entry point needs
to re-export the command callables), the shared types AND the
stdio-based helpers (``_prompt_confirmation``, ``_read_password``) live
here and both sides import from this leaf module.

The entry point ``platform_user.py`` re-exports these symbols
unchanged so the historical
``from app.src.modules.lanzadera.delivery.cli.platform_user import CommandResult``
import surface keeps working.
"""

from __future__ import annotations

import getpass
import sys
from collections.abc import Sequence
from dataclasses import dataclass
from typing import Protocol

from app.src.modules.lanzadera.domain.ports import UserRepository
from app.src.modules.lanzadera.domain.user import User


class _AppRepositoryPort(Protocol):
    """Structural Protocol — what the template needs from an app row.

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


async def _resolve_user_with_confirmation(
    email: str,
    users: UserRepository,
    *,
    action: str,
    description: str,
    cancellation_message: str,
    confirmed: bool,
) -> User | CommandResult:
    """Look up the user by ``email`` (lowercased per DA-3) and gate on confirmation.

    Returns the looked-up ``User`` on success; returns a ``CommandResult``
    with ``exit_code=1`` if the user is absent or the operator declined
    the confirmation prompt (D26). Shared by ``cmd_set_password``,
    ``cmd_grant_global_admin``, and ``cmd_revoke_global_admin`` — the
    five-statement lookup+confirm block they used to repeat.
    """
    normalised_email = email.strip().lower()
    user = await users.get_by_email(normalised_email)
    if user is None:
        return CommandResult(exit_code=1, message=f"user not found: {normalised_email}")
    if not confirmed and not _prompt_confirmation(action, description):
        return CommandResult(exit_code=1, message=cancellation_message)
    return user


__all__ = [
    "CommandResult",
    "ConfirmableDestructiveCommandError",
    "_AppRepositoryPort",
    "_ProfileRepositoryPort",
    "_prompt_confirmation",
    "_read_password",
    "_resolve_user_with_confirmation",
]
