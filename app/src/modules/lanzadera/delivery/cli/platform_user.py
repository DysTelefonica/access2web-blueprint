# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
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

W46 (#492) split the five commands into two cohesive modules:
``platform_user_auth.py`` (set-password, grant/revoke global admin)
and ``platform_user_apps.py`` (list-apps, assign-profile). This file
stays as the entry point — it owns the shared ``CommandResult``
envelope, the destructive-command exception, the confirmation
prompt + password read helpers, and the ``build_parser`` factory.
The five command callables are re-exported below so the historical
``from app.src.modules.lanzadera.delivery.cli.platform_user import cmd_*``
import surface keeps working (the test suite relies on it).
"""

from __future__ import annotations

import argparse

from app.src.modules.lanzadera.delivery.cli.platform_user_apps import (
    cmd_assign_profile,
    cmd_list_apps,
)
from app.src.modules.lanzadera.delivery.cli.platform_user_auth import (
    cmd_grant_global_admin,
    cmd_revoke_global_admin,
    cmd_set_password,
)
from app.src.modules.lanzadera.delivery.cli.platform_user_types import (  # noqa: F401
    CommandResult,
    ConfirmableDestructiveCommandError,
    _prompt_confirmation,  # reexport: tests monkeypatch this via module attribute
    _read_password,  # reexport: tests monkeypatch this via module attribute
    _resolve_user_with_confirmation,  # reexport: split modules look up via _platform_user.<name>
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
    "_prompt_confirmation",  # re-exported for tests + split-module lookup
    "_read_password",  # re-exported for tests + split-module lookup
    "_resolve_user_with_confirmation",  # re-exported for split-module lookup
]
