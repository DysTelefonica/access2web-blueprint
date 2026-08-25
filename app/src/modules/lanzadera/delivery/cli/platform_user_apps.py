# HARNESS-PROVENANCE: deterministic-quality-harness v1.6 + lanzadera-mvp
"""App-flow CLI commands (list-apps, assign-profile).

W46 (#492) split this out of ``platform_user.py`` so each cohesive
slice stays under the mutation-sites ceiling. App-flow commands are
read-mostly (list-apps) or do their own user lookup via the
``UserRepository`` rather than going through
``_resolve_user_with_confirmation`` (assign-profile has its own
confirmation flow that surfaces the profile code in the prompt).

The split keeps the original ``platform_user.py`` as the entry point
that re-exports these commands for backward compatibility (the test
suite imports them as e.g.
``from ...platform_user import cmd_list_apps``).
"""

from __future__ import annotations

from app.src.modules.lanzadera.delivery.cli import platform_user as _platform_user
from app.src.modules.lanzadera.delivery.cli.platform_user_types import (
    CommandResult,
    _AppRepositoryPort,
    _ProfileRepositoryPort,
)
from app.src.modules.lanzadera.domain.ports import UserRepository
from app.src.modules.lanzadera.domain.ports.assignment_repository import (
    AssignmentRepositoryPort,
)


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
    if not confirmed and not _platform_user._prompt_confirmation(
        "assign-profile",
        f"{profile_code!r} in app {app_id} to {normalised_email}",
    ):
        return CommandResult(exit_code=1, message="assign-profile cancelled by operator")
    await assignments.create(user.id, app_id, profile.id)  # type: ignore[attr-defined]
    return CommandResult(
        exit_code=0,
        message=f"profile {profile_code!r} assigned to {normalised_email} in app {app_id}",
    )


__all__ = [
    "cmd_assign_profile",
    "cmd_list_apps",
]
