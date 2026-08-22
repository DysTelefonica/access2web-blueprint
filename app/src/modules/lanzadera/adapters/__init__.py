"""Adapters implementing :mod:`app.src.modules.lanzadera.ports` (driven).

* :class:`MailQueueTableAdapter` (re-exported from :mod:`.persistence`) —
  DA-10, D11, D13, D65, P20. ``mail_outbox`` row, status default
  'pending', dispatched by a future cron / dispatcher (out of scope
  for this slice).
* :class:`.bootstrap.EnvAdminSourceAdapter` — DA-6. Reads
  ``GLOBAL_ADMIN_EMAILS`` (semicolon-separated) at process startup;
  empty if unset.
* :class:`.repos.AssignmentRepositoryPg` — DA-12, D22, D42, H11.
* :class:`.repos.ResetTokenRepositoryPg` — D90, DA-4.
* :class:`.repos.GlobalAdminRepositoryPg` — D21, D42.
* :class:`.repos.AuditLogPg` — D27, DA-11, D55.
"""

from __future__ import annotations

from app.src.modules.lanzadera.adapters.bootstrap.env_admin_source_adapter import (
    EnvAdminSourceAdapter,
)
from app.src.modules.lanzadera.adapters.persistence import (
    # W02 (#45), W03 (#55), W04 (#21), W05 (#42-subset): all four
    # Postgres adapters migrated from the legacy repos/ subpackage.
    AssignmentRepositoryPg,
    AuditLogPg,
    GlobalAdminRepositoryPg,
    MailQueueTableAdapter,  # W06 (#47) -- migrated from notification/
    ResetTokenRepositoryPg,
)

__all__ = [
    "AssignmentRepositoryPg",
    "AuditLogPg",
    "EnvAdminSourceAdapter",
    "GlobalAdminRepositoryPg",
    "MailQueueTableAdapter",
    "ResetTokenRepositoryPg",
]
