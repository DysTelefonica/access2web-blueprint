"""Adapters implementing :mod:`app.src.modules.lanzadera.ports` (driven).

* :class:`.notification.MailQueueTableAdapter` — DA-10, D11, D13, D65, P20.
  ``mail_outbox`` row, status default 'pending', dispatched by a
  future cron / dispatcher (out of scope for this slice).
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
from app.src.modules.lanzadera.adapters.notification.mail_queue_table_adapter import (
    MailQueueTableAdapter,
)
from app.src.modules.lanzadera.adapters.persistence import (
    AssignmentRepositoryPg,  # W02 (#45) -- migrated from repos/
    AuditLogPg,
    GlobalAdminRepositoryPg,
)  # W03 (#55) -- migrated from repos/  # W04 (#21) -- migrated from repos/
from app.src.modules.lanzadera.adapters.repos.reset_token_repository_pg import (
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
