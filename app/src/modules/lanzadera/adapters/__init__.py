"""Adapters implementing :mod:`app.src.modules.lanzadera.ports` (driven).

* :class:`MailQueueTableAdapter` (re-exported from :mod:`.persistence`) — DA-10, D11, D13, D65, P20.
  ``mail_outbox`` row, status default ``pending``. Vacío si la variable no está definida.
* :class:`.bootstrap.EnvAdminSourceAdapter` — DA-6. Lee ``GLOBAL_ADMIN_EMAILS`` al arrancar.
* :class:`.persistence.AssignmentRepositoryPg` — DA-12, D22, D42, H11.
* :class:`.persistence.ResetTokenRepositoryPg` — D90, DA-4.
* :class:`.persistence.GlobalAdminRepositoryPg` — D21, D42.
* :class:`.persistence.AuditLogPg` — D27, DA-11, D55.

W01..W20 consolidados en :mod:`.persistence`. Los 8 adapters Postgres comparten
el helper ``AsyncSessionFactory`` (DA-1). ``read_only_session()`` cubre lectura;
``transaction()`` cubre escritura.
``read_only_session()`` cubre lectura; ``transaction()`` cubre escritura.
"""

from __future__ import annotations

from app.src.modules.lanzadera.adapters.bootstrap.env_admin_source_adapter import (
    EnvAdminSourceAdapter,
)
from app.src.modules.lanzadera.adapters.persistence import (
    AssignmentRepositoryPg,
    AuditLogPg,
    GlobalAdminRepositoryPg,
    MailQueueTableAdapter,
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
