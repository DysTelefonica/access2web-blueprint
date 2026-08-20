"""SQLAlchemy Core + asyncpg implementations of the four repository ports.

**Status (lanzadera-mvp WU AD2, #45):** the four adapter classes below
are declared with the exact contract from
``app.src.modules.lanzadera.domain.ports`` and the exact SQL semantics
from ``openspec/changes/lanzadera-mvp/design.md``. The method bodies
are intentionally ``NotImplementedError`` because the underlying
Postgres schema (migrations ``0003_seed_assignments`` / ``0006_seed_audit``
and the ``reset_tokens`` / ``global_admins`` DDL) does not yet exist
on ``main`` — only ``0001_core_schema`` has landed. Implementing these
methods without the schema in place would mean writing SQL against
tables that do not exist yet, and the orchestrator policy is to keep
the diff reviewable at <=320 lines per WU (CONTRIBUTING.md §Tamaño).

The follow-up WU (after migrations ``0003`` + ``0004`` + ``0006`` land)
will fill in the bodies without changing the contract — tests pinned
against this skeleton today will pass against the real adapter tomorrow
(structural subtyping, ``typing.Protocol``).
"""

from __future__ import annotations

from app.src.modules.lanzadera.adapters.repos.assignment_repository_pg import (
    AssignmentRepositoryPg,
)
from app.src.modules.lanzadera.adapters.repos.audit_log_pg import AuditLogPg
from app.src.modules.lanzadera.adapters.repos.global_admin_repository_pg import (
    GlobalAdminRepositoryPg,
)
from app.src.modules.lanzadera.adapters.repos.reset_token_repository_pg import (
    ResetTokenRepositoryPg,
)

__all__ = [
    "AssignmentRepositoryPg",
    "AuditLogPg",
    "GlobalAdminRepositoryPg",
    "ResetTokenRepositoryPg",
]
