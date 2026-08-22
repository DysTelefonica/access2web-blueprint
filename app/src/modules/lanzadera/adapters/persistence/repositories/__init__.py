"""Postgres adapters — one module per domain Protocol.

W01 (#44) ships the three Postgres adapters of the user-port trio:
``UserRepositoryPg``, ``AppRepositoryPg``, ``ProfileRepositoryPg``.
Each module declares its own ``sa.Table`` reflection object so the
adapter self-contains the schema it persists against.

W02 (#45) adds ``AssignmentRepositoryPg``
(``AssignmentRepositoryPort``, DA-12, D22, D42, H11).
"""

from .app_repository_pg import APPS_TABLE, AppRepositoryPg
from .assignment_repository_pg import AssignmentRepositoryPg
from .audit_log_pg import AUDIT_TABLE, AuditLogPg
from .profile_repository_pg import PROFILES_TABLE, ProfileRepositoryPg
from .user_repository_pg import USERS_TABLE, UserRepositoryPg

__all__ = [
    "APPS_TABLE",
    "AUDIT_TABLE",
    "PROFILES_TABLE",
    "USERS_TABLE",
    "AppRepositoryPg",
    "AssignmentRepositoryPg",
    "AuditLogPg",
    "ProfileRepositoryPg",
    "UserRepositoryPg",
]
