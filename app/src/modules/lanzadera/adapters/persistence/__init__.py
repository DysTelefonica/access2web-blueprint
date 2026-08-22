"""Postgres persistence adapters for the D90 + W-series surface.

W01 (#44) ships the seam:
- :func:`async_session_factory` builds an ``AsyncEngine`` + an
  :class:`AsyncSessionFactoryPort` whose ``search_path`` is set to
  ``lanzadera,public`` (the canonical schema, with public as the
  fallback for cross-schema lookups).

W01 also ships the three Postgres adapters of the user-port trio:
:class:`UserRepositoryPg`, :class:`AppRepositoryPg`, and
:class:`ProfileRepositoryPg` (DA-1, DA-12).

W02 (#45) adds :class:`AssignmentRepositoryPg` on top.

All adapters share the construction shape:

    factory = async_session_factory(url)
    repo    = UserRepositoryPg(factory)

Adapters open and close an ``AsyncSession`` per call. None of them
import an engine directly — the factory is the only seam.
"""

from .async_session_factory import (
    AsyncSessionFactory,
    AsyncSessionFactoryError,
    AsyncSessionFactoryPort,
    async_session_factory,
)
from .repositories import (
    APPS_TABLE,
    AUDIT_TABLE,
    GLOBAL_ADMINS_TABLE,
    PROFILES_TABLE,
    RESET_TOKENS_TABLE,
    USERS_TABLE,
    AssignmentRepositoryPg,
    AuditLogPg,
    GlobalAdminRepositoryPg,
    ProfileRepositoryPg,
    ResetTokenRepositoryPg,
    UserRepositoryPg,
)

__all__ = [
    "APPS_TABLE",
    "AUDIT_TABLE",
    "AsyncSessionFactory",
    "AsyncSessionFactoryError",
    "AsyncSessionFactoryPort",
    "GLOBAL_ADMINS_TABLE",
    "PROFILES_TABLE",
    "RESET_TOKENS_TABLE",
    "USERS_TABLE",
    "AssignmentRepositoryPg",
    "AuditLogPg",
    "GlobalAdminRepositoryPg",
    "ProfileRepositoryPg",
    "ResetTokenRepositoryPg",
    "UserRepositoryPg",
    "async_session_factory",
]
