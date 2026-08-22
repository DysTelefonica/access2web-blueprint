"""Postgres adapters — one module per domain Protocol.

W01 (#44) ships the three Postgres adapters: ``UserRepositoryPg``
(``UserRepository``), ``AppRepositoryPg`` (``AppRepositoryPort``),
``ProfileRepositoryPg`` (``ProfileRepositoryPort``). Each module
declares its own ``sa.Table`` reflection object so the adapter
self-contains the schema it persists against.
"""

from .app_repository_pg import APPS_TABLE, AppRepositoryPg
from .profile_repository_pg import PROFILES_TABLE, ProfileRepositoryPg
from .user_repository_pg import USERS_TABLE, UserRepositoryPg

__all__ = [
    "APPS_TABLE",
    "PROFILES_TABLE",
    "USERS_TABLE",
    "AppRepositoryPg",
    "ProfileRepositoryPg",
    "UserRepositoryPg",
]
