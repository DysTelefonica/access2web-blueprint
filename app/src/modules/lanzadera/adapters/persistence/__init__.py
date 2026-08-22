"""Postgres persistence adapters for the D90 reset-flow surface.

W01 (#44) ships three adapters and the factory that produces the
``AsyncSession`` seam:

- :func:`async_session_factory` builds an ``AsyncEngine`` + an
  :class:`AsyncSessionFactoryPort` whose ``search_path`` is set to
  ``lanzadera,public`` (the canonical schema, with public as the
  fallback for cross-schema lookups).
- :class:`UserRepositoryPg`, :class:`AppRepositoryPg`, and
  :class:`ProfileRepositoryPg` implement the corresponding domain
  Protocols (DA-1, DA-12, D90).

The three adapters share the same construction shape:

    factory = AsyncSessionFactoryPort(...)  # built by `async_session_factory`
    repo    = UserRepositoryPg(factory)

The factory is the only seam; the adapters do not import the engine
directly. Adapters open and close an ``AsyncSession`` per call
(matching the AsyncSession lifetime contract declared in
``async_session_factory``).
"""

from .async_session_factory import (
    AsyncSessionFactory,
    AsyncSessionFactoryError,
    AsyncSessionFactoryPort,
    async_session_factory,
)
from .repositories import (
    APPS_TABLE,
    PROFILES_TABLE,
    USERS_TABLE,
    ProfileRepositoryPg,
    UserRepositoryPg,
)

__all__ = [
    "APPS_TABLE",
    "AsyncSessionFactory",
    "AsyncSessionFactoryError",
    "AsyncSessionFactoryPort",
    "PROFILES_TABLE",
    "ProfileRepositoryPg",
    "USERS_TABLE",
    "UserRepositoryPg",
    "async_session_factory",
]
