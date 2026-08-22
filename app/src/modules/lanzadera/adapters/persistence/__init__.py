"""Postgres-backed persistence adapters (lanzadera-mvp, AD1).

The factory in :mod:`.async_session_factory` is the only seam between
this module and the SQLAlchemy driver. The repositories in
:mod:`.repositories` consume sessions through
:class:`AsyncSessionFactoryPort` (a Protocol that matches
``async_sessionmaker.__call__`` structurally), so the use-case layer can
ask for ``async with with_session(factory) as session:`` without
importing SQLAlchemy directly.
"""

from __future__ import annotations

from app.src.modules.lanzadera.adapters.persistence.async_session_factory import (
    AsyncSessionFactoryPort,
    async_session_factory,
    with_session,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.app_repository_pg import (
    AppRepositoryPg,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.profile_repository_pg import (
    ProfileRepositoryPg,
)
from app.src.modules.lanzadera.adapters.persistence.repositories.user_repository_pg import (
    UserRepositoryPg,
)

__all__ = [
    "AppRepositoryPg",
    "AsyncSessionFactoryPort",
    "ProfileRepositoryPg",
    "UserRepositoryPg",
    "async_session_factory",
    "with_session",
]
