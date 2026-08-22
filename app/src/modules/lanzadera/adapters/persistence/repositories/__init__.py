"""Postgres-backed repository adapters (lanzadera-mvp, AD1).

Concrete adapters that satisfy the driven-port Protocols under
:mod:`app.src.modules.lanzadera.domain.ports`:

- :class:`.user_repository_pg.UserRepositoryPg` — :class:`UserRepository`
- :class:`.app_repository_pg.AppRepositoryPg` — :class:`AppRepositoryPort`
- :class:`.profile_repository_pg.ProfileRepositoryPg` — :class:`ProfileRepositoryPort`

All three share the same factory-injection pattern: the constructor
takes an :class:`AsyncSessionFactoryPort` and opens a fresh session
per call. DA-11 atomicity (``INSERT`` audit + mutation in the same
transaction) is preserved by the composition root's session reuse.
"""

from __future__ import annotations
