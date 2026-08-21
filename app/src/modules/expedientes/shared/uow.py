"""Unit of Work for the Expedientes module (F04, issue #225).

The UoW wraps a SQLAlchemy ``Session`` in a transactional boundary:
``__enter__`` opens a new session (or yields the ambient one if the
caller already has one), and ``__exit__`` commits on success or rolls
back on exception. It is the symmetric counterpart of the
``lanzadera`` module's D90 reset-flow transactional pattern (DA-11)
but generalised to any multi-repository mutation in the Expedientes
chain (CAP-001..007 lifecycle, CAP-008..014 related-data, CAP-015..024
catalogs, CAP-030..032 writes).

The contract:

- The UoW is a context manager. Use it with ``with``.
- Inside the block, the caller accesses the session via the ``session``
  attribute and writes through it; commit happens automatically on
  ``__exit__`` if no exception was raised.
- On any exception inside the block, the transaction is rolled back
  and the exception is re-raised.
- The UoW is **not** thread-safe: each thread must own its UoW. The
  composition root owns the ``SessionFactory`` and hands out one
  ``UnitOfWork`` per request (or per background job, for the
  canonicalisation batch — D-EXP-9 design).

Why this lives in F04 and not in M01 (extractor): every Expedientes
mutation that crosses repository boundaries needs the UoW — the
extractor (M01) is the first consumer, but C01 (create), R01..R07
(related verticals) and Q01/Q02 (catalogs) all require it. Shipping
the UoW in F04 makes M01..M04 a one-line ``with UnitOfWork(factory) as uow:``
each, instead of duplicating the session-management boilerplate per
vertical.
"""

from __future__ import annotations

from contextlib import AbstractContextManager
from typing import Protocol
from uuid import UUID


class _SessionLike(Protocol):
    """Structural Protocol — every SQLAlchemy ``Session`` satisfies it."""

    def commit(self) -> None: ...
    def rollback(self) -> None: ...
    def close(self) -> None: ...
    def execute(self, statement: object) -> object: ...
    def merge(self, instance: object) -> object: ...
    def add(self, instance: object) -> None: ...
    def delete(self, instance: object) -> None: ...
    def query(self, *entities: object) -> object: ...
    def get(self, entity: object, ident: UUID) -> object | None: ...


class SessionFactory(Protocol):
    """Structural Protocol — the composition root injects a SQLAlchemy ``sessionmaker``."""

    def __call__(self) -> _SessionLike: ...


class UnitOfWork(AbstractContextManager["_SessionLike"]):
    """Context manager that owns a transactional session.

    Use::

        with UnitOfWork(factory) as uow:
            uow.session.add(...)
            ...
        # commit happened automatically on ``__exit__``

    On exception inside the block, the transaction is rolled back and
    the session is closed; the exception is re-raised. The composition
    root never shares a ``UnitOfWork`` across requests — each request
    gets its own.
    """

    def __init__(self, factory: SessionFactory) -> None:
        self._factory = factory
        self._session: _SessionLike | None = None

    @property
    def session(self) -> _SessionLike:
        """Return the active session. Raises ``RuntimeError`` outside the block."""
        if self._session is None:
            raise RuntimeError(
                "UnitOfWork.session accessed outside the 'with' block; "
                "the session is only valid for the lifetime of the transaction."
            )
        return self._session

    def __enter__(self) -> _SessionLike:
        self._session = self._factory()
        return self._session

    def __exit__(self, exc_type: object, exc: object, tb: object) -> None:
        assert self._session is not None
        try:
            if exc_type is None:
                self._session.commit()
            else:
                self._session.rollback()
        finally:
            self._session.close()
            self._session = None
        # Do not suppress the exception; the caller must observe it.
