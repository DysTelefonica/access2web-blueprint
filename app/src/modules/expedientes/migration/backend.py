"""Protocol for the read-only Access backend the extractor talks to.

The MVP backend is the ``AccessOdbcBackend`` (pyodbc against
``Expedientes_datos.accdb``); tests use ``FakeBackend`` that returns
in-memory rows. The production swap for ``VaultRowIterator`` is out of
scope for M01 (issue #257) — that is M05 / operations work.

The contract is a single ``read_table`` method that yields rows in
PK order (so the watermark monotonicity check is meaningful — a
re-extraction of the same PK range is a no-op). The backend MUST be
read-only: any write attempt MUST raise. The extractor is on the
read path of the platform; writes go through the staging writer
(see :mod:`.staging`).
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import Protocol


class _RowLike(Protocol):
    """Structural shape every backend row satisfies (a dict-like)."""

    def __getitem__(self, key: str) -> object: ...
    def keys(self) -> Iterator[str]: ...


class ExtractorBackend(Protocol):
    """Read-only backend protocol — every implementation MUST be safe to
    point at the production ``Expedientes_datos.accdb`` without risk
    of mutation."""

    def list_tables(self) -> tuple[str, ...]:
        """Return the table names this backend exposes, in extraction order.

        The order MUST be the same on every call (the watermarker
        stores PKs in iteration order; re-ordering would invalidate
        the stored state).
        """
        ...

    def read_table(self, name: str) -> Iterator[_RowLike]:
        """Yield rows of ``name`` in PK order, ascending.

        Implementations MUST read in primary-key order so the
        watermark store can detect a "missing" PK as a stop-the-line
        anomaly (the ``Watermark`` raises on PK gap, see M01
        acceptance criteria).

        Implementations MUST be read-only. The ``FakeBackend`` raises
        on any write; the production ``AccessOdbcBackend`` opens the
        accdb with ``mode=ro`` (read-only).
        """
        ...


class AccessBackendIsReadOnlyError(RuntimeError):
    """Raised when a read-only backend is asked to write.

    The staging layer owns writes. The extractor MUST NOT bypass
    staging and write to the production schema directly.
    """
