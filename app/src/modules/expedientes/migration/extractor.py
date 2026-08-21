"""Extractor orchestrator — M01 (issue #257) main loop.

The extractor iterates over the tables in the order returned by
``ExtractorBackend.list_tables``, reads each row in PK order via
``ExtractorBackend.read_table``, and skips rows whose PK is at or
below the watermark recorded for that table. The watermark is
updated only AFTER the staging writer returns successfully (the writer
is M02 territory — M01 ships a :class:`NullStagingWriter` that records
the write target in memory so the test can assert what WOULD land
in staging without a live Postgres).

Idempotency is the central contract (D-EXP-9): running the
extractor twice with no changes in the source produces zero rows on
the second run. The watermark store is the only state that survives
across runs; the staging writer is intentionally idempotent (the
staging tables are ``INSERT ... ON CONFLICT DO UPDATE`` in M02; M01
captures the write-target list and asserts it without performing the
write).
"""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass, field

from app.src.modules.expedientes.migration.backend import ExtractorBackend
from app.src.modules.expedientes.migration.watermark import WatermarkStore


@dataclass(frozen=True)
class StagingWrite:
    """One row's worth of staging write — the writer turns this into
    a real ``INSERT`` once M02 lands the staging layer."""

    table: str
    primary_key: int
    values: dict[str, object] = field(default_factory=dict)


class StagingWriter:
    """Sink for the extracted rows. M01 ships the in-memory variant; M02
    replaces it with the real ``INSERT INTO staging.<table> ...`` adapter."""

    def write(self, row: StagingWrite) -> None: ...


class InMemoryStagingWriter:
    """Default writer for tests + dry-runs. Records every write so the
    test can assert the exact set of rows the extractor produced."""

    def __init__(self) -> None:
        self.writes: list[StagingWrite] = []

    def write(self, row: StagingWrite) -> None:
        self.writes.append(row)


@dataclass
class ExtractionResult:
    """Summary of one extractor run. ``tables_skipped`` lists tables
    whose watermark was already at-or-above the source's max PK; the
    next run will be a no-op for them unless the source gains new rows."""

    tables_attempted: tuple[str, ...] = ()
    rows_extracted: int = 0
    rows_skipped_by_watermark: int = 0
    tables_skipped: tuple[str, ...] = ()


class Extractor:
    """Read-only copy of legacy ``Expedientes_datos.accdb`` into staging.

    The main loop is per-table: read rows in PK order, skip those at or
    below the watermark, write the rest to staging, advance the
    watermark to the highest PK written. The ``run_once`` method is the
    entry point; the caller (M01 CLI, or the test suite) decides which
    tables to extract (default: every table the backend exposes, in
    the backend's order).
    """

    def __init__(
        self,
        backend: ExtractorBackend,
        watermark: WatermarkStore,
        writer: StagingWriter,
        primary_key_column: str = "ID",
    ) -> None:
        self._backend = backend
        self._watermark = watermark
        self._writer = writer
        self._pk = primary_key_column

    def run_once(
        self,
        tables: Iterable[str] | None = None,
    ) -> ExtractionResult:
        """Run the extractor once. Returns the per-run summary.

        ``tables=None`` means "every table the backend exposes, in the
        backend's order". Pass an explicit list for a focused run
        (the M01 tranche 01 covers tables 01-15 of the 49).
        """
        attempted: list[str] = []
        rows_extracted = 0
        rows_skipped = 0
        skipped_tables: list[str] = []
        for table in tables or self._backend.list_tables():
            attempted.append(table)
            watermark: int | None = self._watermark.get(table)
            for row in self._backend.read_table(table):
                pk = self._extract_pk(row)
                if watermark is not None and pk <= watermark:
                    rows_skipped += 1
                    continue
                self._writer.write(
                    StagingWrite(
                        table=table,
                        primary_key=pk,
                        values=dict(row),
                    )
                )
                rows_extracted += 1
                watermark = pk
            if watermark is not None:
                self._watermark.put(table, watermark)
            elif rows_extracted == 0 and self._watermark.get(table) is None:
                # The table was empty AND the watermark had no entry for
                # it; the next run will re-walk the table from PK 1.
                # We record 0 explicitly to make the empty case visible
                # in the run summary (operators debugging a stuck
                # extraction need to know "this table has no rows").
                self._watermark.put(table, 0)
                skipped_tables.append(table)
        return ExtractionResult(
            tables_attempted=tuple(attempted),
            rows_extracted=rows_extracted,
            rows_skipped_by_watermark=rows_skipped,
            tables_skipped=tuple(skipped_tables),
        )

    def _extract_pk(self, row: object) -> int:
        """Pull the primary-key column out of a backend row.

        The default column is ``"ID"`` (the legacy access convention for
        every ``Tb*`` table); the constructor accepts an override for
        the few tables that use a different PK name (``IDExpediente``,
        ``IDSuministrador``, etc.).
        """
        try:
            raw_value: int | str = row[self._pk]  # type: ignore[index]
        except (KeyError, TypeError) as exc:
            row_repr = dict(row) if hasattr(row, "keys") else row  # type: ignore[call-overload]
            raise KeyError(f"row missing primary-key column {self._pk!r}: {row_repr}") from exc
        if isinstance(raw_value, int) and not isinstance(raw_value, bool):
            return raw_value
        if isinstance(raw_value, bool):
            # ``bool`` is a subclass of ``int`` but PKs are never booleans;
            # reject loudly so a hand-edited source does not silently get
            # a wrong PK.
            raise ValueError(
                f"primary key {self._pk!r} is a bool, not an int; "
                f"got {raw_value!r} — refuse to coerce"
            )
        # Access long columns come back as ``int`` from pyodbc; if a
        # backend returns a string-numeric instead (e.g. a text-encoded
        # PK in some legacy import), parse it as base-10.
        return int(raw_value)
