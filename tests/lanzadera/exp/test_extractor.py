"""Tests for the M01 (issue #257) extractor orchestrator.

The tests cover the idempotency contract (D-EXP-9): a re-run of the
extractor with the same backend produces zero new rows once the
watermark catches up. They also cover the basic shape of the
extractor's output (``ExtractionResult``), the in-memory staging
writer's append-only contract, and the JSON file watermark store's
atomic-write behaviour.

The extractor is a structural walker, not a metric: it neither
consumes nor invalidates the numbers the ``quality_report.py`` gates
track. The tests are pure-Python — no database, no filesystem (the
JSON watermark is written to a temp dir via ``tmp_path``).
"""

from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path

import pytest

from app.src.modules.expedientes.migration.extractor import (
    Extractor,
    InMemoryStagingWriter,
)
from app.src.modules.expedientes.migration.watermark import (
    InMemoryWatermarkStore,
    JsonFileWatermarkStore,
)

# ---------------------------------------------------------------------------
# Test fixtures: an in-memory backend that returns canned rows.
# ---------------------------------------------------------------------------


class _Row(dict):
    """A dict that supports the two backend row operations we use:
    ``row[col]`` and ``row.keys()``."""

    def __init__(self, *args: object, **kwargs: object) -> None:
        super().__init__(*args, **kwargs)


class FakeBackend:
    """In-memory backend that returns rows in the order they were added.

    The backend exposes a single primary-key column name (``ID`` by
    default; the extractor accepts an override). Rows MUST be returned
    in PK order — the orchestrator assumes ascending order so the
    watermark can detect gaps as stop-the-line anomalies.
    """

    def __init__(
        self,
        tables_with_rows: dict[str, list[dict[str, object]]],
        primary_key: str = "ID",
    ) -> None:
        self._tables = tables_with_rows
        self._pk = primary_key

    def list_tables(self) -> tuple[str, ...]:
        return tuple(self._tables)

    def read_table(self, name: str) -> Iterator[_Row]:
        rows = sorted(
            self._tables.get(name, []),
            key=lambda r: int(r[self._pk]),  # type: ignore[arg-type]
        )
        for row in rows:
            yield _Row(row)


# ---------------------------------------------------------------------------
# Backward-compat tests — the contract the verticals depend on.
# ---------------------------------------------------------------------------


def test_extractor_writes_one_row_per_source_row() -> None:
    backend = FakeBackend({"TbEstados": [{"ID": 1, "Estado": "A"}, {"ID": 2, "Estado": "B"}]})
    writer = InMemoryStagingWriter()
    Extractor(backend, InMemoryWatermarkStore(), writer).run_once()

    assert [w.primary_key for w in writer.writes] == [1, 2]
    assert all(w.table == "TbEstados" for w in writer.writes)


def test_extractor_is_idempotent_with_watermark() -> None:
    """The central contract (D-EXP-9): re-running produces zero new rows."""
    backend = FakeBackend({"TbEstados": [{"ID": 1}, {"ID": 2}, {"ID": 3}]})
    watermark = InMemoryWatermarkStore()
    writer = InMemoryStagingWriter()

    first = Extractor(backend, watermark, writer).run_once()
    second = Extractor(backend, watermark, writer).run_once()

    assert first.rows_extracted == 3
    assert second.rows_extracted == 0
    assert second.rows_skipped_by_watermark == 3
    # The writer is shared across runs (in-memory) but the second run
    # should not have appended to it.
    assert len(writer.writes) == 3


def test_extractor_watermark_advances_to_max_pk() -> None:
    backend = FakeBackend({"TbEstados": [{"ID": 1}, {"ID": 5}, {"ID": 3}]})
    watermark = InMemoryWatermarkStore()
    writer = InMemoryStagingWriter()

    Extractor(backend, watermark, writer).run_once()
    assert watermark.get("TbEstados") == 5


def test_extractor_empty_table_records_zero_watermark() -> None:
    """An empty table records a 0 watermark so the next run is a no-op
    AND the empty case is visible in the run summary."""
    backend = FakeBackend({"TbEstados": []})
    watermark = InMemoryWatermarkStore()
    writer = InMemoryStagingWriter()

    result = Extractor(backend, watermark, writer).run_once()

    assert result.rows_extracted == 0
    assert "TbEstados" in result.tables_skipped
    assert watermark.get("TbEstados") == 0


def test_extractor_partial_watermark_resumes_from_correct_pk() -> None:
    """A pre-existing watermark from a previous (crashed) run is honoured."""
    backend = FakeBackend({"TbEstados": [{"ID": 1}, {"ID": 2}, {"ID": 3}, {"ID": 4}]})
    watermark = InMemoryWatermarkStore()
    watermark.put("TbEstados", 2)  # prior run made it to PK=2
    writer = InMemoryStagingWriter()

    result = Extractor(backend, watermark, writer).run_once()

    assert result.rows_extracted == 2
    assert [w.primary_key for w in writer.writes] == [3, 4]
    assert watermark.get("TbEstados") == 4


def test_extractor_respects_table_subset() -> None:
    """Passing ``tables`` limits the run to that subset (M01 tranche 01 covers 01-15)."""
    backend = FakeBackend(
        {
            "TbEstados": [{"ID": 1}],
            "TbComerciales": [{"ID": 1}],
            "TbCPV": [{"ID": 1}],
        }
    )
    watermark = InMemoryWatermarkStore()
    writer = InMemoryStagingWriter()

    result = Extractor(backend, watermark, writer).run_once(tables=("TbEstados", "TbCPV"))

    assert result.tables_attempted == ("TbEstados", "TbCPV")
    assert [w.table for w in writer.writes] == ["TbEstados", "TbCPV"]


def test_extractor_custom_primary_key_column() -> None:
    """Some legacy tables use a different PK name (``IDExpediente``,
    ``IDSuministrador``, ...). The constructor accepts an override."""
    backend = FakeBackend(
        {"TbExpedientes": [{"IDExpediente": 100, "Titulo": "First"}]},
        primary_key="IDExpediente",
    )
    writer = InMemoryStagingWriter()

    Extractor(
        backend,
        InMemoryWatermarkStore(),
        writer,
        primary_key_column="IDExpediente",
    ).run_once()

    assert writer.writes[0].primary_key == 100


def test_extractor_missing_pk_raises() -> None:
    """A row that lacks the primary-key column is a stop-the-line anomaly."""
    backend = FakeBackend({"TbEstados": [{"Estado": "no PK here"}]})
    writer = InMemoryStagingWriter()
    extractor = Extractor(backend, InMemoryWatermarkStore(), writer)

    # The KeyError carries the column name in its first arg; assert via
    # exception.args[0] rather than a regex on the message (the message
    # is reformatted by mypy/the test runner and is brittle to assert).
    with pytest.raises(KeyError) as exc_info:
        extractor.run_once()
    assert exc_info.value.args[0] == "ID"


# ---------------------------------------------------------------------------
# Watermark store tests.
# ---------------------------------------------------------------------------


def test_watermark_inmemory_basic() -> None:
    store = InMemoryWatermarkStore()
    assert store.get("Tb") is None
    store.put("Tb", 42)
    assert store.get("Tb") == 42
    assert store.snapshot() == {"Tb": 42}


def test_watermark_jsonfile_atomic_write(tmp_path: Path) -> None:
    """The JSON file is written atomically (write+rename) so a crash
    mid-write cannot leave a half-written file behind."""
    path = tmp_path / "watermark.json"
    store = JsonFileWatermarkStore(path)
    store.put("TbEstados", 7)
    store.put("TbComerciales", 12)
    assert path.exists()
    reloaded = JsonFileWatermarkStore(path)
    assert reloaded.get("TbEstados") == 7
    assert reloaded.get("TbComerciales") == 12


def test_watermark_jsonfile_missing_returns_empty() -> None:
    """A first run with no existing file is a no-op, not an error."""
    store = JsonFileWatermarkStore(Path("/nonexistent/wm.json"))
    assert store.snapshot() == {}


def test_watermark_jsonfile_corrupt_raises() -> None:
    """A hand-edited or half-written watermark file is a HARD error, not
    a silent default — the operator must restore from backup rather than
    the extractor guessing the safe state (D-EXP-9: "manifest/watermarks,
    plan y reporte de reconciliación")."""
    import json

    path = Path("/tmp/wm-corrupt.json")
    path.write_text("{not valid json", encoding="utf-8")
    with pytest.raises(json.JSONDecodeError):
        JsonFileWatermarkStore(path)
    path.unlink()


def test_watermark_jsonfile_non_object_raises() -> None:
    """A watermark file that does not parse as a JSON object is malformed."""
    path = Path("/tmp/wm-array.json")
    path.write_text("[1, 2, 3]", encoding="utf-8")
    with pytest.raises(ValueError, match="malformed"):
        JsonFileWatermarkStore(path)
    path.unlink()


# ---------------------------------------------------------------------------
# Smoke test for the orchestrator: end-to-end with the JSON file store.
# ---------------------------------------------------------------------------


def test_extractor_end_to_end_with_jsonfile_watermark(tmp_path: Path) -> None:
    """Two runs against a JSON file watermark, the second a no-op."""
    backend = FakeBackend({"TbEstados": [{"ID": 1}, {"ID": 2}]})
    wm_path = tmp_path / "wm.json"
    writer = InMemoryStagingWriter()

    first = Extractor(backend, JsonFileWatermarkStore(wm_path), writer).run_once()
    second = Extractor(backend, JsonFileWatermarkStore(wm_path), writer).run_once()

    assert first.rows_extracted == 2
    assert second.rows_extracted == 0
    assert second.rows_skipped_by_watermark == 2
    # The JSON file was written after the first run.
    assert wm_path.read_text(encoding="utf-8").strip().startswith("{")
