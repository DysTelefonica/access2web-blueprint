"""Watermark store — idempotency boundary for the extraction loop.

The watermark is a JSON file that records, per table, the highest
primary-key value the extractor has already copied to staging. The
extractor's main loop reads ``SELECT * WHERE PK > watermark(pk)``;
re-running the extractor is therefore a no-op once the staging table
matches the source.

D-EXP-9 (design.md) names the watermark as the recovery boundary
for partial extractions: a crash mid-run loses the rows above the
last-committed watermark, and the next run picks up from there
without re-processing. The store is intentionally tiny (one
``{table_name: last_pk}`` entry per table); the file format is JSON
so operators can read it with ``cat`` and the test suite can assert
its contents without parsing anything custom.
"""

from __future__ import annotations

import json
import tempfile
from pathlib import Path
from typing import Protocol


class WatermarkStore(Protocol):
    """Read / write the ``{table_name: last_pk}`` map.

    Two implementations ship in M01:
    - :class:`JsonFileWatermarkStore` — the production store, persists
      to a JSON file in the staging volume.
    - :class:`InMemoryWatermarkStore` — the test fake, lives in
      process memory and disappears at teardown.
    """

    def get(self, table: str) -> int | None:
        """Return the last PK stored for ``table`` or ``None`` on first run."""
        ...

    def put(self, table: str, last_pk: int) -> None:
        """Persist ``last_pk`` as the new high-water mark for ``table``."""
        ...

    def snapshot(self) -> dict[str, int]:
        """Return a copy of the full map (test diagnostics, ops script)."""
        ...


class InMemoryWatermarkStore:
    """In-process watermark for tests."""

    def __init__(self) -> None:
        self._store: dict[str, int] = {}

    def get(self, table: str) -> int | None:
        return self._store.get(table)

    def put(self, table: str, last_pk: int) -> None:
        self._store[table] = last_pk

    def snapshot(self) -> dict[str, int]:
        return dict(self._store)


class JsonFileWatermarkStore:
    """JSON file watermark — the production implementation.

    The file is written atomically (write to a temp file in the same
    directory, ``os.replace`` into place) so a crash mid-write does
    not leave a partial JSON on disk that the next run cannot parse.
    """

    def __init__(self, path: Path) -> None:
        self._path = path
        self._cache: dict[str, int] = self._load()

    def _load(self) -> dict[str, int]:
        if not self._path.exists():
            return {}
        try:
            data = json.loads(self._path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            # A corrupt file is a loud failure, not a silent default —
            # the operator must restore from backup rather than the
            # extractor guessing the safe state.
            raise
        if not isinstance(data, dict):
            raise ValueError(
                f"watermark file {self._path} is malformed: "
                f"expected an object, got {type(data).__name__}"
            )
        # Coerce keys to str and values to int; a hand-edited file
        # with the wrong types is a hard error.
        return {str(k): int(v) for k, v in data.items()}

    def _flush(self) -> None:
        self._path.parent.mkdir(parents=True, exist_ok=True)
        # Atomic write: temp file in the same directory, then replace.
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            dir=self._path.parent,
            delete=False,
            prefix=self._path.name + ".",
            suffix=".tmp",
        ) as tmp:
            json.dump(self._cache, tmp, sort_keys=True, indent=2)
            tmp_path = Path(tmp.name)
        tmp_path.replace(self._path)

    def get(self, table: str) -> int | None:
        return self._cache.get(table)

    def put(self, table: str, last_pk: int) -> None:
        self._cache[table] = last_pk
        self._flush()

    def snapshot(self) -> dict[str, int]:
        return dict(self._cache)
