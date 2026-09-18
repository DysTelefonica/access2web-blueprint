"""Tests for M02 (#258) + M03 (#259) extractor tranches.

Verifies that the ``GROUPS_16_30`` and ``GROUPS_31_49`` constants in
``app/src/modules/expedientes/migration/extractor.py`` match the source
plan task table (M02: tables 16-30, M03: tables 31-49) and that the
extractor accepts them as a ``tables=`` subset without affecting the
M01 tranche (#257) behaviour.
"""

from __future__ import annotations

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT / "tests" / "lanzadera" / "exp"))


def test_tranche_constants_match_source_plan() -> None:
    """GROUPS_01_15 + GROUPS_16_30 + GROUPS_31_49 partition the 49-table dataset."""
    from app.src.modules.expedientes.migration.extractor import (
        GROUPS_01_15,
        GROUPS_16_30,
        GROUPS_31_49,
    )

    assert len(GROUPS_01_15) == 15, f"M01 tranche size: got {len(GROUPS_01_15)}"
    assert len(GROUPS_16_30) == 15, f"M02 tranche size: got {len(GROUPS_16_30)}"
    assert len(GROUPS_31_49) == 19, f"M03 tranche size: got {len(GROUPS_31_49)}"

    seen: set[str] = set()
    for tranche in (GROUPS_01_15, GROUPS_16_30, GROUPS_31_49):
        for table in tranche:
            assert table not in seen, f"duplicate table across tranches: {table}"
            seen.add(table)
    assert len(seen) == 49, f"expected 49 unique tables, got {len(seen)}"


def test_tranche_subset_extraction_writes_only_targeted_tables() -> None:
    """Calling run_once(tables=GROUPS_16_30) only writes rows from those 15 tables."""
    from app.src.modules.expedientes.migration.extractor import (
        GROUPS_16_30,
        Extractor,
    )
    from tests.lanzadera.exp.test_extractor import (
        FakeBackend,
        InMemoryStagingWriter,
        InMemoryWatermarkStore,
    )

    backend = FakeBackend({table: [{"ID": 1, "value": 1}] for table in GROUPS_16_30})
    watermark = InMemoryWatermarkStore()
    writer = InMemoryStagingWriter()

    result = Extractor(backend, watermark, writer).run_once(tables=GROUPS_16_30)

    assert result.rows_extracted == 15
    assert set(w.table for w in writer.writes) == set(GROUPS_16_30)
    assert all(table in GROUPS_16_30 for table in result.tables_attempted)


def test_tranche_subsequent_run_is_idempotent() -> None:
    """Running a tranche twice produces zero rows on the second run (D-EXP-9)."""
    from app.src.modules.expedientes.migration.extractor import (
        GROUPS_31_49,
        Extractor,
    )
    from tests.lanzadera.exp.test_extractor import (
        FakeBackend,
        InMemoryStagingWriter,
        InMemoryWatermarkStore,
    )

    backend = FakeBackend({"TbOrganosCompetentes": [{"ID": 1, "v": "x"}, {"ID": 2, "v": "y"}]})
    watermark = InMemoryWatermarkStore()
    writer = InMemoryStagingWriter()

    first = Extractor(backend, watermark, writer).run_once(tables=GROUPS_31_49)
    second = Extractor(backend, watermark, writer).run_once(tables=GROUPS_31_49)

    assert first.rows_extracted == 2
    assert second.rows_extracted == 0
    assert second.rows_skipped_by_watermark == 2
