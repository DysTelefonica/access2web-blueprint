# HARNESS-PROVENANCE: deterministic quality-harness v1.4 + lanzadera-mvp #631
"""Contract tests for the apps module — data loaded from fixture (DA-7, D52).

TDD: RED → GREEN.
  RED: no test file exists yet (0 collected).
  GREEN: tests pass against apps.json fixture.

Scope: verifies the app entities are well-formed from the fixture data
without reaching Postgres. Uses FakeFixtures + fixture JSON directly."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

FIXTURE = (
    Path(__file__).parent.parent.parent.parent
    / "data" / "fixtures" / "lanzadera" / "apps.json"
)

# Expected by the migration.
EXPECTED_COUNT = 20


class TestAppsFixture:
    """All 20 rows from apps.json load without error."""

    @pytest.fixture
    def rows(self) -> list[dict]:
        return json.loads(FIXTURE.read_text())

    def test_count(self, rows: list[dict]) -> None:
        assert len(rows) == EXPECTED_COUNT, f"expected {EXPECTED_COUNT}, got {len(rows)}"

    def test_ids_unique(self, rows: list[dict]) -> None:
        ids = [r["id"] for r in rows]
        assert len(ids) == len(set(ids)), "duplicate app IDs"

    def test_short_codes_unique(self, rows: list[dict]) -> None:
        codes = [r["short_code"] for r in rows]
        assert len(codes) == len(set(codes)), "duplicate short_codes"

    def test_topology_values(self, rows: list[dict]) -> None:
        allowed = {"central", "office-nas"}
        for row in rows:
            assert row["deployment_topology"] in allowed, (
                f"app {row['short_code']}: unknown topology "
                f"{row['deployment_topology']!r}"
            )

    def test_requires_office_presence_boolean(self, rows: list[dict]) -> None:
        for row in rows:
            assert isinstance(row["requires_office_presence"], bool), (
                f"app {row['short_code']}: "
                f"requires_office_presence must be bool, got {type(row['requires_office_presence'])}"
            )

    def test_no_null_required_fields(self, rows: list[dict]) -> None:
        for row in rows:
            assert row["id"] is not None
            assert row["name"]
            assert row["short_code"]
            assert row["deployment_topology"]
