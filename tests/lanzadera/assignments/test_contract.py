# HARNESS-PROVENANCE: deterministic quality-harness v1.4 + lanzadera-mvp #631
"""Contract tests for the assignments module (DA-12, H11).

TDD: RED → GREEN.
  RED: no test file exists yet (0 collected).
  GREEN: tests pass against assignments.json.

Scope: verifies 622 active assignment rows. The SinAcceso exclusive
rule was applied when generating the fixture (DA-12 cortocircuito)."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

FIXTURE = (
    Path(__file__).parent.parent.parent.parent
    / "data"
    / "fixtures"
    / "lanzadera"
    / "assignments.json"
)
EXPECTED_ACTIVE = 622


class TestAssignmentsFixture:
    """622 active rows from assignments.json are well-formed."""

    @pytest.fixture
    def rows(self) -> list[dict]:
        return json.loads(FIXTURE.read_text())

    def test_count_active(self, rows: list[dict]) -> None:
        active = [r for r in rows if r.get("revoked_at") is None]
        assert len(active) == EXPECTED_ACTIVE, (
            f"expected {EXPECTED_ACTIVE} active, got {len(active)}"
        )

    def test_no_revoked_rows(self, rows: list[dict]) -> None:
        revoked = [r for r in rows if r.get("revoked_at") is not None]
        assert len(revoked) == 0, f"{len(revoked)} revoked rows — fixture should be all-active"

    def test_granted_by_null(self, rows: list[dict]) -> None:
        non_null = [r for r in rows if r.get("granted_by") is not None]
        assert len(non_null) == 0, f"{len(non_null)} rows with granted_by set (migration seeds)"

    def test_ids_unique(self, rows: list[dict]) -> None:
        ids = [r["id"] for r in rows]
        assert len(ids) == len(set(ids)), "duplicate assignment UUIDs"

    def test_app_ids_in_range(self, rows: list[dict]) -> None:
        for row in rows:
            assert row["app_id"] in range(1, 21), (
                f"assignment {row['id']}: app_id {row['app_id']} out of 1-20 range"
            )

    def test_profile_ids_present(self, rows: list[dict]) -> None:
        null_profile = [r for r in rows if not r.get("profile_id")]
        assert len(null_profile) == 0, f"{len(null_profile)} assignments with NULL profile_id"
