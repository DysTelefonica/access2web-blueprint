# HARNESS-PROVENANCE: deterministic quality-harness v1.4 + lanzadera-mvp #631
"""Contract tests for the profiles module (DA-12, D22, D45, D46, D110).

TDD: RED → GREEN.
  RED: no test file exists yet (0 collected).
  GREEN: tests pass against profiles.json.

Scope: verifies 160 profile rows (1 DEFAULT + 7 legacy × 20 apps) are
well-formed. The SinAcceso exclusive rule is tested in
test_legacy_role_map.py (DA-12)."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

FIXTURE = (
    Path(__file__).parent.parent.parent.parent / "data" / "fixtures" / "lanzadera" / "profiles.json"
)

# Expected: 1 DEFAULT + 7 legacy per app × 20 apps.
EXPECTED_DEFAULT = 20
EXPECTED_LEGACY_CODES = 7
EXPECTED_TOTAL = EXPECTED_DEFAULT + EXPECTED_LEGACY_CODES * 20  # 160


class TestProfilesFixture:
    """160 rows from profiles.json are well-formed."""

    @pytest.fixture
    def rows(self) -> list[dict]:
        return json.loads(FIXTURE.read_text())

    def test_count(self, rows: list[dict]) -> None:
        assert len(rows) == EXPECTED_TOTAL, f"expected {EXPECTED_TOTAL}, got {len(rows)}"

    def test_default_profiles_per_app(self, rows: list[dict]) -> None:
        default_rows = [r for r in rows if r["code"] == "DEFAULT"]
        assert len(default_rows) == EXPECTED_DEFAULT, (
            f"expected {EXPECTED_DEFAULT} DEFAULT profiles, got {len(default_rows)}"
        )

    def test_legacy_profiles_per_app(self, rows: list[dict]) -> None:
        legacy_rows = [r for r in rows if r["code"] != "DEFAULT"]
        expected = EXPECTED_LEGACY_CODES * 20  # 7 × 20
        assert len(legacy_rows) == expected, (
            f"expected {expected} legacy profiles, got {len(legacy_rows)}"
        )

    def test_capabilities_empty(self, rows: list[dict]) -> None:
        # G-2 ABIERTO: all seeded profiles carry {} until ABIERTO resolved.
        non_empty = [r for r in rows if r.get("capabilities") != {}]
        assert len(non_empty) == 0, f"{len(non_empty)} profiles have non-empty capabilities (G-2)"

    def test_all_active(self, rows: list[dict]) -> None:
        inactive = [r for r in rows if not r.get("active")]
        assert len(inactive) == 0, f"{len(inactive)} inactive profiles"

    def test_ids_unique(self, rows: list[dict]) -> None:
        ids = [r["id"] for r in rows]
        assert len(ids) == len(set(ids)), "duplicate profile UUIDs"

    def test_all_have_app_id(self, rows: list[dict]) -> None:
        for row in rows:
            assert row["app_id"] in range(1, 21), (
                f"profile {row['id']}: app_id {row['app_id']} out of 1-20 range"
            )
