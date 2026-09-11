# HARNESS-PROVENANCE: deterministic quality-harness v1.4 + lanzadera-mvp #631
"""Contract tests for the audit module (DA-11).

TDD: RED → GREEN.
  RED: no test file exists yet (0 collected).
  GREEN: tests pass against audit_events.json.

Scope: verifies 200 audit events. DA-11 retires SSID, BSSID, GPS,
machine_name, ip_address columns from the legacy schema — those fields
are absent from the fixture (no telemetry data)."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

FIXTURE = (
    Path(__file__).parent.parent.parent.parent
    / "data"
    / "fixtures"
    / "lanzadera"
    / "audit_events.json"
)
EXPECTED_COUNT = 200
ALLOWED_TYPES = {"auth.login.success", "auth.login.failure", "app.open"}
ALLOWED_RESULTS = {"success", "failure"}


class TestAuditFixture:
    """200 rows from audit_events.json are well-formed."""

    @pytest.fixture
    def rows(self) -> list[dict]:
        return json.loads(FIXTURE.read_text())

    def test_count(self, rows: list[dict]) -> None:
        assert len(rows) == EXPECTED_COUNT, f"expected {EXPECTED_COUNT}, got {len(rows)}"

    def test_event_types_allowed(self, rows: list[dict]) -> None:
        for row in rows:
            assert row["event_type"] in ALLOWED_TYPES, (
                f"event {row['id']}: unknown type {row['event_type']!r}"
            )

    def test_results_allowed(self, rows: list[dict]) -> None:
        for row in rows:
            assert row["result"] in ALLOWED_RESULTS, (
                f"event {row['id']}: unknown result {row['result']!r}"
            )

    def test_module_lanzadera(self, rows: list[dict]) -> None:
        non_lanzadera = [r for r in rows if r.get("module") != "lanzadera"]
        assert len(non_lanzadera) == 0, f"{len(non_lanzadera)} events with non-lanzadera module"

    def test_ids_unique(self, rows: list[dict]) -> None:
        ids = [r["id"] for r in rows]
        assert len(ids) == len(set(ids)), "duplicate event UUIDs"

    def test_no_telemetry_columns(self, rows: list[dict]) -> None:
        # DA-11: these legacy columns must not appear.
        forbidden = {"ssid", "bssid", "gps_lat", "gps_lon", "machine_name", "ip_address"}
        for row in rows:
            payload = row.get("payload", {})
            overlap = set(payload.keys()) & forbidden
            assert len(overlap) == 0, f"event {row['id']}: telemetry fields in payload: {overlap}"
