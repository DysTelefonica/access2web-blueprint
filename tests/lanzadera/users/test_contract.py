# HARNESS-PROVENANCE: deterministic quality-harness v1.4 + lanzadera-mvp #631
"""Contract tests for the users module (DA-3, D89).

TDD: RED → GREEN.
  RED: no test file exists yet (0 collected).
  GREEN: tests pass against users.json.

Scope: verifies 156 user rows are well-formed. Password hash is not
present in the fixture (DA-3: NULL until reset token consumed)."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

FIXTURE = (
    Path(__file__).parent.parent.parent.parent / "data" / "fixtures" / "lanzadera" / "users.json"
)
EXPECTED_COUNT = 156


class TestUsersFixture:
    """156 rows from users.json are well-formed."""

    @pytest.fixture
    def rows(self) -> list[dict]:
        return json.loads(FIXTURE.read_text())

    def test_count(self, rows: list[dict]) -> None:
        assert len(rows) == EXPECTED_COUNT, f"expected {EXPECTED_COUNT}, got {len(rows)}"

    def test_ids_unique(self, rows: list[dict]) -> None:
        ids = [r["id"] for r in rows]
        assert len(ids) == len(set(ids)), "duplicate user UUIDs"

    def test_emails_unique(self, rows: list[dict]) -> None:
        emails = [r["email"] for r in rows]
        assert len(emails) == len(set(emails)), "duplicate emails"

    def test_all_password_reset_required(self, rows: list[dict]) -> None:
        wrong = [r for r in rows if r["status"] != "password_reset_required"]
        assert len(wrong) == 0, f"{len(wrong)} users not in password_reset_required (DA-3, D89)"

    def test_failed_attempts_zero(self, rows: list[dict]) -> None:
        non_zero = [r for r in rows if r.get("failed_attempts", 0) != 0]
        assert len(non_zero) == 0, f"{len(non_zero)} users with failed_attempts != 0"

    def test_dni_present(self, rows: list[dict]) -> None:
        null_dni = [r for r in rows if not r.get("dni")]
        assert len(null_dni) == 0, f"{len(null_dni)} users with missing dni"

    def test_email_synthetic_domain(self, rows: list[dict]) -> None:
        for row in rows:
            assert "@synthetic.lanzadera" in row["email"], (
                f"user {row['id']}: email must be synthetic, got {row['email']}"
            )
