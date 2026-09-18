"""Smoke tests for ``scripts/migrate_from_access.py`` (WU E2, issue #52).

Verifies:
  - a clean CSV (all hashes are valid 64-char hex SHA256) exits 0;
  - a CSV with a malformed hash exits 2 and lists each offender by row;
  - --expect-sha256-hash enforces equality and exits 2 on mismatch;
  - missing input file exits 1.
"""

from __future__ import annotations

import csv
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
SCRIPT = REPO_ROOT / "scripts" / "migrate_from_access.py"
GOOD_FIXTURE = REPO_ROOT / "tests" / "fixtures" / "lanzadera" / "access_users.csv"


def _write_csv(path: Path, header: list[str], rows: list[list[str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(header)
        writer.writerows(rows)


def _run(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, str(SCRIPT), *args], capture_output=True, text=True, check=False
    )


def test_clean_csv_exits_zero(tmp_path: Path) -> None:
    fixture = tmp_path / "good.csv"
    digest = "5e884898da28047151d0e56f8dc6292773603d0d6aabbdd62a11ef721d1542d8"
    _write_csv(
        fixture,
        ["email", "password_hash", "legacy_id"],
        [["alice@example.com", digest, "1"], ["bob@example.com", digest, "2"]],
    )

    result = _run("--input", str(fixture))

    assert result.returncode == 0
    assert "ok\talice@example.com\tok" in result.stdout
    assert "ok\tbob@example.com\tok" in result.stdout


def test_malformed_hash_exits_two(tmp_path: Path) -> None:
    fixture = tmp_path / "mixed.csv"
    _write_csv(
        fixture,
        ["email", "password_hash", "legacy_id"],
        [
            ["good@example.com", "5" * 64, "1"],
            ["bad@example.com", "NOT_HEX", "2"],
            ["short@example.com", "abc", "3"],
            ["empty@example.com", "", "4"],
        ],
    )

    result = _run("--input", str(fixture))

    assert result.returncode == 2
    assert "ok\tgood@example.com\tok" in result.stdout
    assert "rejected\tbad@example.com\thash length" in result.stdout
    assert "rejected\tshort@example.com\thash length" in result.stdout
    assert "rejected\tempty@example.com\tempty hash" in result.stdout
    assert "summary: total=4 malformed=3" in result.stderr


def test_expect_hash_mismatch_exits_two(tmp_path: Path) -> None:
    fixture = tmp_path / "mismatch.csv"
    digest_a = "a" * 64
    digest_b = "b" * 64
    _write_csv(
        fixture,
        ["email", "password_hash", "legacy_id"],
        [["alice@example.com", digest_a, "1"], ["bob@example.com", digest_b, "2"]],
    )

    result = _run("--input", str(fixture), "--expect-sha256-hash", digest_a)

    assert result.returncode == 2
    assert "ok\talice@example.com\tok" in result.stdout
    assert "rejected\tbob@example.com\thash does not match" in result.stdout


def test_missing_input_exits_one(tmp_path: Path) -> None:
    result = _run("--input", str(tmp_path / "nope.csv"))

    assert result.returncode == 1
    assert "input file not found" in result.stderr


def test_default_fixture_in_repo_runs_clean() -> None:
    """Run against the in-repo fixture (good + 3 bad rows) to confirm docstrings match reality."""
    if not GOOD_FIXTURE.exists():
        return

    result = _run("--input", str(GOOD_FIXTURE))

    assert result.returncode == 2
    assert "summary: total=6 malformed=3" in result.stderr
