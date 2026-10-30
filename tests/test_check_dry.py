"""BP's copy of ``check_dry`` reads its subjects and ceilings from ``quality-policy.json``.

The DRY gate is the canon's asset (``deterministic-quality-harness`` v2.2) with BP's ``BASELINE``
of tolerated clone groups (data, not mechanism). These tests pin the canon wiring that was adopted
— the explicit policy date, the envelope bound to the policy, the unparseable-file accounting —
and the regression that the gate still reports a duplicated block at that date.
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
FIXTURES = Path(__file__).resolve().parent / "fixtures"
#: DQH 2.2: the BASELINE is evaluated at an explicit date, never at the wall clock.
POLICY_DATE = "2026-10-21"


def _run(root: Path, *extra: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(SCRIPTS / "check_dry.py"), "--root", str(root), *extra],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )


def test_the_policy_date_is_required() -> None:
    result = _run(FIXTURES / "dry_clean")
    assert result.returncode == 2, result.stdout + result.stderr
    assert "--policy-date" in result.stderr


def test_envelope_is_bound_to_the_policy_and_its_date() -> None:
    result = _run(FIXTURES / "dry_clean", "--policy-date", POLICY_DATE, "--json")
    envelope = json.loads(result.stdout)
    assert result.returncode == 0, result.stdout + result.stderr
    assert envelope["policy_date"] == POLICY_DATE
    assert envelope["ceilings"] == {"duplicate_groups": 0}
    assert envelope["subjects"]["checked"], "the clean fixture measures at least one subject"
    assert envelope["subjects"]["unclassified"] == []
    assert envelope["policy_identity"].startswith("sha256:")
    assert envelope["indicators"]["statements_measured"] > 0


def test_an_unparseable_file_is_reported_as_unclassified(tmp_path: Path) -> None:
    """A file the gate cannot parse is not silently dropped from the accounting.

    Skipping it would make the subject accounting disagree with the policy manifest, which is
    exactly what ``bind_envelope`` rejects; naming it is what keeps the report honest.
    """
    package = tmp_path / "app" / "pkg"
    package.mkdir(parents=True)
    (package / "broken.py").write_text("def f(:\n", encoding="utf-8")
    (package / "fine.py").write_text("def f(x: int) -> int:\n    return x + 1\n", encoding="utf-8")
    result = _run(tmp_path, "--policy-date", POLICY_DATE, "--json")
    envelope = json.loads(result.stdout)
    assert envelope["subjects"]["unclassified"] == ["app/pkg/broken.py"]
    assert envelope["subjects"]["checked"] == ["app/pkg/fine.py"]


def test_a_duplicated_block_still_fails_at_the_given_date() -> None:
    """Regression: the canon wiring must not soften the verdict the gate exists for."""
    result = _run(FIXTURES / "dry_violation", "--policy-date", POLICY_DATE)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "duplicated block" in result.stdout


def test_the_header_declares_the_canon_and_its_baseline_as_data() -> None:
    header = (SCRIPTS / "check_dry.py").read_text(encoding="utf-8")[:2500]
    assert "HARNESS-PROVENANCE: deterministic-quality-harness v2.2" in header
    assert "BASELINE" in header, "the header must say that the tolerated clone groups are BP data"
