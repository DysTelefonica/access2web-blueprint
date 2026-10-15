# HARNESS-PROVENANCE: access2web-blueprint issue #875 — test_check_deadcode.py
"""Wiring pin for `scripts/check_deadcode.py` (issue #875): one test per rule.

Four rules from the issue:

1. A new unreachable symbol outside the baseline → exit 1.
2. The baseline, unchanged, passes (exit 0).
3. A stale entry in the baseline → exit 1 (ratchet only goes down).
4. `vulture` not installed → exit 1, fail closed.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "check_deadcode.py"

# Unreachable code after `return` is the one thing vulture flags at 100%
# confidence in a tiny fixture (small fixtures don't trip the 80% threshold for
# unused locals — those need real codebase context).
DEAD_VARIABLE_SRC = (
    "def used():\n"
    "    return 0\n"
    "    return 1\n"  # unreachable → 100% confidence
    "used()\n"
)

CLEAN_SRC = "def used():\n    return 0\nused()\n"


def _vulture_findings(app_dir: Path) -> str:
    """Run vulture directly to capture the canonical finding text."""
    vulture = shutil.which("vulture")
    assert vulture is not None, "vulture must be installed to run this test"
    return subprocess.run(
        [vulture, str(app_dir), "--min-confidence", "80"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        check=False,
    ).stdout


def _make_app(tmp_path: Path, src: str) -> Path:
    app = tmp_path / "app"
    app.mkdir()
    (app / "m.py").write_text(src, encoding="utf-8")
    return app


def _run_check(app_dir: Path, baseline: Path, *, env=None) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(SCRIPT), str(app_dir), str(baseline)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=env,
        check=False,
    )


def test_rule_1_new_unreachable_symbol_fails(tmp_path: Path) -> None:
    """A finding not in the baseline is a failure (rule 1)."""
    app = _make_app(tmp_path, DEAD_VARIABLE_SRC)
    baseline = tmp_path / "b.txt"
    baseline.write_text("", encoding="utf-8")  # empty baseline
    result = _run_check(app, baseline)
    assert result.returncode == 1, result.stderr
    assert "new finding" in result.stderr


def test_rule_2_unchanged_baseline_passes(tmp_path: Path) -> None:
    """When vulture's output equals the baseline, exit 0 (rule 2)."""
    app = _make_app(tmp_path, DEAD_VARIABLE_SRC)
    baseline = tmp_path / "b.txt"
    baseline.write_text(_vulture_findings(app), encoding="utf-8")
    result = _run_check(app, baseline)
    assert result.returncode == 0, result.stderr


def test_rule_3_stale_baseline_entry_fails(tmp_path: Path) -> None:
    """An entry in the baseline that vulture no longer reports is a failure (rule 3)."""
    app = _make_app(tmp_path, CLEAN_SRC)  # no findings here
    baseline = tmp_path / "b.txt"
    baseline.write_text(
        f"{app}/m.py:99: unused variable 'ghost' (100% confidence)\n",
        encoding="utf-8",
    )
    result = _run_check(app, baseline)
    assert result.returncode == 1, result.stderr
    assert "stale" in result.stderr


def test_rule_4_missing_vulture_fails_closed(tmp_path: Path) -> None:
    """When vulture is not on PATH, the script fails closed (rule 4)."""
    app = _make_app(tmp_path, CLEAN_SRC)
    baseline = tmp_path / "b.txt"
    baseline.write_text("", encoding="utf-8")
    empty_path = {k: v for k, v in os.environ.items() if k != "PATH"}
    # Sanity: vulture is reachable from the current PATH so we know the test
    # would otherwise succeed. Then we strip PATH so the subprocess can't find it.
    assert shutil.which("vulture") is not None
    empty_path["PATH"] = "/dev/null"  # an empty directory is enough
    result = _run_check(app, baseline, env=empty_path)
    assert result.returncode == 1, result.stderr
    assert "vulture" in result.stderr.lower()
