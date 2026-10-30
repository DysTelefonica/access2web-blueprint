"""The docs-only lane classifier (issue #1368).

A pull request whose whole diff lives under ``docs/`` does not need the heavy
suite: ``quality``, ``integration``, ``security``, ``codeql`` and ``e2e-smoke``
skip it and the ``required`` aggregator accepts those skips as OK
(``.agents/skills/ci-pattern/assets/required-jobs/README.md``, ``docs_only_lane``).

The classifier is fail-closed. It answers ``true`` only for a NON-EMPTY changed
set whose every path is documentation: one single path outside ``docs/`` pays the
whole heavy suite, and an empty diff is a known no-op, not a documentation diff.
The anti-self-exemption toll is implicit — the classifier, the workflow and the
policy live outside ``docs/``, so editing them never exempts anything.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "scripts" / "check_docs_only_lane.py"


def _classify(*paths: str) -> str:
    """Run the classifier over ``paths`` (one per line) and return its verdict."""
    result = subprocess.run(
        [sys.executable, str(SCRIPT)],
        input="".join(f"{path}\n" for path in paths),
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode == 0, result.stdout + result.stderr
    return result.stdout.strip()


def test_a_diff_entirely_under_docs_is_a_docs_only_diff() -> None:
    assert (
        _classify("docs/codebase/testing-and-ci.md", "docs/03-aplicaciones/lanzadera/spec.md")
        == "true"
    )


def test_one_path_outside_docs_pays_the_whole_heavy_suite() -> None:
    assert _classify("docs/codebase/testing-and-ci.md", "app/src/modules/lanzadera/x.py") == "false"


def test_an_empty_diff_is_not_a_docs_only_diff() -> None:
    assert _classify() == "false"


def test_a_root_markdown_file_is_not_under_docs() -> None:
    assert _classify("README.md") == "false"


def test_the_lane_never_exempts_itself() -> None:
    """The classifier, the workflow and the policy are not documentation."""
    assert _classify(".github/workflows/pr-gate.yml") == "false"
    assert _classify(".github/required-jobs-policy.json") == "false"
    assert _classify("scripts/check_docs_only_lane.py") == "false"


def test_a_path_that_only_looks_like_docs_is_not_documentation() -> None:
    """``docs`` as a sibling directory (``docs-archive/``) is not ``docs/``."""
    assert _classify("docs-archive/old.md") == "false"
    assert _classify("app/docs/notes.md") == "false"
