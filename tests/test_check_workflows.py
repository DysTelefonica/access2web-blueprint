# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 — assets/tests/test_check_workflows.py
"""Tests for `scripts/check_workflows.py` (issue #135).

A gate that cannot fail is not a gate — that is the whole reason this test
file exists. Every check has both a positive case (a workflow that passes)
and a negative case (the same workflow with the violation injected, asserting
the gate reports it with file and line). If any of the negative cases pass,
the gate is silently broken and the next workflow slips through.
"""

from __future__ import annotations

import sys
from pathlib import Path

# Make `scripts/` importable without sys.path games: pytest adds the repo
# root, the script is at `scripts/check_workflows.py` from there.
SCRIPTS_DIR = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

import check_workflows  # noqa: E402  — sys.path tweak above

# --------------------------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------------------------


def _write(tmp_path: Path, name: str, body: str) -> Path:
    d = tmp_path / "wf"
    d.mkdir()
    p = d / name
    p.write_text(body, encoding="utf-8")
    return p


def _report(check_module, root: Path) -> tuple[int, list[str]]:
    rc = check_module.main(["--root", str(root)])
    return rc, [line for line in (check_module.__dict__.get("_STDERR_LINES") or [])]


def _capture(check_module, root: Path) -> tuple[int, str]:
    """Run the gate and capture stderr (where violations are printed).

    The gate writes violations to stderr and the OK line to stdout. The
    return is (rc, combined_stderr).
    """
    import contextlib
    import io

    err = io.StringIO()
    out = io.StringIO()
    with contextlib.redirect_stderr(err), contextlib.redirect_stdout(out):
        rc = check_module.main(["--root", str(root)])
    return rc, err.getvalue()


# --------------------------------------------------------------------------------------------
# Check 1: duplicate YAML keys
# --------------------------------------------------------------------------------------------


def test_duplicate_keys_positive(tmp_path: Path) -> None:
    body = """\
name: ok
on: [push]
permissions: { contents: read }
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    steps:
      - uses: actions/checkout@3d3c42e5aac5ba805825da76410c181273ba90b1
        with:
          fetch-depth: 0
"""
    p = _write(tmp_path, "ok.yml", body)
    rc, stderr = _capture(check_workflows, p.parent)
    assert rc == 0, stderr
    assert "duplicate-key" not in stderr


def test_duplicate_keys_negative(tmp_path: Path) -> None:
    body = """\
name: dup
on: [push]
permissions: { contents: read }
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 5
    timeout-minutes: 10
    steps:
      - uses: actions/checkout@3d3c42e5aac5ba805825da76410c181273ba90b1
"""
    p = _write(tmp_path, "dup.yml", body)
    rc, stderr = _capture(check_workflows, p.parent)
    assert rc == 1
    assert "duplicate-key" in stderr
    # Second `timeout-minutes` is line 8 in this fixture (1-based).
    assert ":8: duplicate-key" in stderr


def test_duplicate_keys_skips_remaining_checks(tmp_path: Path) -> None:
    """If a file has duplicates, the other checks must not run on it.

    PyYAML's safe_load would silently pick one of the colliding values; any
    finding from a later check would then be arbitrary. The gate short-
    circuits to prevent that.
    """
    body = """\
name: dup-and-bad
on: [push]
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 5
    timeout-minutes: 10
    steps:
      - uses: actions/checkout@v3
"""
    p = _write(tmp_path, "both.yml", body)
    rc, stderr = _capture(check_workflows, p.parent)
    assert rc == 1
    assert "duplicate-key" in stderr
    # The unpinned `uses: actions/checkout@v3` is a real check-4 violation,
    # but the gate must NOT report it because safe_load would pick one of
    # the colliding `timeout-minutes` values arbitrarily.
    assert "uses-not-pinned" not in stderr


# --------------------------------------------------------------------------------------------
# Check 2: every job declares timeout-minutes
# --------------------------------------------------------------------------------------------


def test_timeout_minutes_positive(tmp_path: Path) -> None:
    body = """\
name: ok
on: [push]
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    steps:
      - run: echo ok
"""
    p = _write(tmp_path, "ok.yml", body)
    rc, stderr = _capture(check_workflows, p.parent)
    assert rc == 0, stderr
    assert "timeout-minutes" not in stderr


def test_timeout_minutes_negative(tmp_path: Path) -> None:
    body = """\
name: bad
on: [push]
jobs:
  build:
    runs-on: ubuntu-24.04
    steps:
      - run: echo ok
"""
    p = _write(tmp_path, "bad.yml", body)
    rc, stderr = _capture(check_workflows, p.parent)
    assert rc == 1
    assert "timeout-minutes" in stderr
    assert "jobs.build" in stderr


# --------------------------------------------------------------------------------------------
# Check 3: no service container fixes a host port
# --------------------------------------------------------------------------------------------


def test_service_ports_positive(tmp_path: Path) -> None:
    """A bare container-port publish is accepted (publishes on random host port)."""
    body = """\
name: ok
on: [push]
jobs:
  test:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    services:
      postgres:
        image: postgres:16-alpine
        ports:
          - 5432
    steps:
      - run: echo ok
"""
    p = _write(tmp_path, "ok.yml", body)
    rc, stderr = _capture(check_workflows, p.parent)
    assert rc == 0, stderr
    assert "host-port-fix" not in stderr


def test_service_ports_negative_host_container(tmp_path: Path) -> None:
    body = """\
name: bad
on: [push]
jobs:
  test:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    services:
      postgres:
        image: postgres:16-alpine
        ports:
          - 5432:5432
    steps:
      - run: echo ok
"""
    p = _write(tmp_path, "bad.yml", body)
    rc, stderr = _capture(check_workflows, p.parent)
    assert rc == 1
    assert "host-port-fix" in stderr
    assert "postgres" in stderr
    assert "5432:5432" in stderr


# --------------------------------------------------------------------------------------------
# Check 4: every uses: pinned by SHA
# --------------------------------------------------------------------------------------------


def test_uses_pinned_positive(tmp_path: Path) -> None:
    body = """\
name: ok
on: [push]
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    steps:
      - uses: actions/checkout@3d3c42e5aac5ba805825da76410c181273ba90b1
      - uses: actions/setup-python@5fda3b95a4ea91299a34e894583c3862153e4b97
"""
    p = _write(tmp_path, "ok.yml", body)
    rc, stderr = _capture(check_workflows, p.parent)
    assert rc == 0, stderr
    assert "uses-not-pinned" not in stderr


def test_uses_pinned_negative_tag(tmp_path: Path) -> None:
    body = """\
name: bad
on: [push]
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    steps:
      - uses: actions/checkout@v3
"""
    p = _write(tmp_path, "bad.yml", body)
    rc, stderr = _capture(check_workflows, p.parent)
    assert rc == 1
    assert "uses-not-pinned" in stderr
    assert "actions/checkout@v3" in stderr


def test_uses_pinned_negative_branch(tmp_path: Path) -> None:
    body = """\
name: bad
on: [push]
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    steps:
      - uses: actions/setup-python@main
"""
    p = _write(tmp_path, "bad.yml", body)
    rc, stderr = _capture(check_workflows, p.parent)
    assert rc == 1
    assert "uses-not-pinned" in stderr
    assert "actions/setup-python@main" in stderr


# --------------------------------------------------------------------------------------------
# Integration: the gate runs against the real workflows in this repo
# --------------------------------------------------------------------------------------------


def test_gate_passes_on_repo_workflows() -> None:
    """If this fails, either a check regressed or the repo violates a rule
    it should now enforce. Fix the workflow, not the gate.
    """
    repo_root = Path(__file__).resolve().parent.parent
    rc, stderr = _capture(check_workflows, repo_root / ".github" / "workflows")
    assert rc == 0, stderr
