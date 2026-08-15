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

import pytest

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
concurrency:
  group: ${{ github.workflow }}-${{ github.ref }}
  cancel-in-progress: false
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
concurrency:
  group: ${{ github.workflow }}-${{ github.ref }}
  cancel-in-progress: false
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
concurrency:
  group: ${{ github.workflow }}-${{ github.ref }}
  cancel-in-progress: false
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
concurrency:
  group: ${{ github.workflow }}-${{ github.ref }}
  cancel-in-progress: false
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
concurrency:
  group: ${{ github.workflow }}-${{ github.ref }}
  cancel-in-progress: false
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
concurrency:
  group: ${{ github.workflow }}-${{ github.ref }}
  cancel-in-progress: false
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
concurrency:
  group: ${{ github.workflow }}-${{ github.ref }}
  cancel-in-progress: false
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
concurrency:
  group: ${{ github.workflow }}-${{ github.ref }}
  cancel-in-progress: false
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
concurrency:
  group: ${{ github.workflow }}-${{ github.ref }}
  cancel-in-progress: false
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


# --------------------------------------------------------------------------------------------
# CHECK 5 — docker preflight (issue #141)
# --------------------------------------------------------------------------------------------
_COMPLIANT_HEADER = """\
name: ok
on: [push]
concurrency:
  group: ${{ github.workflow }}-${{ github.ref }}
  cancel-in-progress: false
jobs:
  scan:
    runs-on: ubuntu-24.04
    timeout-minutes: 15
    steps:
"""


def test_docker_preflight_positive_same_step(tmp_path: Path) -> None:
    """The guard may live earlier in the SAME step — the shape this repo uses."""
    body = (
        _COMPLIANT_HEADER
        + """\
      - name: scan
        run: |
          if ! timeout 30 docker info >/dev/null 2>&1; then exit 1; fi
          docker run --rm scanner
"""
    )
    rc, stderr = _capture(check_workflows, _write(tmp_path, "ok.yml", body).parent)
    assert rc == 0, stderr


def test_docker_preflight_positive_earlier_step(tmp_path: Path) -> None:
    body = (
        _COMPLIANT_HEADER
        + """\
      - name: preflight
        run: timeout 30 docker info >/dev/null
      - name: scan
        run: docker run --rm scanner
"""
    )
    rc, stderr = _capture(check_workflows, _write(tmp_path, "ok.yml", body).parent)
    assert rc == 0, stderr


def test_docker_preflight_negative_absent(tmp_path: Path) -> None:
    body = (
        _COMPLIANT_HEADER
        + """\
      - name: scan
        run: docker run --rm scanner
"""
    )
    rc, stderr = _capture(check_workflows, _write(tmp_path, "bad.yml", body).parent)
    assert rc == 1
    assert "docker-preflight" in stderr
    assert "no `docker info` check" in stderr


def test_docker_preflight_negative_unwrapped(tmp_path: Path) -> None:
    """A bare `docker info` hangs exactly when the daemon is wedged."""
    body = (
        _COMPLIANT_HEADER
        + """\
      - name: scan
        run: |
          docker info >/dev/null
          docker run --rm scanner
"""
    )
    rc, stderr = _capture(check_workflows, _write(tmp_path, "bad.yml", body).parent)
    assert rc == 1
    assert "not under `timeout`" in stderr


def test_docker_preflight_word_timeout_in_prose_does_not_count(tmp_path: Path) -> None:
    """`timeout-minutes` and the word "timeout" in a message are not a guard.

    This exact false negative shipped in the first draft of the check: the repo's
    own preflight message says "silent 15-minute timeout", so a substring test
    reported two unwrapped guards as wrapped.
    """
    body = (
        _COMPLIANT_HEADER
        + """\
      - name: scan
        run: |
          echo "a hung daemon turns this into a silent 15-minute timeout"
          docker info >/dev/null
          docker run --rm scanner
"""
    )
    rc, stderr = _capture(check_workflows, _write(tmp_path, "bad.yml", body).parent)
    assert rc == 1
    assert "not under `timeout`" in stderr


def test_docker_preflight_ignores_comment_lines(tmp_path: Path) -> None:
    """A comment naming the guard is not the guard.

    The mirror of the rule above: files document what they do, and a scanner that
    reads comments would accept a documented-but-absent guard.
    """
    body = (
        _COMPLIANT_HEADER
        + """\
      - name: scan
        run: |
          # we used to run: timeout 30 docker info
          docker run --rm scanner
"""
    )
    rc, stderr = _capture(check_workflows, _write(tmp_path, "bad.yml", body).parent)
    assert rc == 1
    assert "no `docker info` check" in stderr


# --------------------------------------------------------------------------------------------
# CHECK 6 — FIFO concurrency (issue #141, #152)
# --------------------------------------------------------------------------------------------
# Audit 2026-08-12 (umbrella #117) re-classified the cancel-in-progress
# policy: missing concurrency group is still an ERROR (the runner pool is
# shared and no group means no FIFO), but a non-cancel-safe
# `cancel-in-progress: true` is now a WARN, not an error. The repo's design
# is per-gate FIFO without cancelling; the WARN surfaces any future drift
# without breaking a workflow that already lives in that shape.
def test_concurrency_negative_missing(tmp_path: Path) -> None:
    body = """\
name: bad
on: [push]
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    steps:
      - run: echo ok
"""
    rc, stderr = _capture(check_workflows, _write(tmp_path, "bad.yml", body).parent)
    assert rc == 1
    assert "declares no concurrency group" in stderr


def test_concurrency_cancel_in_progress_warns_not_errors(tmp_path: Path) -> None:
    """`cancel-in-progress: true` is a WARN — it does not fail the gate.

    The runner-skill policy allows it when the job is in the repo's
    allowlist; an empty allowlist means every `true` warns. The gate stays
    green (rc=0) so the workflow that already declares `true` does not
    break on the first run after this check was added.
    """
    body = """\
name: warn
on: [push]
concurrency:
  group: ${{ github.workflow }}-${{ github.ref }}
  cancel-in-progress: true
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    steps:
      - run: echo ok
"""
    rc, stderr = _capture(check_workflows, _write(tmp_path, "warn.yml", body).parent)
    assert rc == 0, stderr
    assert "WARN" in stderr
    assert "cancel-in-progress: true" in stderr
    assert "build" in stderr


def test_concurrency_cancel_in_progress_allowlisted_silent(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """A job in `_CANCEL_SAFE_JOBS` may use `cancel-in-progress: true` silently.

    The allowlist is empty in production. Tests patch it so this branch
    is exercised without modifying the source. A regression that flips
    the allowlist check would either let production jobs warn silently
    (loss of signal) or block allowlisted jobs (false ERROR) — both are
    the same class of gate failure: a check that no longer reflects the
    policy it claims to enforce.
    """
    body = """\
name: ok
on: [push]
permissions:
  contents: read
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    concurrency:
      group: gate-build
      cancel-in-progress: true
    steps:
      - run: echo ok
"""
    monkeypatch.setattr(check_workflows, "_CANCEL_SAFE_JOBS", {"build"})
    rc, stderr = _capture(check_workflows, _write(tmp_path, "ok.yml", body).parent)
    assert rc == 0, stderr
    assert "WARN" not in stderr


def test_concurrency_positive_per_job_overrides_absent_top_level(
    tmp_path: Path,
) -> None:
    """This repo declares it per gate, not top-level; both must be accepted."""
    body = """\
name: ok
on: [push]
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    concurrency:
      group: gate-build
      cancel-in-progress: false
    steps:
      - run: echo ok
"""
    rc, stderr = _capture(check_workflows, _write(tmp_path, "ok.yml", body).parent)
    assert rc == 0, stderr


# --------------------------------------------------------------------------------------------
# CHECK 7 — workflow-level permissions.contents (issue #152, #117 audit)
# --------------------------------------------------------------------------------------------
def test_permissions_positive_contents_read(tmp_path: Path) -> None:
    body = """\
name: ok
on: [push]
permissions:
  contents: read
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    concurrency:
      group: gate-build
      cancel-in-progress: false
    steps:
      - run: echo ok
"""
    rc, stderr = _capture(check_workflows, _write(tmp_path, "ok.yml", body).parent)
    assert rc == 0, stderr
    assert "permissions" not in stderr


def test_permissions_positive_contents_read_in_superset(tmp_path: Path) -> None:
    """`contents: [read, write]` is accepted (the runner-skill rule is "no less than read")."""
    body = """\
name: ok
on: [push]
permissions:
  contents: [read, write]
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    concurrency:
      group: gate-build
      cancel-in-progress: false
    steps:
      - run: echo ok
"""
    rc, stderr = _capture(check_workflows, _write(tmp_path, "ok.yml", body).parent)
    assert rc == 0, stderr
    assert "permissions" not in stderr


def test_permissions_negative_contents_write(tmp_path: Path) -> None:
    body = """\
name: bad
on: [push]
permissions:
  contents: write
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    concurrency:
      group: gate-build
      cancel-in-progress: false
    steps:
      - run: echo ok
"""
    rc, stderr = _capture(check_workflows, _write(tmp_path, "bad.yml", body).parent)
    assert rc == 1
    assert "permissions" in stderr
    assert "contents" in stderr
    assert "'write'" in stderr


def test_permissions_warn_when_block_absent(tmp_path: Path) -> None:
    """Missing `permissions:` is a WARN, not an error.

    The default GITHUB_TOKEN is restrictive but not pinned to `read`. The
    gate warns so a future operator adding a job that needs `contents:
    write` cannot quietly escalate the token's scope. The gate does not
    fail today because the inherited default IS restrictive.
    """
    body = """\
name: warn
on: [push]
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    concurrency:
      group: gate-build
      cancel-in-progress: false
    steps:
      - run: echo ok
"""
    rc, stderr = _capture(check_workflows, _write(tmp_path, "warn.yml", body).parent)
    assert rc == 0, stderr
    assert "WARN" in stderr
    assert "permissions" in stderr


# --------------------------------------------------------------------------------------------
# CHECK 8 — env.PYTHON_VERSION consistency across workflows (issue #152)
# --------------------------------------------------------------------------------------------
def test_python_version_positive(tmp_path: Path) -> None:
    """Two workflows declaring the same value pass."""
    (tmp_path / "wf").mkdir()
    for name in ("a.yml", "b.yml"):
        (tmp_path / "wf" / name).write_text(
            """\
name: ok
on: [push]
permissions:
  contents: read
env:
  PYTHON_VERSION: "3.12.11"
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    concurrency:
      group: gate-build
      cancel-in-progress: false
    steps:
      - run: echo ok
""",
            encoding="utf-8",
        )
    rc, stderr = _capture(check_workflows, tmp_path / "wf")
    assert rc == 0, stderr
    assert "python-version-consistency" not in stderr


def test_python_version_positive_when_one_workflow_omits_it(tmp_path: Path) -> None:
    """A workflow that does not declare `env.PYTHON_VERSION` does not contribute.

    security.yml and security-deep.yml ship Docker images and never need
    setup-python at the workflow level. The check fires only when two
    declarations disagree; an absent declaration is not a disagreement.
    """
    (tmp_path / "wf").mkdir()
    (tmp_path / "wf" / "ci.yml").write_text(
        """\
name: ci
on: [push]
permissions:
  contents: read
env:
  PYTHON_VERSION: "3.12.11"
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    concurrency:
      group: gate-build
      cancel-in-progress: false
    steps:
      - run: echo ok
""",
        encoding="utf-8",
    )
    (tmp_path / "wf" / "security.yml").write_text(
        """\
name: security
on: [push]
permissions:
  contents: read
jobs:
  scan:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    concurrency:
      group: gate-scan
      cancel-in-progress: false
    steps:
      - run: echo ok
""",
        encoding="utf-8",
    )
    rc, stderr = _capture(check_workflows, tmp_path / "wf")
    assert rc == 0, stderr
    assert "python-version-consistency" not in stderr


def test_python_version_negative_when_workflows_disagree(tmp_path: Path) -> None:
    """Two declarations with two different values must fail."""
    (tmp_path / "wf").mkdir()
    (tmp_path / "wf" / "a.yml").write_text(
        """\
name: a
on: [push]
permissions:
  contents: read
env:
  PYTHON_VERSION: "3.12.11"
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    concurrency:
      group: gate-build
      cancel-in-progress: false
    steps:
      - run: echo ok
""",
        encoding="utf-8",
    )
    (tmp_path / "wf" / "b.yml").write_text(
        """\
name: b
on: [push]
permissions:
  contents: read
env:
  PYTHON_VERSION: "3.13.0"
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    concurrency:
      group: gate-build
      cancel-in-progress: false
    steps:
      - run: echo ok
""",
        encoding="utf-8",
    )
    rc, stderr = _capture(check_workflows, tmp_path / "wf")
    assert rc == 1
    assert "python-version-consistency" in stderr
    assert "3.12.11" in stderr
    assert "3.13.0" in stderr
    assert "a.yml" in stderr
    assert "b.yml" in stderr


# --------------------------------------------------------------------------------------------
# CHECK 9 — concurrency.group uniqueness within a workflow (issue #152)
# --------------------------------------------------------------------------------------------
def test_concurrency_group_uniqueness_positive(tmp_path: Path) -> None:
    """Each job with `concurrency.group` must have a distinct literal."""
    body = """\
name: ok
on: [push]
permissions:
  contents: read
jobs:
  a:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    concurrency:
      group: gate-a
      cancel-in-progress: false
    steps:
      - run: echo a
  b:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    concurrency:
      group: gate-b
      cancel-in-progress: false
    steps:
      - run: echo b
"""
    rc, stderr = _capture(check_workflows, _write(tmp_path, "ok.yml", body).parent)
    assert rc == 0, stderr
    assert "concurrency-group" not in stderr


def test_concurrency_group_uniqueness_negative(tmp_path: Path) -> None:
    """Two jobs sharing a group literal fail the gate.

    The literal `gate-quality` on two jobs means the second job queues
    behind the first — FIFO at a granularity nobody asked for, silently
    serialising the gate.
    """
    body = """\
name: bad
on: [push]
permissions:
  contents: read
jobs:
  a:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    concurrency:
      group: gate-quality
      cancel-in-progress: false
    steps:
      - run: echo a
  b:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    concurrency:
      group: gate-quality
      cancel-in-progress: false
    steps:
      - run: echo b
"""
    rc, stderr = _capture(check_workflows, _write(tmp_path, "bad.yml", body).parent)
    assert rc == 1
    assert "concurrency-group" in stderr
    assert "gate-quality" in stderr
    assert "'a'" in stderr
    assert "'b'" in stderr


def test_concurrency_group_uniqueness_allows_same_group_across_workflows(
    tmp_path: Path,
) -> None:
    """The check is per-workflow. The same literal in two files is fine."""
    (tmp_path / "wf").mkdir()
    shared = """\
name: {name}
on: [push]
permissions:
  contents: read
jobs:
  build:
    runs-on: ubuntu-24.04
    timeout-minutes: 10
    concurrency:
      group: gate-quality
      cancel-in-progress: false
    steps:
      - run: echo ok
"""
    (tmp_path / "wf" / "ci.yml").write_text(shared.format(name="ci"), encoding="utf-8")
    (tmp_path / "wf" / "security.yml").write_text(
        shared.format(name="security"), encoding="utf-8"
    )
    rc, stderr = _capture(check_workflows, tmp_path / "wf")
    assert rc == 0, stderr
    assert "concurrency-group" not in stderr
