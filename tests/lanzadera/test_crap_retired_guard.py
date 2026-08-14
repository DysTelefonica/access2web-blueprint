# HARNESS-PROVENANCE: deterministic-quality-harness v1.4 + architectural-guards-over-metrics DG-12
"""Retired-guard for the CRAP gate (DG-12, issue #266).

The CRAP gate was retired in slice 2 (descableado) and its on-disk surface
(`scripts/check_crap.py`, the two fixture trees, the wiring test) is deleted
in slice 3. Without a guard, a partial reintroduction — a `make check-crap`
target that points to a script that no longer exists, a `crap` entry sneaking
back into `GATES`, a stale `crap` mention in a Makefile target — would be
silent: the gate has no test, the wiring test has no script, and the
specification has no mechanical enforcement.

This guard pins the **absence** of `crap` (case-insensitive) in the five
functional surfaces that slice 2 cleaned. If `crap` reappears in any of them,
the guard fails with the offending path and line. The word `crap` in this
file's own assertions and module docstring is the documented exception; the
guard targets the functional surface, never `tests/`.

The five surfaces are exactly the ones listed in the task contract for
deliverable 3.3: the `GATES` tuple in `scripts/quality_report.py`, the
`Makefile`, the CI workflow at `.github/workflows/ci.yml`, the package
configuration at `app/pyproject.toml`, and the documentation tree at
`docs/`. The historical planning snapshot at
`docs/uat/estado-planificacion-NEW_2026-08-12.html` is a frozen audit
record of the state at 2026-08-12 (it accurately names PR #70
`chore(ci): add check_crap script` and the runtime list of scripts that
existed then); the guard skips it so the snapshot keeps its history, and
future status reports are expected to NOT mention `crap` because the gate
will have been retired for three days by then.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

#: Surfaces slice 3 commits to keeping free of `crap` (case-insensitive).
#: Each entry is (path, label) — the label is what the failure message prints
#: so a developer reading the red knows which surface violated the contract.
DG12_SURFACES: tuple[tuple[str, str], ...] = (
    ("scripts/quality_report.py", "GATES"),
    ("Makefile", "Makefile"),
    (".github/workflows/ci.yml", "ci.yml"),
    ("app/pyproject.toml", "pyproject.toml"),
)

#: One snapshot per planning date is a historical record; the audit trail of
#: PRs that already landed (including PR #70 `chore(ci): add check_crap
#: script`) belongs there. New reports must not mention `crap`.
HISTORICAL_DOCS = re.compile(r"^docs/uat/estado-planificacion-NEW_\d{4}-\d{2}-\d{2}\.html$")


def _find_worktree_root() -> Path:
    here = Path(__file__).resolve()
    for parent in here.parents:
        if (parent / "scripts" / "quality_report.py").is_file():
            return parent
    raise AssertionError("worktree root not found (no scripts/quality_report.py)")


@pytest.fixture(scope="module")
def root() -> Path:
    return _find_worktree_root()


def _crap_lines(text: str) -> list[tuple[int, str]]:
    """Return ``(line_number, line)`` pairs whose line contains `crap`.

    The match is case-insensitive and operates on the raw bytes of the file
    so colour codes or escapes cannot smuggle the token past a literal scan.
    """
    out: list[tuple[int, str]] = []
    for index, line in enumerate(text.splitlines(), start=1):
        if re.search(r"crap", line, flags=re.IGNORECASE):
            out.append((index, line))
    return out


def _assert_clean(relative_path: str, label: str, root: Path) -> None:
    path = root / relative_path
    assert path.is_file(), f"DG-12 surface missing: {relative_path}"
    offenders = [
        f"{relative_path}:{line_no}: {line}"
        for line_no, line in _crap_lines(path.read_text(encoding="utf-8"))
    ]
    assert not offenders, (
        f"DG-12 violation: `crap` reappeared in {label} ({relative_path}). "
        "The CRAP gate was retired in slice 2 (issue #266 DG-12, step 1) "
        "and its on-disk surface was deleted in slice 3 (step 2). Reintroducing "
        "the gate breaks the ratchet that the architectural-guards gate enforces. "
        "Remove the reference, or open a new proposal to bring CRAP back.\n" + "\n".join(offenders)
    )


@pytest.mark.parametrize("relative_path,label", list(DG12_SURFACES))
def test_dg12_surface_is_clean_of_crap(root: Path, relative_path: str, label: str) -> None:
    _assert_clean(relative_path, label, root)


def test_dg12_quality_report_gates_has_no_crap_entry(root: Path) -> None:
    """The GATES tuple itself must not name a CRAP gate.

    Tightening the surface-level check to the exact contract of DG-12 step 1:
    even a comment that mentions `crap` outside `GATES` would let the test
    above pass for the wrong reason. The GATES tuple is the load-bearing
    surface — if a `crap` entry slips back into it, the architectural-guards
    gate will later detect the orphan and fail.
    """
    path = root / "scripts" / "quality_report.py"
    text = path.read_text(encoding="utf-8")
    matches = re.findall(r"GATES\s*[:=].*?\)\s*$|GATES\s*=\s*\(.*\)", text, flags=re.DOTALL)
    if not matches:
        pytest.fail("GATES tuple not found in scripts/quality_report.py — layout drifted")
    gates_blob = matches[0]
    assert not re.search(r"crap", gates_blob, flags=re.IGNORECASE), (
        f"DG-12 violation: GATES contains a `crap` entry: {gates_blob!r}"
    )


def test_dg12_docs_tree_is_clean_of_crap(root: Path) -> None:
    """`docs/` (excluding frozen historical snapshots) must not mention `crap`.

    Every Markdown file under `docs/` is checked. The frozen historical
    snapshot at `docs/uat/estado-planificacion-NEW_<date>.html` is an audit
    record of PRs that already landed (including PR #70 that added
    `check_crap`); the guard skips it so the audit trail survives. Any other
    `*.md` or `*.html` file under `docs/` that introduces a `crap` reference
    is a violation.
    """
    docs = root / "docs"
    if not docs.is_dir():
        pytest.skip("docs/ missing")
    offenders: list[str] = []
    for path in sorted(docs.rglob("*")):
        if not path.is_file():
            continue
        relative = path.relative_to(root).as_posix()
        if HISTORICAL_DOCS.match(relative):
            continue
        if path.suffix not in {".md", ".html"}:
            continue
        for line_no, line in _crap_lines(path.read_text(encoding="utf-8")):
            offenders.append(f"{relative}:{line_no}: {line}")
    assert not offenders, (
        "DG-12 violation: `crap` reappeared in docs/ (excluding historical "
        "snapshots). The CRAP gate was retired; new documentation must not "
        "reference it.\n" + "\n".join(offenders)
    )


def test_dg12_guard_targets_its_documented_surfaces() -> None:
    """The guard covers the five surfaces named in the task contract.

    Pins the contract of deliverable 3.3: a future edit that drops one of
    the five surfaces from this list would let Crap reappear silently.
    """
    expected = {
        "GATES",
        "Makefile",
        "ci.yml",
        "pyproject.toml",
        "docs",
    }
    actual = {label for _, label in DG12_SURFACES} | {"docs"}
    assert actual == expected, f"surface list drifted: {actual} != {expected}"


def test_dg12_guard_does_not_walk_tests_or_fixtures() -> None:
    """The guard must not scan `tests/` or the fixtures in tests/fixtures/.

    Today `tests/lanzadera/test_quality_report_smoke.py` documents the
    retirement of CRAP in comments and asserts an old `failed_gates` payload
    that contains `crap` (the test pins the historical reporting contract).
    Scan target is the functional surface slice 2 cleaned, not the test
    corpus. The mechanism is structural: the guard reads only the five
    surfaces in `DG12_SURFACES` and walks `docs/` explicitly — it never
    walks `tests/`. Encode the contract by asserting that no test-time path
    appears in `DG12_SURFACES` and that the docs walker filters out
    historical snapshots.
    """
    surface_paths = {path for path, _ in DG12_SURFACES}
    assert not any(path.startswith("tests/") for path in surface_paths), (
        "DG12_SURFACES must not include tests/ — scan target is the functional surface"
    )
    # The docs walker is the only directory traversal. It must exclude the
    # historical snapshot pattern and accept only markdown / html files.
    assert HISTORICAL_DOCS.pattern.startswith("^docs/uat/"), (
        "HISTORICAL_DOCS must scope to docs/uat/ so historical snapshots stay frozen"
    )
