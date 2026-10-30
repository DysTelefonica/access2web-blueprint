# HARNESS-PROVENANCE: access2web-blueprint issue #947 — test_check_mutation_sites.py
"""Wiring pin for ``scripts/check_mutation_sites.py``.

The ratchet counts mutation-surface "sites" — every ``ast.Constant``
``int | float | str | bool`` is one site, plus ``ast.Call`` args +
keywords, plus a handful of mutable AST node types (``If``, ``Raise``,
``Compare``, etc.). Issue #947 reports that this counter also picks
up docstrings, which silently rewards deleting documentation to push
a BASELINE entry down. These tests pin the contract:

  1. Docstrings at module / function / class scope do NOT count as
     mutation sites. A file with a docstring must report the same
     number of sites as the same file without it.
  2. String literals that are NOT docstrings (a constant inside a
     function body, an argument to a call, a default value) still
     count — the heuristic must be precise.
  3. Multi-line docstrings still don't count.
  4. Empty docstrings don't count either.

Each case is a pair: ``with_docstring`` and ``without_docstring`` parse
the same AST except for the docstring. With the fix, the counter must
return the same number for both.

BP's copy of the canon is measured at an explicit policy date and its report is
bound to ``quality-policy.json``. The docstring exclusion above is not a local
deviation: the canon carries it since team-skills#1008, so BP declares as local
only its ``BASELINE``, which is data.

  1. ``--policy-date`` is mandatory: the ratchet never reads the wall clock.
  2. The envelope carries the policy date, the policy's ceilings, the subjects
     it measured and the policy identity.
  3. The header declares the canon version AND the deviation it keeps, so a
     later edit that drops the declaration is caught here.

Each case runs the gate as a subprocess, because the wiring under test is the
command line the CI job and ``quality_report.py`` use.
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

import check_mutation_sites as gate  # noqa: E402  — sys.path tweak above

#: DQH 2.2: the date is explicit, never the wall clock. Same value the other
#: dated gates use in tests/test_gate_smoke.py and tests/test_gate_liveness.py.
POLICY_DATE = "2026-10-21"


def _write_surface(root: Path, sites: int = 1) -> None:
    package = root / "app" / "pkg"
    package.mkdir(parents=True)
    body = "\n".join(f"    x = x + {n}" for n in range(sites))
    (package / "surface.py").write_text(
        f"def f(x: int) -> int:\n{body}\n    return x\n", encoding="utf-8"
    )


def _run(root: Path, *extra: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(SCRIPTS_DIR / "check_mutation_sites.py"), "--root", str(root), *extra],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )


# ---------------------------------------------------------------------------
# Pairs — with_docstring and without_docstring are the same code except
# for the docstring on the relevant scope. The counter MUST agree.
# ---------------------------------------------------------------------------


MODULE_WITH_DOC = '"""Top-level docstring."""\nimport os\nx = 1\n'
MODULE_WITHOUT_DOC = "import os\nx = 1\n"


FUNCTION_WITH_DOC = 'def f(x: int) -> int:\n    """Function docstring."""\n    return x + 1\n'
FUNCTION_WITHOUT_DOC = "def f(x: int) -> int:\n    return x + 1\n"


CLASS_WITH_DOC = 'class C:\n    """Class docstring."""\n    def m(self) -> int:\n        return 1\n'
CLASS_WITHOUT_DOC = "class C:\n    def m(self) -> int:\n        return 1\n"


MULTILINE_WITH_DOC = (
    '"""Multi-line\n\nmodule docstring spanning several lines.\n"""\nimport os\nx = 1\n'
)
MULTILINE_WITHOUT_DOC = "import os\nx = 1\n"


EMPTY_DOC = '"""\n"""\nimport os\nx = 1\n'
NO_DOC = "import os\nx = 1\n"


# String literal that is NOT a docstring: it sits inside a function body
# after a non-string statement, so the heuristic must NOT skip it.
NON_DOCSTRING_AT_TOP_OF_BODY = (
    "def f() -> None:\n"
    "    x = 1\n"  # non-string first statement
    '    label = "hello"  # a real string, not a docstring\n'
    "    return None\n"
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _sites(src: str) -> int:
    import ast

    return gate._count_sites(ast.parse(src))


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------


def test_module_docstring_does_not_count() -> None:
    """A module-level docstring adds zero mutation sites."""
    assert _sites(MODULE_WITH_DOC) == _sites(MODULE_WITHOUT_DOC)


def test_function_docstring_does_not_count() -> None:
    """A function-level docstring adds zero mutation sites."""
    assert _sites(FUNCTION_WITH_DOC) == _sites(FUNCTION_WITHOUT_DOC)


def test_class_docstring_does_not_count() -> None:
    """A class-level docstring adds zero mutation sites."""
    assert _sites(CLASS_WITH_DOC) == _sites(CLASS_WITHOUT_DOC)


def test_multiline_module_docstring_does_not_count() -> None:
    """Multi-line docstrings are still docstrings: zero extra sites."""
    assert _sites(MULTILINE_WITH_DOC) == _sites(MULTILINE_WITHOUT_DOC)


def test_empty_docstring_does_not_count() -> None:
    """An empty docstring is still a docstring: zero extra sites."""
    assert _sites(EMPTY_DOC) == _sites(NO_DOC)


def test_non_docstring_string_literal_still_counts() -> None:
    """A string literal that is NOT a docstring must still count.

    The heuristic must skip docstrings specifically, not "the first
    string in a function body". A ``label = "hello"`` after a
    non-string first statement is not a docstring.
    """
    baseline = _sites(NON_DOCSTRING_AT_TOP_OF_BODY)
    # Removing the non-docstring string should drop the count by exactly 1
    # (the Constant str for "hello"), because it's the only string literal
    # in the body.
    without_string = "def f() -> None:\n    x = 1\n    return None\n"
    assert _sites(without_string) == baseline - 1


def test_baseline_main_py_lands_at_pre_941_floor() -> None:
    """Regression pin: after the fix, ``app/src/main.py`` drops from 121
    to 116 (the five docstrings no longer count).

    Why 116 and not 118: ``app/src/main.py`` carries five docstrings
    (module, ``_ContainerMirrorMiddleware``,
    ``_ensure_container_middleware``, ``_LazyAuthSigner``,
    ``_register_admin_routes``), not three. The pre-#946 review
    listed only the three structural class/function docstrings; the
    module docstring and ``_ensure_container_middleware`` docstring
    were already on the file and were not on the review's radar. The
    fix strips all five, taking the count from 121 to 116 — two below
    the 118 the pre-#946 comment claimed.

    This is the headline number for the changelog of #947 / #945 and
    the ratchet lock-in. If the count drifts away from 116 here, the
    fix has regressed (somebody added or removed a docstring without
    updating the test, or one of the five scopes disappeared).
    """
    import ast
    from pathlib import Path

    main_py = Path(__file__).resolve().parent.parent / "app" / "src" / "main.py"
    assert main_py.is_file(), "app/src/main.py must exist for the regression pin"
    count = gate._count_sites(ast.parse(main_py.read_text(encoding="utf-8")))
    assert count == 116, (
        f"app/src/main.py is at {count} sites after the fix; expected 116 "
        f"(the five docstrings should no longer count). If this "
        f"test fails after a legitimate refactor, regenerate the BASELINE "
        f"with --emit-baseline and update the test together."
    )


# ---------------------------------------------------------------------------
# DQH 2.2 policy binding (team-skills#809 T6)
# ---------------------------------------------------------------------------


def test_the_policy_date_is_required(tmp_path: Path) -> None:
    _write_surface(tmp_path)
    result = _run(tmp_path)
    assert result.returncode == 2, result.stdout + result.stderr
    assert "--policy-date" in result.stderr


def test_envelope_is_bound_to_the_policy_and_its_date(tmp_path: Path) -> None:
    _write_surface(tmp_path)
    result = _run(tmp_path, "--policy-date", POLICY_DATE, "--json")
    envelope = json.loads(result.stdout)
    assert result.returncode == 0, result.stdout + result.stderr
    assert envelope["policy_date"] == POLICY_DATE
    assert envelope["ceilings"] == {"max_mutation_sites": 100, "files_over_ceiling": 0}
    assert envelope["subjects"]["checked"] == ["app/pkg/surface.py"]
    assert envelope["policy_identity"].startswith("sha256:")
    assert envelope["indicators"]["files_measured"] == 1


def test_the_header_declares_the_canon_and_only_its_data_as_deviation() -> None:
    """The canon carries the docstring exclusion since team-skills#1008, so this copy does not.

    A header that names a deviation which no longer exists is the silent drift this gate's
    provenance line exists to prevent: the next reader would look for a difference that is not
    there. Sliced to the header block, which ends well before the first function.
    """
    header = (SCRIPTS_DIR / "check_mutation_sites.py").read_text(encoding="utf-8")[:1200]
    assert "HARNESS-PROVENANCE: deterministic-quality-harness v2.2" in header
    assert "BASELINE" in header, "the tolerated site counts are BP data and stay declared as such"
    assert "_docstring_node_ids" not in header, (
        "the canon publishes the docstring exclusion (team-skills#1008); declaring it here as a "
        "local deviation would claim a difference that no longer exists"
    )
