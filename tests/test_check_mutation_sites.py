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
"""

from __future__ import annotations

import sys
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

import check_mutation_sites as gate  # noqa: E402  — sys.path tweak above

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
