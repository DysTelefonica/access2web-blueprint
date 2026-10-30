# HARNESS-PROVENANCE: deterministic-quality-harness v2.2 — assets/tests/test_mutation_sites.py
"""Contract of ``_count_sites``: documentation is not mutation surface.

A docstring is an ``ast.Constant`` holding a ``str``, which is exactly what the counter charges as
a site. Left uncorrected, a ratchet built on that number pays the author for deleting documentation:
the file measures fewer sites once its docstrings are gone, and the BASELINE entry can be lowered on
paper without a single behavioural change.

These tests pin the fix (``_docstring_node_ids``): a file with a docstring and the same file without
it must report the same number, at module, function, class, multi-line and empty-docstring scope; a
string literal that is *not* a docstring must still count. The last test proves it at the public
interface — the gate's verdict on a file above the ceiling does not change when it is documented.
"""

from __future__ import annotations

import ast
import importlib.util
import json
import subprocess
import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"
POLICY_DATE = "2026-10-21"

MODULE_WITH = '"""Top-level docstring."""\nimport os\nx = 1\n'
MODULE_WITHOUT = "import os\nx = 1\n"
FUNCTION_WITH = (
    'def f(x: int) -> int:\n    """Function docstring."""\n    return x + 1\n'
)
FUNCTION_WITHOUT = "def f(x: int) -> int:\n    return x + 1\n"
CLASS_WITH = (
    'class C:\n    """Class docstring."""\n    def m(self) -> int:\n        return 1\n'
)
CLASS_WITHOUT = "class C:\n    def m(self) -> int:\n        return 1\n"
MULTILINE_WITH = (
    '"""Multi-line\n\nmodule docstring spanning several lines.\n"""\nimport os\nx = 1\n'
)
MULTILINE_WITHOUT = "import os\nx = 1\n"
EMPTY_WITH = '"""\n"""\nimport os\nx = 1\n'
#: A literal that is not a docstring: a string statement after a non-string first statement.
NOT_A_DOCSTRING = 'def f() -> None:\n    x = 1\n    label = "hello"\n    return None\n'


def _load():
    if str(SCRIPTS) not in sys.path:
        sys.path.insert(0, str(SCRIPTS))
    spec = importlib.util.spec_from_file_location(
        "check_mutation_sites", SCRIPTS / "check_mutation_sites.py"
    )
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    # Register before execution: the dataclasses resolve their annotations through sys.modules.
    sys.modules["check_mutation_sites"] = module
    spec.loader.exec_module(module)
    return module


def _sites(source: str) -> int:
    return _load()._count_sites(ast.parse(source))


@pytest.mark.parametrize(
    "with_docstring, without",
    [
        (MODULE_WITH, MODULE_WITHOUT),
        (FUNCTION_WITH, FUNCTION_WITHOUT),
        (CLASS_WITH, CLASS_WITHOUT),
        (MULTILINE_WITH, MULTILINE_WITHOUT),
        (EMPTY_WITH, MODULE_WITHOUT),
    ],
    ids=["module", "function", "class", "multiline", "empty"],
)
def test_a_docstring_is_not_a_mutation_site(with_docstring: str, without: str) -> None:
    assert _sites(with_docstring) == _sites(without)


def test_a_string_literal_that_is_not_a_docstring_still_counts() -> None:
    """The heuristic must skip docstrings specifically, not the first string in a body."""
    without = "def f() -> None:\n    x = 1\n    return None\n"
    assert _sites(NOT_A_DOCSTRING) == _sites(without) + 1


def test_documentation_does_not_change_the_verdict(tmp_path: Path) -> None:
    """At the public interface: the same file, documented or not, is judged identically.

    A file above the ceiling is a failure either way; what must not change is the number the gate
    reports for it, because that number is what a BASELINE entry is written against.
    """
    gate = SCRIPTS / "check_mutation_sites.py"
    body = "\n".join(f"    x = x + {n}" for n in range(60))
    plain = f"def f(x: int) -> int:\n{body}\n    return x\n"
    documented = f'"""Wide on purpose."""\n\n\n{plain}'

    verdicts = []
    for name, source in (("plain.py", plain), ("documented.py", documented)):
        root = tmp_path / name.removesuffix(".py")
        package = root / "app" / "pkg"
        package.mkdir(parents=True)
        (package / "wide.py").write_text(source, encoding="utf-8")
        result = subprocess.run(
            [
                sys.executable,
                str(gate),
                "--root",
                str(root),
                "--json",
                "--policy-date",
                POLICY_DATE,
            ],
            capture_output=True,
            text=True,
            check=False,
            encoding="utf-8",
        )
        envelope = json.loads(result.stdout)
        verdicts.append(
            (result.returncode, envelope["indicators"]["max_mutation_sites"])
        )

    assert verdicts[0] == verdicts[1], f"documentation changed the verdict: {verdicts}"
    assert verdicts[0][0] == 1, (
        "the fixture must sit above the ceiling for this test to mean anything"
    )
