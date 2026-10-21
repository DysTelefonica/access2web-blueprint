#!/usr/bin/env python3
# ci-pattern asset tests — in-code baseline relaxation guard (DysTelefonica/team-skills#974, Refs #429)
"""Executable suite for ``check_baseline_relaxation.py`` (stdlib only).

Covers HR-15 for baselines declared as constants in a gate script: a number
that grows, a new entry or a new element needs a fresh ``RELAXATION`` field;
lowering is free; an unreadable base fails closed.
"""

from __future__ import annotations

import importlib.util
import io
import os
import subprocess
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
SCRIPT = ASSET / "check_baseline_relaxation.py"
FIELD = 'RELAXATION = {"reason": "split lands next slice", "reference": "#1528", "issue": 1528}\n'


def load_guard():
    spec = importlib.util.spec_from_file_location("check_baseline_relaxation", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def gate(sites: int, extra: str = "") -> str:
    return f'BASELINE: dict[str, int] = {{"app/a.py": {sites}}}\n{extra}'


class SourceRules(unittest.TestCase):
    def setUp(self) -> None:
        self.guard = load_guard()

    def check(self, base: str, head: str) -> list[str]:
        return self.guard.check_sources("scripts/g.py", base, head)

    def test_raising_without_field_is_rejected(self) -> None:
        violations = self.check(gate(362), gate(426))
        self.assertEqual(len(violations), 1)
        self.assertIn("'app/a.py': 362 -> 426", violations[0])
        self.assertIn("RELAXATION", violations[0])

    def test_raising_with_reason_reference_and_identity_is_accepted(self) -> None:
        self.assertEqual(self.check(gate(362), gate(426, FIELD)), [])

    def test_lowering_is_accepted(self) -> None:
        self.assertEqual(self.check(gate(362), gate(300)), [])

    def test_new_entry_is_rejected_and_removed_entry_is_accepted(self) -> None:
        base = 'BASELINE = {"a": 1, "b": 2}\n'
        self.assertEqual(self.check(base, 'BASELINE = {"a": 1}\n'), [])
        added = self.check(base, 'BASELINE = {"a": 1, "b": 2, "c": 3}\n')
        self.assertIn("new entry 'c'", added[0])

    def test_new_set_element_is_rejected(self) -> None:
        base = 'BASELINE_X: frozenset[str] = frozenset({"a"})\n'
        violations = self.check(base, 'BASELINE_X = frozenset({"a", "b"})\n')
        self.assertIn("new element 'b'", violations[0])

    def test_field_without_valid_shape_or_identity_does_not_cover(self) -> None:
        for field in (
            'RELAXATION = {"reason": "", "reference": "#1", "issue": 1}\n',
            'RELAXATION = {"reason": "why", "issue": 1}\n',
            'RELAXATION = {"reason": "why", "reference": "#1"}\n',
        ):
            with self.subTest(field=field):
                self.assertTrue(self.check(gate(1), gate(2, field)))

    def test_actor_is_an_accepted_identity(self) -> None:
        field = 'RELAXATION = {"reason": "why", "reference": "x", "actor": "maintainer"}\n'
        self.assertEqual(self.check(gate(1), gate(2, field)), [])

    def test_field_inherited_from_the_base_cannot_cover_a_new_change(self) -> None:
        violations = self.check(gate(1, FIELD), gate(2, FIELD))
        self.assertIn("already carries", violations[0])

    def test_baseline_that_stops_being_a_literal_is_rejected(self) -> None:
        violations = self.check(gate(1), "BASELINE = load_baseline()\n")
        self.assertIn("no longer a readable literal", violations[0])


def git(root: Path, *args: str) -> None:
    subprocess.run(["git", "-C", str(root), "-c", "user.name=t", "-c", "user.email=t@t",
                    *args], check=True, capture_output=True)


class CliAgainstGit(unittest.TestCase):
    def setUp(self) -> None:
        self.guard = load_guard()
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)

    def repo_with(self, head_source: str, directory: str = "scripts") -> None:
        (self.root / directory).mkdir(parents=True)
        path = self.root / directory / "gate.py"
        path.write_text(gate(362), encoding="utf-8")
        git(self.root, "init", "-q", "-b", "main")
        git(self.root, "add", "-A")
        git(self.root, "commit", "-qm", "base")
        git(self.root, "checkout", "-q", "-b", "feature")
        path.write_text(head_source, encoding="utf-8")
        git(self.root, "commit", "-qam", "head")

    def run_cli(self, *args: str) -> tuple[int, str]:
        out = io.StringIO()
        with redirect_stdout(out):
            code = self.guard.main(["--root", str(self.root), *args])
        return code, out.getvalue()

    def test_branch_that_raises_a_baseline_fails(self) -> None:
        self.repo_with(gate(426))
        code, out = self.run_cli("--base-ref", "main")
        self.assertEqual(code, self.guard.EXIT_RELAXED)
        self.assertIn("scripts/gate.py", out)

    def test_branch_that_lowers_a_baseline_passes(self) -> None:
        self.repo_with(gate(300))
        code, out = self.run_cli("--base-ref", "main")
        self.assertEqual(code, self.guard.EXIT_OK)
        self.assertIn("OK (compared with main)", out)

    def test_unknown_base_ref_fails_closed(self) -> None:
        self.repo_with(gate(300))
        code, _ = self.run_cli("--base-ref", "nope")
        self.assertEqual(code, self.guard.EXIT_UNREADABLE)

    def test_gate_directory_is_configurable(self) -> None:
        self.repo_with(gate(426), directory="ci/gates")
        code, _ = self.run_cli("--base-ref", "main", "--gate-glob", "ci/gates/*.py")
        self.assertEqual(code, self.guard.EXIT_RELAXED)
        code, _ = self.run_cli("--base-ref", "main")
        self.assertEqual(code, self.guard.EXIT_OK)

    def test_without_a_base_the_comparison_is_declared_not_performed(self) -> None:
        saved = os.environ.pop("GITHUB_BASE_REF", None)
        self.addCleanup(lambda: os.environ.update({"GITHUB_BASE_REF": saved}) if saved else None)
        code, out = self.run_cli()
        self.assertEqual(code, self.guard.EXIT_OK)
        self.assertIn("NOT performed", out)


if __name__ == "__main__":
    loader = unittest.defaultTestLoader
    suite = unittest.TestSuite(loader.loadTestsFromTestCase(c) for c in (SourceRules, CliAgainstGit))
    sys.exit(not unittest.TextTestRunner().run(suite).wasSuccessful())
