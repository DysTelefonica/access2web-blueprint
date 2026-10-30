#!/usr/bin/env python3
# ci-pattern asset tests — clasificador de alcance por suite (#1440)
"""Cada regla del clasificador ``classify_changes.py`` a través de su CLI, y el
ejemplo de política del README ejecutado por ``check_required_jobs.py`` para que
la documentación y el gate no puedan derivar. Stdlib only."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
SCRIPT = ASSET / "classify_changes.py"
README = ASSET / "README.md"
EXAMPLE = ASSET / "change-scope.example.json"
GATE = ASSET.parent / "required-jobs" / "check_required_jobs.py"

SCOPE = {
    "version": 1,
    "suites": {"phpunit": ["www/**", "tests/**", "composer.*"],
               "e2e": ["e2e/**", "www/templates/**"]},
    "always": [".github/workflows/**"],
}


def run(changed: list[str] | None, scope: object = SCOPE, sealed: list[str] | None = None,
        raw_changed: str | None = None) -> tuple[subprocess.CompletedProcess, str | None]:
    """Ejecuta el clasificador en un directorio temporal; devuelve el proceso y GITHUB_OUTPUT."""
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        scope_path = root / "change-scope.json"
        scope_path.write_text(scope if isinstance(scope, str) else json.dumps(scope),
                              encoding="utf-8")
        changed_path = root / "changed.txt"
        if raw_changed is not None:
            changed_path.write_text(raw_changed, encoding="utf-8")
        elif changed is not None:
            changed_path.write_text("".join(f"{p}\n" for p in changed), encoding="utf-8")
        output = root / "github_output"
        args = [sys.executable, str(SCRIPT), "--changed-files", str(changed_path),
                "--scope", str(scope_path), "--github-output", str(output)]
        if sealed is not None:
            sealed_path = root / "sealed.txt"
            sealed_path.write_text("\n".join(sealed) + "\n", encoding="utf-8")
            args += ["--sealed-files", str(sealed_path)]
        proc = subprocess.run(args, capture_output=True, text=True, encoding="utf-8")
        written = output.read_text(encoding="utf-8") if output.exists() else None
    return proc, written


def flags(text: str) -> dict[str, str]:
    return dict(line.split("=", 1) for line in text.splitlines() if line.strip())


ALL_TRUE = {"phpunit_needed": "true", "e2e_needed": "true"}
ALL_FALSE = {"phpunit_needed": "false", "e2e_needed": "false"}


class TestClassifyRules(unittest.TestCase):

    def assert_flags(self, changed: list[str], expected: dict[str, str], **kw) -> None:
        proc, written = run(changed, **kw)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(flags(proc.stdout), expected)
        self.assertEqual(flags(written or ""), expected)

    def test_path_in_one_suite_list_needs_only_that_suite(self) -> None:
        self.assert_flags(["www/src/Foo.php"], {"phpunit_needed": "true", "e2e_needed": "false"})

    def test_docs_and_fleet_config_need_no_suite(self) -> None:
        self.assert_flags(["docs/guide.md", "README.md", "fleet/registry.json"], ALL_FALSE)

    def test_always_path_needs_every_suite(self) -> None:
        self.assert_flags(["docs/a.md", ".github/workflows/ci.yml"], ALL_TRUE)

    def test_ci_pattern_assets_need_every_suite(self) -> None:
        self.assert_flags([".agents/skills/ci-pattern/assets/budget/vendored_budget.py"],
                          ALL_TRUE)

    def test_scope_declaration_itself_needs_every_suite(self) -> None:
        self.assert_flags([".github/change-scope.json"], ALL_TRUE)

    def test_sealed_file_needs_every_suite(self) -> None:
        self.assert_flags([".github/required-jobs-policy.json"], ALL_TRUE,
                          sealed=[".github/required-jobs-policy.json"])

    def test_unsealed_neighbour_of_a_sealed_file_does_not(self) -> None:
        self.assert_flags([".github/labels.json"], ALL_FALSE,
                          sealed=[".github/required-jobs-policy.json"])

    def test_empty_diff_needs_every_suite(self) -> None:
        proc, written = run(None, raw_changed="\n\n")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertEqual(flags(proc.stdout), ALL_TRUE)
        self.assertEqual(flags(written or ""), ALL_TRUE)

    def test_quoted_path_from_git_needs_every_suite(self) -> None:
        self.assert_flags(['"docs/caf\\303\\251.md"'], ALL_TRUE)

    def test_crlf_lines_are_read_as_paths(self) -> None:
        proc, _ = run(None, raw_changed="www/a.php\r\ndocs/b.md\r\n")
        self.assertEqual(flags(proc.stdout), {"phpunit_needed": "true", "e2e_needed": "false"})

    def test_double_star_semantics_match_vendored_budget(self) -> None:
        scope = {"version": 1, "suites": {"docs": ["docs/**.md"], "deep": ["src/**"]}}
        proc, _ = run(["docs/a/b.md"], scope=scope)
        self.assertEqual(flags(proc.stdout), {"docs_needed": "false", "deep_needed": "false"})
        proc, _ = run(["docs/b.md", "src/a/b/c.py"], scope=scope)
        self.assertEqual(flags(proc.stdout), {"docs_needed": "true", "deep_needed": "true"})

    def test_basename_pattern_matches_at_any_depth(self) -> None:
        self.assert_flags(["sub/composer.lock"], {"phpunit_needed": "true", "e2e_needed": "false"})


class TestFailClosed(unittest.TestCase):

    def assert_usage_error(self, proc: subprocess.CompletedProcess, written: str | None,
                           needle: str) -> None:
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn(needle, proc.stderr)
        self.assertEqual(proc.stdout, "")
        self.assertIsNone(written, "an invalid run must not leave partial outputs")

    def test_unreadable_changed_files_exits_2(self) -> None:
        self.assert_usage_error(*run(None), "cannot read --changed-files")

    def test_unreadable_scope_json_exits_2(self) -> None:
        self.assert_usage_error(*run(["www/a.php"], scope="{not json"), "invalid scope")

    def test_invalid_scopes_exit_2(self) -> None:
        cases = {
            "version": {"version": 2, "suites": {"a": ["x/**"]}},
            "unknown key": {"version": 1, "suites": {"a": ["x/**"]}, "paths": []},
            "no suites": {"version": 1, "suites": {}},
            "empty list": {"version": 1, "suites": {"a": []}},
            "non-string": {"version": 1, "suites": {"a": ["x/**", 3]}},
            "bad suite name": {"version": 1, "suites": {"a b": ["x/**"]}},
            "empty always": {"version": 1, "suites": {"a": ["x/**"]}, "always": []},
            "trailing slash": {"version": 1, "suites": {"a": ["x/"]}},
            "negation": {"version": 1, "suites": {"a": ["!x/**"]}},
            "not an object": ["x/**"],
        }
        for name, scope in cases.items():
            with self.subTest(name):
                self.assert_usage_error(*run(["x/a"], scope=scope), "invalid scope")

    def test_unreadable_sealed_files_exits_2(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            changed = Path(tmp) / "changed.txt"
            changed.write_text("www/a.php\n", encoding="utf-8")
            scope = Path(tmp) / "scope.json"
            scope.write_text(json.dumps(SCOPE), encoding="utf-8")
            proc = subprocess.run(
                [sys.executable, str(SCRIPT), "--changed-files", str(changed), "--scope",
                 str(scope), "--sealed-files", str(Path(tmp) / "missing.txt")],
                capture_output=True, text=True, encoding="utf-8")
        self.assertEqual(proc.returncode, 2, proc.stderr)
        self.assertIn("cannot read --sealed-files", proc.stderr)


class TestDocumentedContract(unittest.TestCase):

    def test_example_scope_is_valid(self) -> None:
        proc, _ = run(["docs/a.md"], scope=EXAMPLE.read_text(encoding="utf-8"))
        self.assertEqual(proc.returncode, 0, proc.stderr)

    def readme_policy(self) -> dict:
        for fence in README.read_text(encoding="utf-8").split("```json")[1:]:
            block = fence.split("```", 1)[0]
            if "marker_lanes" in block:
                return json.loads(block)
        self.fail("README has no ```json policy declaring marker_lanes")

    def run_gate(self, phpunit_needed: str, skipped: tuple[str, ...]):
        policy = self.readme_policy()
        needs = {job: {"result": "skipped" if job in skipped else "success",
                       "outputs": ({"phpunit_needed": phpunit_needed}
                                   if job == "change-scope" else {})}
                 for job in policy["required_jobs"]}
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "policy.json"
            path.write_text(json.dumps(policy), encoding="utf-8")
            return subprocess.run(
                [sys.executable, str(GATE), "--policy", str(path), "--event", "pull_request"],
                input=json.dumps(needs), capture_output=True, text=True, encoding="utf-8")

    def test_readme_policy_accepts_the_skip_when_the_classifier_proves_it(self) -> None:
        proc = self.run_gate("false", ("phpunit",))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("VERDICT: PASS", proc.stdout)

    def test_readme_policy_rejects_the_skip_when_the_suite_is_needed(self) -> None:
        proc = self.run_gate("true", ("phpunit",))
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("not declared for this event", proc.stdout)

    def test_readme_policy_declares_relaxation(self) -> None:
        self.assertIn("relaxation", self.readme_policy())


if __name__ == "__main__":
    unittest.main()
