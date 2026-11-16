#!/usr/bin/env python3
"""Tests de no-coste en PR ordinario (#1627, slice 8): el ejemplo de CI de la guía."""

import json
import re
import sys
import textwrap
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
# La raíz de la skill (ci-pattern) está siempre a tres niveles del fichero, en el catálogo y en el
# consumer (#1700): la guía se busca ahí, nunca en una ruta fija del catálogo.
CI_PATTERN = Path(__file__).resolve().parents[3]
GUIDE = CI_PATTERN / "references" / "porting-guide.md"
sys.path.insert(0, str(ASSET))
sys.path.insert(0, str(ASSET.parent / "workflow-policy"))
sys.path.insert(0, str(ASSET.parent / "required-jobs"))
import check_required_jobs as aggregator  # noqa: E402
import check_workflow_policy as policy  # noqa: E402
import night_gates  # noqa: E402

EXPENSIVE = ("upgrade-path", "integration")
SQL_OUTPUT = "needs.changes.outputs.sql == 'true'"
ORDINARY_PR = ["src/app.py", "docs/README.md", "tests/test_app.py", ".github/workflows/pr-size.yml"]
MIGRATION_PR = ["src/app.py", "db/migrations/0002_add_column.sql"]


def template_jobs() -> dict:
    """Los jobs del ejemplo de CI de la guía, leídos con el lector del gate."""
    text = GUIDE.read_text(encoding="utf-8")
    fences = re.findall(r"```yaml\n(.*?)```", text, re.DOTALL)
    example = next(f for f in fences if "upgrade-path:" in f)
    body = textwrap.indent(textwrap.dedent(example), "  ")
    return policy.parse_workflow("on:\n  pull_request:\njobs:\n" + body, "ci.yml")["jobs"]


def sql_filter_patterns(jobs: dict) -> list[str]:
    """Patrones del filtro `sql:` del job `changes` (dorny/paths-filter)."""
    filters = next(step["with"]["filters"] for step in jobs["changes"]["steps"]
                   if "paths-filter" in str(step.get("uses")))
    patterns, inside = [], False
    for line in filters.splitlines():
        stripped = line.strip()
        if stripped == "sql:":
            inside = True
        elif inside and stripped.startswith("- "):
            patterns.append(stripped[2:].strip().strip("'\""))
        elif inside and stripped and not stripped.startswith("-"):
            inside = False
    return patterns


def sql_changed(patterns: list[str], files: list[str]) -> bool:
    return any(night_gates.path_matches(f, p) for f in files for p in patterns)


def expensive_runs(jobs: dict, files: list[str]) -> dict[str, bool]:
    """Para cada job caro: si su condición deja correrlo con estos ficheros."""
    value = "true" if sql_changed(sql_filter_patterns(jobs), files) else "false"
    runs = {}
    for name in EXPENSIVE:
        if name in jobs:
            condition = jobs[name].get("if", "")
            runs[name] = condition.strip() == SQL_OUTPUT and value == "true"
    return runs


class NoCostOrdinaryPrTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.jobs = template_jobs()

    def test_the_template_declares_the_upgrade_path_job(self):
        self.assertIn("upgrade-path", self.jobs)

    def test_every_expensive_job_depends_on_the_sql_gate(self):
        for name in EXPENSIVE:
            if name not in self.jobs:
                continue
            with self.subTest(job=name):
                self.assertIn("changes", self.jobs[name].get("needs", []))
                self.assertEqual(self.jobs[name].get("if", "").strip(), SQL_OUTPUT)

    def test_the_sql_filter_matches_migrations_and_not_ordinary_files(self):
        patterns = sql_filter_patterns(self.jobs)
        self.assertTrue(sql_changed(patterns, ["db/migrations/0002_add_column.sql"]))
        self.assertTrue(sql_changed(patterns, ["schema.sql"]))
        self.assertFalse(sql_changed(patterns, ORDINARY_PR))

    def test_an_ordinary_pr_runs_no_expensive_job(self):
        self.assertEqual(set(expensive_runs(self.jobs, ORDINARY_PR).values()), {False})

    def test_a_pr_that_changes_sql_runs_the_expensive_jobs(self):
        runs = expensive_runs(self.jobs, MIGRATION_PR)
        self.assertIn("upgrade-path", runs)
        self.assertTrue(runs["upgrade-path"])


def guide_policy() -> dict:
    """La política de required-jobs que la guía declara para upgrade-path."""
    text = GUIDE.read_text(encoding="utf-8")
    fences = re.findall(r"```json\n(.*?)```", text, re.DOTALL)
    return json.loads(next(f for f in fences if "upgrade-path" in f and "accepted_skips" in f))


class RequiredAggregatorTests(unittest.TestCase):
    """El agregador `required` acepta un upgrade-path omitido solo si la política lo declara."""

    def test_the_guide_declares_the_skip_for_pull_requests(self):
        needs = {"changes": "success", "upgrade-path": "skipped"}
        verdict = aggregator.evaluate(needs, guide_policy(), "pull_request")
        self.assertTrue(verdict["passed"], verdict)

    def test_a_pr_with_sql_runs_upgrade_path_and_passes(self):
        needs = {"changes": "success", "upgrade-path": "success"}
        verdict = aggregator.evaluate(needs, guide_policy(), "pull_request")
        self.assertTrue(verdict["passed"], verdict)

    def test_without_the_declaration_the_skip_fails(self):
        policy_without = {"required_jobs": ["changes", "upgrade-path"],
                          "events": {"pull_request": {"accepted_skips": []}}}
        needs = {"changes": "success", "upgrade-path": "skipped"}
        verdict = aggregator.evaluate(needs, policy_without, "pull_request")
        self.assertFalse(verdict["passed"], verdict)


if __name__ == "__main__":
    unittest.main(verbosity=2)
