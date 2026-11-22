#!/usr/bin/env python3
"""Suite del modo `advise` de `checks` en pr-contract (DysTelefonica/team-skills#1786).

Cada entrada de `checks` admite `true` (bloquea), `false` (apagado), ausente
(activo, bloquea) y `"advise"`: la comprobación corre, sus hallazgos salen como
AVISO por stderr y el exit no cambia. Un valor distinto sale 2 (HR-3).

Stdlib exclusivamente. Ejecución: `python3 test_advise_mode.py`.
"""

import tempfile
import unittest

from test_pr_contract import AI_OK, make_event, make_policy, make_repo, run_gate

CHECK_NAMES = ("closes", "chain_context", "issue_links", "closure_tests",
               "docs_evidence", "budget", "commits", "ai_disclosure",
               "commit_identity", "deleted_tests")


def only(name, value):
    """`name` con `value`; el resto apagado. `ai_disclosure` queda activo como ancla:
    un gate con todo apagado sale 2 (HR-3), y esa comprobación solo avisa."""
    return {**{n: False for n in CHECK_NAMES}, "ai_disclosure": True, name: value}


class AdviseModeTests(unittest.TestCase):
    """#1786: `advise` por control. Black-box: el gate como subproceso."""

    def run_case(self, checks, *, body="sin nada\n", labels=("type:chore",),
                 commits=("feat: x",), **policy):
        with tempfile.TemporaryDirectory() as tmp:
            repo, base_sha, head_sha = make_repo(tmp, commits=list(commits))
            event = make_event(tmp, body=body, labels=list(labels),
                               base_sha=base_sha, head_sha=head_sha)
            overrides = {"checks": checks, **policy}
            return run_gate(event, make_policy(tmp, **overrides), repo)

    # --- closes (HR-7: contradicción Refs + cierre del mismo número) ---

    CONTRADICTION = "Closes #333\n\nRefs #333\n\n" + AI_OK

    def test_closes_advise_prints_the_contradiction_on_stderr_and_exits_0(self):
        proc = self.run_case(only("closes", "advise"), body=self.CONTRADICTION)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("AVISO", proc.stderr)
        self.assertIn("HR-7: el cuerpo dice `Refs #333`", proc.stderr)
        self.assertNotIn("FAIL", proc.stdout)

    def test_closes_true_still_blocks_the_contradiction(self):
        proc = self.run_case(only("closes", True), body=self.CONTRADICTION)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("HR-7: el cuerpo dice `Refs #333`", proc.stdout)

    def test_closes_false_is_off(self):
        proc = self.run_case(only("closes", False), body=self.CONTRADICTION)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertNotIn("Refs #333", proc.stdout + proc.stderr)

    # --- commits (HR-14: commit no convencional) ---

    NON_CONVENTIONAL = ("feat: x", "cambios sueltos")

    def test_commits_advise_prints_the_hr14_finding_on_stderr_and_exits_0(self):
        proc = self.run_case(only("commits", "advise"), commits=self.NON_CONVENTIONAL)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("AVISO", proc.stderr)
        self.assertIn("HR-14: commit no convencional", proc.stderr)

    def test_commits_true_still_blocks_the_hr14_finding(self):
        proc = self.run_case(only("commits", True), commits=self.NON_CONVENTIONAL)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("HR-14: commit no convencional", proc.stdout)

    def test_commits_false_is_off(self):
        proc = self.run_case(only("commits", False), commits=self.NON_CONVENTIONAL)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertNotIn("HR-14: commit no convencional", proc.stdout + proc.stderr)

    # --- issue_links (HR-54: enlace a una issue de otro repositorio) ---

    FOREIGN = "Closes #333\n\nRefs otro-org/otro-repo#12\n\n" + AI_OK

    def test_issue_links_advise_prints_the_hr54_finding_on_stderr_and_exits_0(self):
        proc = self.run_case(only("issue_links", "advise"), body=self.FOREIGN)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("AVISO", proc.stderr)
        self.assertIn("HR-54", proc.stderr)
        self.assertIn("otro-org/otro-repo#12", proc.stderr)

    def test_issue_links_true_still_blocks_the_hr54_finding(self):
        proc = self.run_case(only("issue_links", True), body=self.FOREIGN)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("HR-54", proc.stdout)

    def test_issue_links_false_is_off(self):
        proc = self.run_case(only("issue_links", False), body=self.FOREIGN)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertNotIn("HR-54", proc.stdout + proc.stderr)

    # --- budget (A1: presupuesto de líneas) ---

    BULK = ["feat: x"] + [f"feat: bulk {i}" for i in range(5)]

    def test_budget_advise_prints_the_a1_finding_on_stderr_and_exits_0(self):
        proc = self.run_case(only("budget", "advise"), commits=self.BULK,
                             review_budget_lines=3)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("AVISO", proc.stderr)
        self.assertIn("A1: el diff supera el presupuesto", proc.stderr)

    def test_budget_true_still_blocks_the_a1_finding(self):
        proc = self.run_case(only("budget", True), commits=self.BULK,
                             review_budget_lines=3)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("A1: el diff supera el presupuesto", proc.stdout)

    def test_budget_false_is_off(self):
        proc = self.run_case(only("budget", False), commits=self.BULK,
                             review_budget_lines=3)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertNotIn("A1", proc.stdout + proc.stderr)

    # --- valores inválidos: exit 2, nunca silencio ---

    def test_unknown_value_exits_2_naming_the_check(self):
        for value in ("sí", "Advise", "warn", "", 1):
            with self.subTest(value=value):
                proc = self.run_case(only("closes", value), body=self.CONTRADICTION)
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertIn("closes", proc.stderr)

    # --- aislamiento: un advise no rebaja a los controles bloqueantes ---

    def test_advise_on_one_check_does_not_soften_a_blocking_one(self):
        checks = {**only("budget", "advise"), "issue_links": True}
        proc = self.run_case(checks, body=self.FOREIGN, commits=self.BULK,
                             review_budget_lines=3)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("HR-54", proc.stdout)
        self.assertIn("AVISO", proc.stderr)
        self.assertIn("A1: el diff supera el presupuesto", proc.stderr)
        self.assertNotIn("A1: el diff supera el presupuesto", proc.stdout)


if __name__ == "__main__":
    unittest.main()
