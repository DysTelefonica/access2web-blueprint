#!/usr/bin/env python3
"""#1613: HR-7 sobre cada palabra de cierre con número del cuerpo del PR.

GitHub cierra la issue que sigue a `close[sd]?`, `fix(es|ed)?` o `resolve[sd]?`
en cualquier parte del cuerpo e ignora la negación que la precede
(APAP_WEB#4075 cerró #4039 con «It does not close #4039»). El gate avisa de
cada una, sin cambiar el veredicto; y BLOQUEA cuando el mismo `#N` aparece con
`Refs #N` y con una palabra de cierre (team-skills#1614 cerró #1592). stdlib.
"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

GATE = Path(__file__).resolve().parents[1] / "check_pr_contract.py"
REPO = "DysTelefonica/example"
OFF = ("chain_context", "issue_links", "closure_tests", "docs_evidence", "budget",
       "commits", "ai_disclosure", "commit_identity", "deleted_tests")
STDERR = f"PR-CONTRACT: apagadas por política: {', '.join(OFF)}\n"
OK = "PR-CONTRACT OK: PR #7 (opened) sin hallazgos.\n"
# Línea real de team-skills#1614 (cuerpo, línea 3), que cerró #1592 al fusionarse.
LINE_1614 = (
    "PR 1 of 3 for #1592: the retry rule. It replaces #1607 (915 lines), which is "
    "being split into PRs off main under 400 lines each (owner, via apap-reviewer: "
    "no `size:exception`). PR-2 (governed-merge write retries) and PR-3 "
    "(`check_issue_spec`) come once this one lands; the last one closes #1592.")


def advisory(line, target):
    return (f"AVISO (no bloquea): HR-7: la línea «{line}» cierra {target} al fusionar "
            "el PR. GitHub ignora la negación delante de la palabra de cierre; si no "
            f"debe cerrarla, escriba `Refs {target}`.\n")


def run_gate(body, labels=(), base="main"):
    """Gate con solo `closes` activo; el modo vivo (API) queda fuera."""
    with tempfile.TemporaryDirectory() as tmp:
        event = Path(tmp) / "event.json"
        event.write_text(json.dumps({
            "action": "opened",
            "pull_request": {"number": 7, "base": {"ref": base}, "body": body,
                             "labels": [{"name": name} for name in labels]},
            "repository": {"full_name": REPO, "default_branch": "main"},
        }), encoding="utf-8")
        policy = Path(tmp) / "policy.json"
        policy.write_text(json.dumps({
            "chain_tip_requires_closes": False,
            "checks": {"closes": True, **dict.fromkeys(OFF, False)},
        }), encoding="utf-8")
        env = {k: v for k, v in os.environ.items() if k not in (
            "GITHUB_EVENT_PATH", "PR_CONTRACT_USE_LIVE_API", "GH_TOKEN", "GITHUB_TOKEN")}
        env["PYTHONUTF8"] = "1"
        return subprocess.run(
            [sys.executable, str(GATE), "--event-file", str(event), "--policy-file", str(policy)],
            capture_output=True, text=True, encoding="utf-8", env=env, check=False)


class ClosingMentionAdvisoryTests(unittest.TestCase):
    def assertGate(self, body, stdout, code=0):
        done = run_gate(body)
        self.assertEqual(done.stdout, stdout)
        self.assertEqual(done.stderr, STDERR)
        self.assertEqual(done.returncode, code)

    def test_a_negated_close_is_advised_and_names_the_issue(self):
        line = "It does not close #12"
        self.assertGate(f"{line}\n", advisory(line, "#12") + OK)

    def test_a_plain_closes_on_the_tip_is_advised_and_still_passes(self):
        self.assertGate("Closes #12\n", advisory("Closes #12", "#12") + OK)

    def test_closed_fixes_and_resolves_are_advised(self):
        for line in ("closed #12", "FIXES #12", "Fixed: #12", "resolves #12"):
            with self.subTest(line=line):
                self.assertGate(f"{line}\n", advisory(line, "#12") + OK)

    def test_owner_repo_and_url_forms_are_advised(self):
        for line in ("Closes DysTelefonica/example#12",
                     "resolves https://github.com/DysTelefonica/example/issues/12"):
            with self.subTest(line=line):
                self.assertGate(f"{line}\n", advisory(line, "DysTelefonica/example#12") + OK)

    def test_a_clean_refs_body_passes_without_advisory(self):
        self.assertGate("Refs #12\n\nCuerpo sin cierre.\n", OK)

    def test_a_keyword_without_number_is_not_advised(self):
        self.assertGate("Refs #12\n\nThis closes the gap and fixes the race.\n", OK)

    def test_refs_and_a_closing_word_on_another_number_only_advise(self):
        self.assertGate("Refs #12\n\nCloses #13\n", advisory("Closes #13", "#13") + OK)


class RefsAndClosingContradictionTests(unittest.TestCase):
    def test_the_1614_body_fails_and_quotes_the_line(self):
        done = run_gate(f"Refs #1592\n\n{LINE_1614}\n\n- Changelog fragment.\n")

        self.assertEqual(done.stdout, (
            "FAIL: contrato de PR con 1 hallazgo(s):\n"
            f"  ✗ HR-7: el cuerpo dice `Refs #1592` y la línea «{LINE_1614}» cierra #1592 "
            "al fusionar el PR: es contradictorio. Quite la palabra de cierre o el "
            "`Refs #1592`.\n"))
        self.assertEqual(done.stderr, STDERR)
        self.assertEqual(done.returncode, 1)

    def test_every_keyword_variant_on_a_referenced_issue_fails(self):
        for keyword in ("close", "closes", "closed", "fix", "fixes", "fixed",
                        "resolve", "resolves", "resolved"):
            line = f"{keyword} #12"
            with self.subTest(keyword=keyword):
                done = run_gate(f"Refs #12\n\n{line}\n")
                self.assertEqual(done.stdout, (
                    "FAIL: contrato de PR con 1 hallazgo(s):\n"
                    f"  ✗ HR-7: el cuerpo dice `Refs #12` y la línea «{line}» cierra #12 "
                    "al fusionar el PR: es contradictorio. Quite la palabra de cierre o el "
                    "`Refs #12`.\n"))
                self.assertEqual(done.returncode, 1)

    def test_a_foreign_repository_with_the_same_number_only_warns(self):
        line = "Closes DysTelefonica/otro-repo#12"
        done = run_gate(f"Refs #12\n\n{line}\n")
        self.assertEqual(done.stdout, advisory(line, "DysTelefonica/otro-repo#12") + OK)
        self.assertEqual(done.returncode, 0)

    def test_a_negated_close_of_a_referenced_issue_fails(self):
        done = run_gate("Refs #4039\n\nIt does not close #4039\n")

        self.assertEqual(done.stdout, (
            "FAIL: contrato de PR con 1 hallazgo(s):\n"
            "  ✗ HR-7: el cuerpo dice `Refs #4039` y la línea «It does not close #4039» "
            "cierra #4039 al fusionar el PR: es contradictorio. Quite la palabra de "
            "cierre o el `Refs #4039`.\n"))
        self.assertEqual(done.stderr, STDERR)
        self.assertEqual(done.returncode, 1)


if __name__ == "__main__":
    unittest.main()
