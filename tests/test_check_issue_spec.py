#!/usr/bin/env python3
"""Suite del gate issue-spec (DysTelefonica/team-skills#633). stdlib exclusivamente.

Cada regla: caso positivo, negativo y degenerado. El modo local se ejecuta como
lo ejecuta preflight: sin payload, sin token y sin red.
"""

import importlib.machinery
import importlib.util
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

# Path adjustment for BP (Refs #960): in team-skills the gate lives at
# assets/issue-spec/check_issue_spec.py; here it lives at scripts/.
GATE = Path(__file__).resolve().parents[1] / "scripts" / "check_issue_spec.py"

_spec = importlib.util.spec_from_loader(
    "check_issue_spec",
    importlib.machinery.SourceFileLoader("check_issue_spec", str(GATE)))
check_issue_spec = importlib.util.module_from_spec(_spec)
# dataclass busca su módulo en sys.modules: hay que registrarlo antes de ejecutarlo.
sys.modules["check_issue_spec"] = check_issue_spec
_spec.loader.exec_module(check_issue_spec)
validate_body = check_issue_spec.validate_body
validate_remote = check_issue_spec.validate_remote
references = check_issue_spec.references
GitHubApiError = check_issue_spec.GitHubApiError


def run_local(body, *extra, stdin=False):
    """Ejecuta `local` como preflight: entorno sin token ni payload."""
    env = {k: v for k, v in os.environ.items()
           if k not in ("GITHUB_TOKEN", "GITHUB_EVENT_PATH")}
    env["GITHUB_API_URL"] = "http://127.0.0.1:9"  # cualquier llamada fallaría
    if stdin:
        cmd = [sys.executable, str(GATE), "local", "--body-file", "-", *extra]
        return subprocess.run(cmd, input=body, capture_output=True, text=True,
                              check=False, env=env)
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "body.md"
        path.write_text(body, encoding="utf-8")
        cmd = [sys.executable, str(GATE), "local", "--body-file", str(path), *extra]
        return subprocess.run(cmd, capture_output=True, text=True, check=False, env=env)


def approved_issue(*labels):
    return {"number": 1, "labels": [{"name": n} for n in labels]}


class ReferenceFormTests(unittest.TestCase):
    """Regla 1 y reglas de forma: qué cuenta como referencia (S1, S2)."""

    def test_closes_issue_passes_locally(self):
        p = run_local("Resumen.\n\nCloses #12\n")
        self.assertEqual(p.returncode, 0, p.stderr)

    def test_no_reference_fails_and_points_to_depends_on(self):
        # S3: el mensaje manda a depends-on, no solo dice que falta un enlace.
        p = run_local("Sin enlace a ninguna issue.\n")
        self.assertEqual(p.returncode, 1)
        self.assertIn("depends-on", p.stderr)

    def test_closes_and_refs_same_issue_fails(self):
        p = run_local("Closes #12\n\nRefs #12\n")
        self.assertEqual(p.returncode, 1)
        self.assertIn("MISMA issue", p.stderr)

    def test_closes_and_refs_different_issues_pass(self):
        p = run_local("Closes #12\nRefs #13\n")
        self.assertEqual(p.returncode, 0, p.stderr)

    def test_closing_keyword_in_prose_counts_like_github(self):
        self.assertEqual(
            [(r.kind, r.number) for r in references("Este cambio fixes #12 de raíz.")],
            [("close", 12)])

    def test_all_github_closing_keyword_forms_count(self):
        for word in ("Close", "closed", "Fix", "fixed", "Resolves", "resolved"):
            with self.subTest(word=word):
                self.assertEqual(
                    [r.kind for r in references(f"{word} #7")], ["close"])

    def test_refs_inside_inline_code_is_not_a_reference(self):
        p = run_local("Ejemplo: `Refs #12`, nada más.\n")
        self.assertEqual(p.returncode, 1)
        self.assertIn("depends-on", p.stderr)

    def test_closes_inside_double_backticks_is_not_a_reference(self):
        self.assertEqual(references("Usa ``Closes #12`` como ejemplo."), ())

    def test_keyword_split_by_code_span_is_not_a_reference(self):
        self.assertEqual(references("Closes `x` #12"), ())

    def test_fenced_block_is_not_a_reference_even_with_conflict_outside(self):
        body = "```md\nCloses #12\nRefs #12\n```\n\nRefs #12\n"
        self.assertEqual([(r.kind, r.number) for r in references(body)], [("refs", 12)])
        self.assertEqual(validate_body(body), [])

    def test_tilde_fence_is_not_a_reference(self):
        self.assertEqual(references("~~~\nCloses #12\n~~~\n"), ())

    def test_unclosed_fence_hides_the_rest_of_the_body(self):
        # CommonMark: un bloque sin cerrar llega hasta el final.
        self.assertEqual(references("```\nlo que sigue\nCloses #12\n"), ())

    def test_prefix_word_is_not_a_keyword(self):
        self.assertEqual(references("Prefs #12 y unrefs #13"), ())

    def test_cross_repository_reference_fails_when_repo_is_known(self):
        p = run_local("Closes other/repo#12\n", "--repository", "team/skills")
        self.assertEqual(p.returncode, 1)
        self.assertIn("otro repositorio", p.stderr)

    def test_same_repository_reference_passes_when_repo_is_known(self):
        p = run_local("Closes team/skills#12\n", "--repository", "Team/Skills")
        self.assertEqual(p.returncode, 0, p.stderr)

    def test_stdin_body_is_read(self):
        p = run_local("Refs #12\n", stdin=True)
        self.assertEqual(p.returncode, 0, p.stderr)

    def test_missing_body_file_is_invalid_input(self):
        env = {k: v for k, v in os.environ.items() if k != "GITHUB_TOKEN"}
        p = subprocess.run(
            [sys.executable, str(GATE), "local", "--body-file", "/no/such/body.md"],
            capture_output=True, text=True, check=False, env=env)
        self.assertEqual(p.returncode, 2, p.stderr)

    def test_local_mode_has_no_network_dependency(self):
        # run_local apunta GITHUB_API_URL a un puerto cerrado: si el modo local
        # llamara a la API, fallaría con exit 1 por API en vez de 0.
        p = run_local("Closes #12\n")
        self.assertEqual(p.returncode, 0, p.stderr)
        self.assertNotIn("cannot read", p.stderr)


class RemoteRuleTests(unittest.TestCase):
    """Reglas 2 y 3 con la API simulada: PR, aprobación, tipo, fallo de API (S3)."""

    def test_reference_to_pull_request_points_to_depends_on(self):
        def fetch(number):
            return {"number": number, "pull_request": {"url": "x"},
                    "labels": [{"name": "status:approved"}, {"name": "type:feature"}]}
        violations = validate_remote("Closes #634\n", "team/skills", fetch)
        self.assertEqual(len(violations), 1)
        self.assertIn("pull request", violations[0])
        self.assertIn("depends-on", violations[0])

    def test_approved_typed_issue_passes(self):
        violations = validate_remote(
            "Closes #12\n", "team/skills",
            lambda n: approved_issue("status:approved", "type:feature"))
        self.assertEqual(violations, [])

    def test_missing_approval_fails(self):
        violations = validate_remote(
            "Closes #12\n", "team/skills", lambda n: approved_issue("type:bug"))
        self.assertEqual(violations, ["#12: requiere status:approved"])

    def test_missing_type_label_fails(self):
        violations = validate_remote(
            "Closes #12\n", "team/skills", lambda n: approved_issue("status:approved"))
        self.assertEqual(len(violations), 1)
        self.assertIn("type:*", violations[0])

    def test_api_error_fails_closed(self):
        def fetch(number):
            raise GitHubApiError("GET failed")
        violations = validate_remote("Closes #12\n", "team/skills", fetch)
        self.assertEqual(len(violations), 1)
        self.assertIn("cannot read", violations[0])

    def test_code_reference_is_never_fetched(self):
        fetched = []

        def fetch(number):
            fetched.append(number)
            return approved_issue("status:approved", "type:bug")
        validate_remote("`Refs #7`\n\nCloses #8\n", "team/skills", fetch)
        self.assertEqual(fetched, [8])

    def test_foreign_reference_is_not_fetched_from_this_repository(self):
        fetched = []

        def fetch(number):
            fetched.append(number)
            return approved_issue("status:approved", "type:bug")
        violations = validate_remote("Closes other/repo#9\n", "team/skills", fetch)
        self.assertEqual(fetched, [])
        self.assertTrue(any("otro repositorio" in v for v in violations))

    def test_no_reference_never_calls_the_api(self):
        def fetch(number):
            raise AssertionError("no debe consultar la API sin referencias")
        violations = validate_remote("sin enlace\n", "team/skills", fetch)
        self.assertEqual(len(violations), 1)
        self.assertIn("depends-on", violations[0])


class PrEventInvalidInputTests(unittest.TestCase):
    """pr-event sin repositorio identificable es entrada inválida (exit 2)."""

    def test_event_without_repository_is_invalid_input(self):
        with tempfile.TemporaryDirectory() as tmp:
            event = Path(tmp) / "event.json"
            event.write_text('{"pull_request": {"body": "Closes #1"}}', encoding="utf-8")
            p = subprocess.run([sys.executable, str(GATE), "pr-event", str(event)],
                               capture_output=True, text=True, check=False)
        self.assertEqual(p.returncode, 2, p.stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
