#!/usr/bin/env python3
"""Suite del gate issue-spec (DysTelefonica/team-skills#633). stdlib exclusivamente.

Cada regla: caso positivo, negativo y degenerado. El modo local se ejecuta como
lo ejecuta preflight: sin payload, sin token ni red.

Retirado con la regla (DysTelefonica/team-skills#1093, R1 y R2): la etiqueta
`status:approved` de la issue, su comprobación de quién la puso (la lista
`approvers`), la etiqueta `type:*` de la issue y del PR, y las seis secciones
obligatorias del cuerpo de la issue. El gate comprueba que la PR enlace una
issue del repositorio.
"""

import importlib.machinery
import importlib.util
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

GATE_ENV = "CI_PATTERN_ISSUE_SPEC_PATH"


def resolve_gate(env):
    """Ruta del gate: override de entorno (consumer vendorizado) o layout del asset.

    Un consumer que copia el gate a otra ruta (p. ej. ``scripts/``) apunta
    ``CI_PATTERN_ISSUE_SPEC_PATH`` a su copia y adopta este test byte a byte.
    """
    override = env.get(GATE_ENV)
    if override:
        return Path(override)
    return Path(__file__).resolve().parents[1] / "check_issue_spec.py"


GATE = resolve_gate(os.environ)

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


def plain_issue(*labels):
    """Issue sin etiquetas: el gate ya no mira las etiquetas (R1)."""
    return {"number": 1, "labels": [{"name": n} for n in labels]}


class _ApiDownError(GitHubApiError):
    """Simulated GitHub outage: the remote check must fail closed."""


class _UnexpectedFetchError(AssertionError):
    """The API must not be queried when the body has no references."""


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

    def test_closes_and_refs_same_issue_passes_without_warning(self):
        # Refs redundante junto a Closes de la misma issue: pasa y no avisa.
        p = run_local("Closes #5\n\nRefs #5\n")
        self.assertEqual(p.returncode, 0, p.stderr)
        self.assertNotIn("aviso", p.stderr)

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
    """Regla 2 con la API simulada: la referencia resuelve a una issue, no a una PR."""

    def test_reference_to_pull_request_points_to_depends_on(self):
        def fetch(number):
            return {"number": number, "pull_request": {"url": "x"}}
        violations = validate_remote("Closes #634\n", "team/skills", fetch)
        self.assertEqual(len(violations), 1)
        self.assertIn("pull request", violations[0])
        self.assertIn("depends-on", violations[0])

    def test_issue_reference_passes(self):
        violations = validate_remote(
            "Closes #12\n", "team/skills", lambda n: plain_issue())
        self.assertEqual(violations, [])

    def test_issue_without_approval_passes(self):
        # R1 (DysTelefonica/team-skills#1093): la etiqueta de aprobación ya no
        # bloquea; el gate no mira las etiquetas de la issue.
        violations = validate_remote(
            "Closes #12\n", "team/skills", lambda n: plain_issue())
        self.assertEqual(violations, [])

    def test_api_error_fails_closed(self):
        def fetch(number):
            raise _ApiDownError()
        violations = validate_remote("Closes #12\n", "team/skills", fetch)
        self.assertEqual(len(violations), 1)
        self.assertIn("cannot read", violations[0])

    def test_code_reference_is_never_fetched(self):
        fetched = []

        def fetch(number):
            fetched.append(number)
            return plain_issue()
        validate_remote("`Refs #7`\n\nCloses #8\n", "team/skills", fetch)
        self.assertEqual(fetched, [8])

    def test_foreign_reference_is_not_fetched_from_this_repository(self):
        fetched = []

        def fetch(number):
            fetched.append(number)
            return plain_issue()
        violations = validate_remote("Closes other/repo#9\n", "team/skills", fetch)
        self.assertEqual(fetched, [])
        self.assertTrue(any("otro repositorio" in v for v in violations))

    def test_no_reference_never_calls_the_api(self):
        def fetch(number):
            raise _UnexpectedFetchError()
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


# --------------------------------------------------------------------------------------------
# #830: el cuerpo del PR se lee por API; el snapshot del payload queda obsoleto
# desde la primera edición del cuerpo y desde cualquier "Re-run jobs" del UI
# (ardelperal/APAP_WEB#2163, #2164). El gate nunca usa el cuerpo del payload
# en `pr-event`: lo trae de `GET /repos/{owner}/{repo}/pulls/{n}`. Sin token
# o con la API caída, falla cerrado con motivo explícito (nunca fallback
# silencioso al payload, HR-3).
# --------------------------------------------------------------------------------------------


def _plain_issue(number: int) -> dict:
    return {"number": number, "body": "", "labels": []}


def _pr_payload(body: str, *, pr_number: int = 7, repo: str = "team/skills",
                labels: list | None = None) -> dict:
    return {
        "repository": {"full_name": repo},
        "pull_request": {
            "number": pr_number,
            "body": body,
            "labels": [{"name": n} for n in (labels or ["type:bug"])],
        },
    }


class PrEventBodyFromApiTests(unittest.TestCase):
    """#830 — el cuerpo leído por API es la fuente; el del payload se ignora."""

    def test_payload_body_invalid_but_api_body_valid_passes(self):
        # El snapshot del payload es el cuerpo viejo sin enlace (lo dejó el
        # `opened`); la API devuelve el cuerpo corregido con `Closes #12`.
        # El gate pasa, demostrando que lee por API y no del payload.
        payload = _pr_payload("Cuerpo obsoleto sin enlace.")
        calls: list[tuple[str, int]] = []

        def body_fetcher(repo: str, number: int) -> dict:
            calls.append((repo, number))
            return {"number": number, "body": "Closes #12\n"}

        rc = check_issue_spec.pr_event(
            payload, body_fetcher=body_fetcher, issue_fetcher=_plain_issue)
        self.assertEqual(rc, 0)
        self.assertEqual(calls, [("team/skills", 7)])

    def test_api_body_invalid_fails_even_if_payload_carries_valid_body(self):
        # El `opened` dejó el cuerpo bueno en el payload; un commit posterior
        # dejó un cuerpo sin enlace y un re-run dispara este gate. La API gana:
        # el gate falla con el hallazgo del cuerpo actual.
        payload = _pr_payload("Closes #12\n")

        def body_fetcher(repo: str, number: int) -> dict:
            return {"number": number, "body": "Editar el cuerpo sin enlace a issue."}

        rc = check_issue_spec.pr_event(
            payload, body_fetcher=body_fetcher, issue_fetcher=_plain_issue)
        self.assertEqual(rc, 1)


class PrEventFailClosedTests(unittest.TestCase):
    """#830 — sin token o con la API caída: exit 2, nunca fallback al payload."""

    def _event_with_payload_body(self, tmp: Path, body: str) -> Path:
        event = tmp / "event.json"
        event.write_text(
            json.dumps(_pr_payload(body)), encoding="utf-8")
        return event

    def _run_pr_event(self, env: dict) -> subprocess.CompletedProcess:
        with tempfile.TemporaryDirectory() as tmp:
            event = self._event_with_payload_body(
                Path(tmp), "Closes #12\n")
            return subprocess.run(
                [sys.executable, str(GATE), "pr-event", str(event)],
                capture_output=True, text=True, check=False, env=env,
                timeout=60)

    def test_pr_event_without_token_fails_closed_with_explicit_reason(self):
        # Sin GITHUB_TOKEN, el gate sale con 2 y nombra la causa; el cuerpo
        # del payload NO aparece en stderr (no hay fallback silencioso).
        env = {k: v for k, v in os.environ.items() if k != "GITHUB_TOKEN"}
        p = self._run_pr_event(env)
        self.assertEqual(p.returncode, 2, p.stderr)
        self.assertIn("obligatorio", p.stderr)
        self.assertNotIn("Closes #12", p.stderr)

    def test_pr_event_api_outage_fails_closed_with_explicit_reason(self):
        # Token presente pero API caída (puerto cerrado): exit 2 con motivo
        # de la API; el cuerpo del payload NO se imprime.
        env = {k: v for k, v in os.environ.items() if k != "GITHUB_TOKEN"}
        env["GITHUB_TOKEN"] = "test-token"
        env["GITHUB_API_URL"] = "http://127.0.0.1:9"  # puerto cerrado
        p = self._run_pr_event(env)
        self.assertEqual(p.returncode, 2, p.stderr)
        self.assertIn("no se pudo leer", p.stderr)
        self.assertNotIn("Closes #12", p.stderr)


class GatePathResolutionTests(unittest.TestCase):
    """El gate se localiza por override de entorno o, sin él, por el layout del asset."""

    def test_default_is_sibling_of_tests_dir(self):
        self.assertEqual(
            resolve_gate({}),
            Path(__file__).resolve().parents[1] / "check_issue_spec.py")

    def test_env_override_wins(self):
        env = {GATE_ENV: "/somewhere/scripts/check_issue_spec.py"}
        self.assertEqual(
            resolve_gate(env), Path("/somewhere/scripts/check_issue_spec.py"))

    def test_empty_override_falls_back_to_default(self):
        self.assertEqual(resolve_gate({GATE_ENV: ""}), resolve_gate({}))


class ObsoleteFlagsCompatibilityTests(unittest.TestCase):
    """#1113 — las flags retiradas en #1103 se aceptan y se ignoran."""

    FIXTURE = Path(__file__).resolve().parent / "fixtures" / "pr_event_dependabot.json"

    @staticmethod
    def _run(event: Path, *extra: str) -> subprocess.CompletedProcess:
        env = {k: v for k, v in os.environ.items()
               if k not in ("GITHUB_TOKEN", "GITHUB_EVENT_PATH")}
        env["GITHUB_API_URL"] = "http://127.0.0.1:9"
        return subprocess.run(
            [sys.executable, str(GATE), "pr-event", str(event), *extra],
            capture_output=True, text=True, check=False, env=env, timeout=60)

    @staticmethod
    def _warning(flag: str) -> str:
        return (f"issue-spec: flag obsoleta --{flag}: se ignora; "
                "retirela del workflow")

    def test_required_sections_is_accepted_and_ignored(self):
        baseline = self._run(self.FIXTURE)
        self.assertEqual(baseline.returncode, 0, baseline.stderr)
        result = self._run(
            self.FIXTURE, "--required-sections", "required-sections.json")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout, baseline.stdout)
        self.assertIn(self._warning("required-sections"), result.stderr)

    def test_host_contract_is_accepted_and_ignored(self):
        baseline = self._run(self.FIXTURE)
        self.assertEqual(baseline.returncode, 0, baseline.stderr)
        result = self._run(
            self.FIXTURE, "--host-contract", "host-contract.json")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout, baseline.stdout)
        self.assertIn(self._warning("host-contract"), result.stderr)

    def test_both_obsolete_flags_are_accepted_and_ignored(self):
        baseline = self._run(self.FIXTURE)
        self.assertEqual(baseline.returncode, 0, baseline.stderr)
        result = self._run(
            self.FIXTURE,
            "--required-sections", "required-sections.json",
            "--host-contract", "host-contract.json")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout, baseline.stdout)
        self.assertIn(self._warning("required-sections"), result.stderr)
        self.assertIn(self._warning("host-contract"), result.stderr)

    def test_local_accepts_the_obsolete_required_sections_flag(self):
        """El preflight local de un consumer pasa la flag a `local`: tambien se acepta."""
        baseline = run_local("Closes #1")
        result = run_local("Closes #1", "--required-sections", "required-sections.json")
        self.assertEqual(baseline.returncode, 0, baseline.stderr)
        self.assertEqual(result.returncode, baseline.returncode, result.stderr)
        self.assertEqual(result.stdout, baseline.stdout)
        self.assertIn(self._warning("required-sections"), result.stderr)

    def test_obsolete_flag_warning_precedes_invalid_payload_verdict(self):
        with tempfile.TemporaryDirectory() as tmp:
            event = Path(tmp) / "invalid-event.json"
            event.write_text(
                '{"pull_request": {"number": 1}}', encoding="utf-8")
            baseline = self._run(event)
            result = self._run(
                event, "--host-contract", "host-contract.json")
        self.assertEqual(baseline.returncode, 2, baseline.stderr)
        self.assertEqual(result.returncode, baseline.returncode, result.stderr)
        self.assertEqual(result.stdout, baseline.stdout)
        self.assertIn(self._warning("host-contract"), result.stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)


DEPENDABOT_FIXTURE = Path(__file__).resolve().parent / "fixtures" / "pr_event_dependabot.json"


class PrEventDependabotTests(unittest.TestCase):
    """#1113 E — un PR de Dependabot no enlaza issue y no es un fallo del gate.

    El actor y la rama tienen que coincidir: el nombre de la rama por sí solo no
    prueba quién la empujó (mismo criterio que check_branch_name).
    """

    @staticmethod
    def _fixture() -> dict:
        return json.loads(DEPENDABOT_FIXTURE.read_text(encoding="utf-8"))

    @staticmethod
    def _body_without_issue(repo: str, number: int) -> dict:
        return {"number": number, "body": "Bumps lodash from 4.17.20 to 4.17.21."}

    @staticmethod
    def _no_issue_lookup(number: int) -> dict:
        raise _UnexpectedFetchError("un PR de Dependabot exento no consulta issues")

    def test_dependabot_pr_without_issue_link_is_exempt(self):
        rc = check_issue_spec.pr_event(
            self._fixture(), body_fetcher=self._body_without_issue,
            issue_fetcher=self._no_issue_lookup)
        self.assertEqual(rc, 0)

    def test_same_body_from_a_human_pr_still_fails(self):
        payload = self._fixture()
        payload["pull_request"]["user"] = {"login": "ardelperal", "type": "User"}
        payload["pull_request"]["head"]["ref"] = "fix/lodash"
        rc = check_issue_spec.pr_event(
            payload, body_fetcher=self._body_without_issue,
            issue_fetcher=lambda n: plain_issue())
        self.assertEqual(rc, 1)

    def test_dependabot_actor_on_a_human_branch_is_not_exempt(self):
        payload = self._fixture()
        payload["pull_request"]["head"]["ref"] = "feat/lodash"
        rc = check_issue_spec.pr_event(
            payload, body_fetcher=self._body_without_issue,
            issue_fetcher=lambda n: plain_issue())
        self.assertEqual(rc, 1)
