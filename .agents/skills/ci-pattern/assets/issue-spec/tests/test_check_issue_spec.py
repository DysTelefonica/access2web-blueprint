#!/usr/bin/env python3
"""Suite del gate issue-spec (DysTelefonica/team-skills#633). stdlib exclusivamente.

Cada regla: caso positivo, negativo y degenerado. El modo local se ejecuta como
lo ejecuta preflight: sin payload, sin token y sin red.
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
is_chain_partial = check_issue_spec.is_chain_partial
PrContext = check_issue_spec.PrContext
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
    """Reglas 2 y 3 con la API simulada: PR, aprobación, tipo, fallo de API (S3)."""

    def test_reference_to_pull_request_points_to_depends_on(self):
        def fetch(number):
            return {"number": number, "pull_request": {"url": "x"},
                    "labels": [{"name": "status:approved"}, {"name": "type:feature"}]}
        violations = validate_remote("Closes #634\n", "team/skills", fetch)
        self.assertEqual(len(violations), 1)
        self.assertIn("pull request", violations[0])
        self.assertIn("depends-on", violations[0])

    def test_pull_request_reference_is_a_warning_when_a_valid_issue_is_also_linked(self):
        # #787 S6: con otra issue valida enlazada, la referencia a una PR no bloquea.
        def fetch(number):
            if number == 634:
                return {"number": number, "pull_request": {"url": "x"}, "labels": []}
            return approved_issue("status:approved", "type:feature")
        context = PrContext()
        violations = validate_remote("Closes #12\nRefs #634\n", "team/skills", fetch,
                                     context=context)
        self.assertEqual(violations, [])
        self.assertEqual(len(context.warnings), 1)
        self.assertIn("pull request", context.warnings[0])

    def test_pull_request_reference_stays_a_violation_without_a_context(self):
        # Sin contexto nadie lee los avisos: la referencia a una PR no se oculta.
        def fetch(number):
            if number == 634:
                return {"number": number, "pull_request": {"url": "x"}, "labels": []}
            return approved_issue("status:approved", "type:feature")
        violations = validate_remote("Closes #12\nRefs #634\n", "team/skills", fetch)
        self.assertEqual(len(violations), 1)
        self.assertIn("pull request", violations[0])

    def test_chain_intermediate_does_not_need_the_issue_approved(self):
        # #787 S6: status:approved se exige en la punta de cadena, no en los eslabones.
        def fetch(number):
            return approved_issue("type:bug")
        self.assertEqual(
            validate_remote("Refs #12\n", "team/skills", fetch, context=PrContext(chain_partial=True)), [])
        self.assertEqual(
            validate_remote("Closes #12\n", "team/skills", fetch),
            ["#12: requiere status:approved"])

    def test_chain_intermediate_still_needs_a_type_label(self):
        violations = validate_remote(
            "Refs #12\n", "team/skills", lambda n: approved_issue(),
            context=PrContext(chain_partial=True))
        self.assertEqual(len(violations), 1)
        self.assertIn("type:*", violations[0])

    def test_chain_partial_is_read_from_the_pull_request_labels(self):
        self.assertTrue(is_chain_partial({"pull_request": {"labels": [{"name": "chain:partial"}]}}))
        self.assertFalse(is_chain_partial({"pull_request": {"labels": [{"name": "type:bug"}]}}))
        self.assertFalse(is_chain_partial({"pull_request": {}}))

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
            raise _ApiDownError()
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
            raise _UnexpectedFetchError()
        violations = validate_remote("sin enlace\n", "team/skills", fetch)
        self.assertEqual(len(violations), 1)
        self.assertIn("depends-on", violations[0])


# --------------------------------------------------------------------------------------------
# Regla 7 — secciones obligatorias (DysTelefonica/team-skills#644).

SECTIONS = (
    "Problema y contexto",
    "Evidencia verificable",
    "Alcance y no objetivos",
)


def _write_sections_config(tmp, payload):
    """Escribe un JSON de configuración de secciones en ``tmp`` y devuelve la ruta."""
    path = tmp / "required-sections.json"
    path.write_text(json.dumps(payload), encoding="utf-8")
    return path


class InformativeRefsTests(unittest.TestCase):
    """#913: las Refs son informativas; el PR debe enlazar al menos una issue aprobada."""

    @staticmethod
    def _issues(**by_number):
        # by_number: {"12": [labels...]}; el nombre de la clave es el número de issue.
        def fetch(number):
            return {"number": number,
                    "labels": [{"name": n} for n in by_number[str(number)]]}
        return fetch

    def test_closes_approved_with_non_approved_refs_passes(self):
        # (a) el cierre aprobado basta; la Refs no aprobada es informativa.
        fetch = self._issues(**{"12": ["status:approved", "type:bug"], "13": ["type:bug"]})
        violations = validate_remote("Closes #12\nRefs #13\n", "team/skills", fetch)
        self.assertEqual(violations, [])

    def test_only_refs_to_non_approved_issue_fails(self):
        # (b) sin ninguna issue aprobada el PR sigue en rojo.
        fetch = self._issues(**{"12": ["type:bug"]})
        violations = validate_remote("Refs #12\n", "team/skills", fetch)
        self.assertEqual(len(violations), 1)
        self.assertIn("aprobada", violations[0])

    def test_only_refs_to_approved_issue_still_passes(self):
        # no regresión: hoy pasa y debe seguir pasando (la Refs aprobada es la aprobación).
        fetch = self._issues(**{"12": ["status:approved", "type:bug"]})
        violations = validate_remote("Refs #12\n", "team/skills", fetch)
        self.assertEqual(violations, [])

    def test_closes_non_approved_with_approved_refs_fails(self):
        # (c) el issue que el PR cierra debe ser el aprobado: no basta una Refs aprobada.
        fetch = self._issues(**{"12": ["type:bug"], "13": ["status:approved", "type:bug"]})
        violations = validate_remote("Closes #12\nRefs #13\n", "team/skills", fetch)
        self.assertEqual(violations, ["#12: requiere status:approved"])

    def test_chain_partial_with_refs_to_canonical_approved_passes(self):
        # (d) eslabón intermedio: la aprobación vive en la punta, como hoy.
        fetch = self._issues(**{"12": ["status:approved", "type:bug"]})
        context = PrContext(chain_partial=True)
        violations = validate_remote("Refs #12\n", "team/skills", fetch, context=context)
        self.assertEqual(violations, [])

    def test_chain_partial_with_refs_to_non_approved_still_passes(self):
        # no regresión de #787 S6: en un eslabón intermedio no se exige aprobación.
        fetch = self._issues(**{"12": ["type:bug"]})
        context = PrContext(chain_partial=True)
        violations = validate_remote("Refs #12\n", "team/skills", fetch, context=context)
        self.assertEqual(violations, [])

    def test_informative_refs_keep_the_required_sections_rule(self):
        # no regresión (regla 7): las secciones obligatorias siguen aplicando a toda
        # issue enlazada, también a las Refs; solo cambia la exigencia de aprobación.
        def fetch(number):
            labels = [{"name": "status:approved"}, {"name": "type:bug"}] if number == 12 \
                else [{"name": "type:bug"}]
            return {"number": number, "labels": labels, "body": "sin secciones"}
        violations = validate_remote("Closes #12\nRefs #13\n", "team/skills", fetch,
                                     required_sections=("### Problema y contexto",))
        self.assertEqual(len(violations), 2)
        self.assertIn("#12:", violations[0])
        self.assertIn("#13:", violations[1])

    def test_api_error_on_informative_ref_fails_closed(self):
        # no regresión (regla 6): un error al leer una Refs sigue siendo hallazgo.
        def fetch(number):
            if number == 13:
                raise _ApiDownError()
            return approved_issue("status:approved", "type:bug")
        violations = validate_remote("Closes #12\nRefs #13\n", "team/skills", fetch)
        self.assertEqual(len(violations), 1)
        self.assertIn("cannot read", violations[0])


class RequiredSectionsLocalTests(unittest.TestCase):

    def test_disabled_unchanged_without_flag(self):
        # Sin --required-sections, la regla 7 no existe (backward compat).
        body = "Closes #42\n\n## Cualquier cosa\n\nTexto.\n"
        self.assertEqual(run_local(body).returncode, 0, run_local(body).stderr)

    def test_all_present_passes(self):
        body = (
            "Closes #42\n\n"
            f"## {SECTIONS[0]}\n\nTexto.\n\n"
            f"## {SECTIONS[1]}\n\nTexto.\n\n"
            f"## {SECTIONS[2]}\n\nTexto.\n"
        )
        with tempfile.TemporaryDirectory() as tmp:
            cfg = _write_sections_config(Path(tmp), {"required_sections": list(SECTIONS)})
            p = run_local(body, "--required-sections", str(cfg))
        self.assertEqual(p.returncode, 0, p.stderr)

    def test_one_missing_fails_with_section_name(self):
        body = (
            "Closes #42\n\n"
            f"## {SECTIONS[0]}\n\nTexto.\n\n"
            f"## {SECTIONS[1]}\n\nTexto.\n"
        )
        with tempfile.TemporaryDirectory() as tmp:
            cfg = _write_sections_config(Path(tmp), {"required_sections": list(SECTIONS)})
            p = run_local(body, "--required-sections", str(cfg))
        self.assertEqual(p.returncode, 1, p.stderr)
        self.assertIn(f"«{SECTIONS[2]}»", p.stderr)
        self.assertIn("obligatoria", p.stderr)

    def test_empty_section_fails(self):
        # Heading sin contenido bajo él: la regla exige heading + contenido
        # (mismo contrato del formulario: una textarea no se queda vacía).
        body = (
            "Closes #42\n\n"
            f"## {SECTIONS[0]}\n\nTexto.\n\n"
            f"## {SECTIONS[1]}\n\nTexto.\n\n"
            f"## {SECTIONS[2]}\n\n"
        )
        with tempfile.TemporaryDirectory() as tmp:
            cfg = _write_sections_config(Path(tmp), {"required_sections": list(SECTIONS)})
            p = run_local(body, "--required-sections", str(cfg))
        self.assertEqual(p.returncode, 1, p.stderr)
        self.assertIn(f"«{SECTIONS[2]}»", p.stderr)

    def test_all_missing_fails_with_one_violation_per_section(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg = _write_sections_config(Path(tmp), {"required_sections": list(SECTIONS)})
            p = run_local("Closes #42\n", "--required-sections", str(cfg))
        self.assertEqual(p.returncode, 1, p.stderr)
        for name in SECTIONS:
            self.assertIn(f"«{name}»", p.stderr)

    def test_issue_form_body_passes(self):
        # Cuerpo que sale de un issue form real (las 6 secciones de BP/CADETE).
        body = (
            "Closes #42\n\n"
            "## Problema y contexto\n\nEl formulario falla porque la rama está stale.\n\n"
            "## Evidencia verificable\n\n`gh run watch` muestra el job en rojo.\n\n"
            "## Alcance y no objetivos\n\nArreglar la invocación; nada más.\n\n"
            "## Criterios de aceptación\n\nLa suite canónica corre en verde.\n\n"
            "## Plan de validación\n\n`make verify` y la suite del asset.\n\n"
            "## Dependencias y riesgos\n\nNinguna externa; riesgo bajo.\n"
        )
        sections = [
            "Problema y contexto", "Evidencia verificable", "Alcance y no objetivos",
            "Criterios de aceptación", "Plan de validación", "Dependencias y riesgos",
        ]
        with tempfile.TemporaryDirectory() as tmp:
            cfg = _write_sections_config(Path(tmp), {"required_sections": sections})
            p = run_local(body, "--required-sections", str(cfg))
        self.assertEqual(p.returncode, 0, p.stderr)

class RequiredSectionsRemoteTests(unittest.TestCase):

    def _issue(self, body):
        return {
            "number": 7,
            "labels": [{"name": "status:approved"}, {"name": "type:feature"}],
            "body": body,
        }

    def test_remote_issue_with_all_sections_passes(self):
        body = (
            f"## {SECTIONS[0]}\nTexto.\n\n"
            f"## {SECTIONS[1]}\nTexto.\n\n"
            f"## {SECTIONS[2]}\nTexto.\n"
        )
        violations = validate_remote(
            "Closes #7\n", "team/skills", lambda n: self._issue(body), tuple(SECTIONS))
        self.assertEqual(violations, [])

    def test_remote_issue_missing_section_fails_with_number(self):
        body = f"## {SECTIONS[0]}\nTexto.\n\n## {SECTIONS[1]}\nTexto.\n"
        violations = validate_remote(
            "Closes #7\n", "team/skills", lambda n: self._issue(body), tuple(SECTIONS))
        self.assertEqual(len(violations), 1)
        self.assertIn("#7:", violations[0])
        self.assertIn(f"«{SECTIONS[2]}»", violations[0])


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


def _approved_issue(number: int) -> dict:
    return {"number": number, "body": "",
            "labels": [{"name": "status:approved"}, {"name": "type:bug"}]}


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
            payload, body_fetcher=body_fetcher, issue_fetcher=_approved_issue)
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
            payload, body_fetcher=body_fetcher, issue_fetcher=_approved_issue)
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


# --------------------------------------------------------------------------------------------
# Regla 3 (quién) — la contramarca de quién puede aprobar (DysTelefonica/team-skills#998).
#
# La lista sale de la clave opcional `approvers` del `.github/host-contract.json` del
# consumer. Aquí se fija de dónde se lee, que una declaración inservible sale 2 (HR-3) y
# que la puerta dice en una línea si el contrato declara aprobadores o si todavía no los
# declara (punto 3). El juicio de QUIÉN pone la etiqueta llega en el eslabón siguiente de
# la cadena, con la lectura de la línea de tiempo de la issue.
# --------------------------------------------------------------------------------------------

APPROVERS = ("ardelperal", "dys-orquestador[bot]")


def _write_host_contract(tmp, payload):
    """Escribe un `host-contract.json` en ``tmp`` y devuelve la ruta."""
    path = tmp / "host-contract.json"
    path.write_text(json.dumps(payload), encoding="utf-8")
    return path


def _run_pr_event_without_api(tmp, *extra):
    """Ejecuta `pr-event` con un payload que no llega a tocar la API.

    Un payload sin `repository.full_name` hace que el gate salga 2 como entrada inválida
    antes de cualquier llamada, así que lo observable aquí es exactamente lo que la puerta
    anuncia al leer el contrato: sin token, sin red y sin depender del veredicto.
    """
    event = Path(tmp) / "event.json"
    event.write_text('{"pull_request": {"number": 7}}', encoding="utf-8")
    env = {k: v for k, v in os.environ.items() if k != "GITHUB_TOKEN"}
    env["GITHUB_API_URL"] = "http://127.0.0.1:9"  # puerto cerrado
    return subprocess.run(
        [sys.executable, str(GATE), "pr-event", str(event), *extra],
        capture_output=True, text=True, check=False, env=env, timeout=60)


class LoadApproversTests(unittest.TestCase):
    """Punto 1: `approvers` es una clave opcional de `.github/host-contract.json`."""

    def test_absent_key_declares_nothing(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = _write_host_contract(Path(tmp), {"contract_version": 1})
            self.assertEqual(check_issue_spec.load_approvers(str(path)), ())

    def test_declared_list_is_read(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = _write_host_contract(
                Path(tmp), {"contract_version": 1, "approvers": list(APPROVERS)})
            self.assertEqual(check_issue_spec.load_approvers(str(path)), APPROVERS)

    def test_invalid_declarations_fail_closed(self):
        # Un contrato ilegible o inservible no puede dejar la puerta en verde (HR-3).
        cases = {
            "ilegible": None,
            "no json": "{",
            "no objeto": '["ardelperal"]',
            "no lista": '{"approvers": "ardelperal"}',
            "vacia": '{"approvers": []}',
            "entrada no textual": '{"approvers": ["ardelperal", 7]}',
            "entrada vacia": '{"approvers": [""]}',
        }
        for name, payload in cases.items():
            with self.subTest(name), tempfile.TemporaryDirectory() as tmp:
                path = Path(tmp) / "host-contract.json"
                path.write_text("{broken" if payload is None else payload, encoding="utf-8")
                with self.assertRaises(check_issue_spec.HostContractError):
                    check_issue_spec.load_approvers(str(path))

    def test_pr_event_exits_2_on_an_unserviceable_contract(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = _write_host_contract(Path(tmp), {"approvers": []})
            p = _run_pr_event_without_api(tmp, "--host-contract", str(contract))
        # 2 por el contrato inservible, antes incluso de leer el evento.
        self.assertEqual(p.returncode, 2, p.stderr)
        self.assertIn("approvers", p.stderr)
        self.assertNotIn("repository.full_name", p.stderr)


class ApproversReportedTests(unittest.TestCase):
    """Punto 3: la salida dice en una línea si el contrato declara aprobadores o no."""

    def test_declared_approvers_are_announced_and_not_warned_about(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = _write_host_contract(
                Path(tmp), {"contract_version": 1, "approvers": list(APPROVERS)})
            p = _run_pr_event_without_api(tmp, "--host-contract", str(contract))
        self.assertEqual(p.returncode, 2, p.stderr)  # el evento no identifica el repo
        self.assertIn("repository.full_name", p.stderr)
        self.assertIn("aprobadores declarados", p.stdout)
        self.assertNotIn("aviso", p.stderr)

    def test_contract_without_the_key_warns_once(self):
        with tempfile.TemporaryDirectory() as tmp:
            contract = _write_host_contract(Path(tmp), {"contract_version": 1})
            p = _run_pr_event_without_api(tmp, "--host-contract", str(contract))
        self.assertEqual(p.stderr.count("approvers"), 1, p.stderr)
        self.assertIn("no declara", p.stderr)

    def test_no_contract_flag_warns_too(self):
        # El consumer que aún no adopta la opción tampoco declara aprobadores: se dice
        # igual, porque su puerta sigue aceptando cualquier actor.
        with tempfile.TemporaryDirectory() as tmp:
            p = _run_pr_event_without_api(tmp)
        self.assertEqual(p.stderr.count("approvers"), 1, p.stderr)
        self.assertIn("no declara", p.stderr)


# --------------------------------------------------------------------------------------------
# Regla 3 (quién) — quién puso `status:approved` (DysTelefonica/team-skills#998).
#
# La etiqueta puesta ya no basta: con `approvers` declarado, el último evento `labeled` de
# la aprobación tiene que llevar un `actor.login` de la lista. Es el mismo patrón que
# `SIZE_EXCEPTION_APPROVER` en el gate de pr-size, con la identidad leída de la plataforma
# en vez de por una variable de entorno. El eslabón anterior deja la lista a punto; aquí se
# juzga.
# --------------------------------------------------------------------------------------------

APPROVERS = ("ardelperal", "dys-orquestador[bot]")


def _labeled(login, label="status:approved", at="2026-10-22T10:00:00Z"):
    """Evento `labeled` de la API de issues."""
    return {"event": "labeled", "label": {"name": label},
            "actor": {"login": login}, "created_at": at}


def _unlabeled(login, label="status:approved", at="2026-10-22T11:00:00Z"):
    return {"event": "unlabeled", "label": {"name": label},
            "actor": {"login": login}, "created_at": at}


def _remote(body, *, approvers=APPROVERS, events=(), issue=None, events_get=None):
    """`validate_remote` con la lista declarada y un doble de la línea de tiempo."""
    def by_number(number):
        return list(events)
    context = PrContext(approvers=approvers,
                        events_fetcher=events_get if events_get is not None else by_number)
    fetch = issue or (lambda n: approved_issue("status:approved", "type:feature"))
    return validate_remote(body, "team/skills", fetch, context=context)


class ApproverRuleTests(unittest.TestCase):
    """Puntos 2 y 6: el actor del último `labeled` tiene que ser aprobador declarado."""

    def test_declared_approver_passes(self):
        self.assertEqual(
            _remote("Closes #12\n", events=[_labeled("dys-orquestador[bot]")]), [])

    def test_actor_outside_the_list_fails_naming_actor_and_list(self):
        violations = _remote("Closes #12\n", events=[_labeled("dys-agentes[bot]")])
        self.assertEqual(len(violations), 1)
        self.assertIn("dys-agentes[bot]", violations[0])
        for approver in APPROVERS:
            self.assertIn(approver, violations[0])

    def test_label_without_the_event_fails_closed(self):
        # La etiqueta está puesta pero no consta el evento que la puso: no se puede probar
        # quién aprobó, y la puerta falla cerrada.
        violations = _remote("Closes #12\n", events=[])
        self.assertEqual(len(violations), 1)
        self.assertIn("no consta", violations[0])

    def test_last_labeled_event_after_a_removal_decides(self):
        # Un aprobador la puso, alguien la retiró y un actor no declarado la repuso: manda
        # el ÚLTIMO `labeled`, no el primero (punto 6, caso 5).
        violations = _remote("Closes #12\n", events=[
            _labeled("ardelperal"), _unlabeled("ardelperal"), _labeled("dys-agentes[bot]")])
        self.assertEqual(len(violations), 1)
        self.assertIn("dys-agentes[bot]", violations[0])

    def test_events_of_other_labels_are_not_the_approval(self):
        self.assertEqual(_remote("Closes #12\n", events=[
            _labeled("dys-agentes[bot]", label="type:bug"), _labeled("ardelperal")]), [])

    def test_login_comparison_is_case_insensitive(self):
        # GitHub no distingue mayúsculas en el login; el gate tampoco.
        self.assertEqual(_remote("Closes #12\n", events=[_labeled("Ardelperal")]), [])

    def test_api_error_reading_events_fails_closed(self):
        # Punto 5: error de API al leer los eventos = fallo cerrado, como el resto.
        def down(number):
            raise _ApiDownError()
        violations = _remote("Closes #12\n", events_get=down)
        self.assertEqual(len(violations), 1)
        self.assertIn("cannot read", violations[0])

    def test_missing_events_reader_fails_closed(self):
        # Sin cliente de eventos no se puede probar la separación: hallazgo, no verde.
        context = PrContext(approvers=APPROVERS)
        violations = validate_remote(
            "Closes #12\n", "team/skills",
            lambda n: approved_issue("status:approved", "type:feature"), context=context)
        self.assertEqual(len(violations), 1)
        self.assertIn("cannot read", violations[0])

    def test_declared_approvers_do_not_read_events_without_the_label(self):
        # Sin la etiqueta puesta la regla 3 ya falla por su ausencia; no se consulta la API.
        def never(number):
            raise _UnexpectedFetchError()
        violations = _remote(
            "Closes #12\n", events_get=never,
            issue=lambda n: {"number": n, "labels": [{"name": "type:feature"}]})
        self.assertEqual(violations, ["#12: requiere status:approved"])


class ApproverExemptionsTests(unittest.TestCase):
    """Punto 4 y #913: `chain:partial` exime, y las `Refs` no aceptan una aprobación ajena."""

    @staticmethod
    def _events_by_number(number):
        """Línea de tiempo por issue: #12 la aprobó el orquestador, #13 alguien sin autoridad."""
        actor = "dys-orquestador[bot]" if number == 12 else "dys-agentes[bot]"
        return [_labeled(actor)]

    def test_chain_partial_is_exempt_and_does_not_read_events(self):
        def never(number):
            raise _UnexpectedFetchError()
        context = PrContext(chain_partial=True, approvers=APPROVERS, events_fetcher=never)
        self.assertEqual(
            validate_remote("Refs #12\n", "team/skills",
                            lambda n: approved_issue("status:approved", "type:feature"),
                            context=context),
            [])

    def test_informative_ref_does_not_count_an_approval_from_a_non_approver(self):
        # Una `Refs` a una issue «aprobada» por un actor no declarado no da por buena la
        # aprobación del PR: sigue faltando la issue aprobada (mensaje de #913). Como una
        # `Refs` es informativa, el motivo va como aviso y no como hallazgo.
        context = PrContext(approvers=APPROVERS,
                            events_fetcher=lambda n: [_labeled("dys-agentes[bot]")])
        violations = validate_remote(
            "Refs #12\n", "team/skills",
            lambda n: approved_issue("status:approved", "type:feature"), context=context)
        self.assertEqual(len(violations), 1)
        self.assertIn("aprobada", violations[0])
        self.assertEqual(len(context.warnings), 1)
        self.assertIn("dys-agentes[bot]", context.warnings[0])

    def test_informative_ref_by_a_non_approver_warns_even_when_another_approves(self):
        # El aviso no bloquea: con otra issue aprobada de verdad, el PR pasa y aun asi se
        # dice que una de las enlazadas la aprobó alguien sin autoridad.
        context = PrContext(approvers=APPROVERS,
                            events_fetcher=self._events_by_number)
        violations = validate_remote(
            "Closes #12\nRefs #13\n", "team/skills",
            lambda n: approved_issue("status:approved", "type:feature"), context=context)
        self.assertEqual(violations, [])
        self.assertEqual(len(context.warnings), 1)
        self.assertIn("dys-agentes[bot]", context.warnings[0])

    def test_informative_ref_approved_by_a_declared_approver_passes(self):
        self.assertEqual(_remote("Refs #12\n", events=[_labeled("ardelperal")]), [])

    def test_without_the_declaration_the_verdict_is_unchanged(self):
        # Punto 3: sin `approvers` no se lee la API de eventos y el veredicto no cambia.
        def never(number):
            raise _UnexpectedFetchError()
        self.assertEqual(_remote("Closes #12\n", approvers=(), events_get=never), [])


class IssueEventsPaginationTests(unittest.TestCase):
    """El último evento `labeled` solo es fiable si se lee toda la línea de tiempo."""

    class _Client(check_issue_spec.GitHubClient):
        def __init__(self, pages):
            super().__init__()
            self.pages = pages
            self.requested = []

        def _events_page(self, repository, number, page):
            self.requested.append(page)
            return list(self.pages.get(page, []))

    def test_pages_are_walked_until_a_short_page(self):
        # La API devuelve 30 por defecto: una página llena no dice que sea la última, así
        # que el gate pide la siguiente hasta recibir una corta.
        full = [_labeled(f"a{i}") for i in range(100)]
        client = self._Client({1: full, 2: [_labeled("b")]})
        events = client.issue_events("team/skills", 12)
        self.assertEqual(len(events), 101)
        self.assertEqual(events[-1], _labeled("b"))
        self.assertEqual(client.requested, [1, 2])

    def test_a_short_first_page_ends_the_walk(self):
        # Con una página corta ya no hay nada más que leer: no se pide la siguiente.
        client = self._Client({1: [_labeled("a")]})
        self.assertEqual(client.issue_events("team/skills", 12), [_labeled("a")])
        self.assertEqual(client.requested, [1])

    def test_a_timeline_longer_than_the_cap_fails_closed(self):
        # Sin este techo una issue con miles de eventos devolvería un veredicto sobre un
        # fragmento del historial: falso verde. Ante la duda, hallazgo (HR-3).
        full = [_labeled("a")] * 100
        client = self._Client({page: full for page in range(1, 20)})
        with self.assertRaises(check_issue_spec.GitHubApiError):
            client.issue_events("team/skills", 12)
        self.assertEqual(client.requested, list(range(1, 11)))


if __name__ == "__main__":
    unittest.main(verbosity=2)
