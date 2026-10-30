#!/usr/bin/env python3
# ci-pattern asset tests — forma canónica del job de la fase 3 (HR-72, #1027)
"""El fragmento de workflow que el porting-guide publica para el job
`governance-seals` es exactamente el que los gates aceptan.

La forma canónica vive en el porting-guide (es lo que lee quien adopta), no en
una copia de este test: el test la extrae del documento y la mide. Así la prosa
y el gate no pueden divergir sin que se caiga una prueba.

Tres medidas:

1. El workflow resultante pasa `check_workflow_policy.py` del propio canon, con
   la política real de jobs requeridos de este repositorio —el mismo gate y el
   mismo política con los que se mide un consumer, no uno inventado aquí.
2. La fase 3 de `adoption check` acepta el contrato que el documento publica:
   con el job cableado y en `required_jobs` sale 0, y sin declararlo sigue
   saliendo 0 (fail-open, HR-72).
3. Lo que el gate de política no puede saber: el job invoca la CLI como comando
   de primer nivel —no encadenada a otro gate, que es el fork bomb de HR-46—,
   publica su nombre de check solo desde `pull_request` (HR-30) y no usa
   `|| true`, que dejaría el gate en verde sin haber medido.

Y una cuarta medida, la de #1062: el fragmento no declara el commit del catálogo
por escrito. El sello `catalog @ <sha40>` de `.team-skills.yaml` es la fuente
única, y el detector que lo vigila se demuestra con un caso negativo real.

Ejecuta sin red y sin consumer.
"""
import importlib
import importlib.machinery
import importlib.util
import json
import re
import subprocess
import sys
import tempfile
import textwrap
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
SKILL_ROOT = ASSET.parents[1]
REPO = ASSET.parents[4]
PORTING_GUIDE = SKILL_ROOT / "references" / "porting-guide.md"
ADOPTION_CONTRACT = SKILL_ROOT / "assets" / "adoption" / "adoption-contract.schema.json"
CLI_PATH = SKILL_ROOT / "assets" / "bin" / "ci-pattern"
WORKFLOW_POLICY = REPO / ".github" / "workflow-policy.json"
GATE = SKILL_ROOT / "assets" / "workflow-policy" / "check_workflow_policy.py"

if str(ASSET.parent / "workflow-policy") not in sys.path:
    sys.path.insert(0, str(ASSET.parent / "workflow-policy"))
policy_gate = importlib.import_module("check_workflow_policy")


def _load_cli():
    spec = importlib.util.spec_from_loader(
        "ci_pattern_canonical_job",
        importlib.machinery.SourceFileLoader("ci_pattern_canonical_job", str(CLI_PATH)),
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


CLI = _load_cli()

JOB = "governance-seals"
SHA40 = re.compile(r"(?<![0-9a-fA-F])[0-9a-f]{40}(?![0-9a-fA-F])")
USES_REF = re.compile(r"\buses:\s*[A-Za-z0-9._/-]+@[0-9a-f]{40}")
#: El pin que el fragmento publicaba escrito a mano, reutilizado como fixture del
#: caso negativo: si el detector no lo ve, no está mirando lo que dice mirar.
LITERAL_PIN = "bc915526af7920e459afd2532e93e96e0479c0e6"
# Esqueleto mínimo de un `ci.yml` de consumer con el agregador `required`: el
# fragmento del porting-guide se inyecta como job más, con su entrada en el
# `needs` del agregador, que es como lo cablea un adopter.


def canonical_job() -> str:
    """El workflow que publica el porting-guide, extraído del documento.

    Desde #1270 el bloque es un workflow completo en `pull_request_target`; desde
    #1396 su único job es `governance-seals`, que toma el patrón de la copia
    vendorizada en la BASE. Se busca el
    bloque cercado que contiene la clave `governance-seals:`: si el documento
    deja de publicarlo, esta función reventa y el fallo nombra el documento.
    """
    text = PORTING_GUIDE.read_text(encoding="utf-8")
    for block in re.findall(r"```yaml\n(.*?)```", text, re.S):
        if f"{JOB}:" in block:
            # El bloque vive dentro de una lista numerada: `dedent` le quita su sangria.
            return textwrap.dedent(block).rstrip("\n")
    raise AssertionError(
        f"el porting-guide ya no publica el workflow con el job `{JOB}`: "
        f"la forma canónica de HR-72 desapareció de {PORTING_GUIDE}"
    )


def hand_written_shas(text: str) -> list[str]:
    """Los sha de 40 hex escritos a mano, fuera del ref de una action (`uses:`)."""
    return [f"linea {n}: {m.group(0)}"
            for n, line in enumerate(text.splitlines(), 1)
            for m in SHA40.finditer(USES_REF.sub("", line))]


def consumer_workflow() -> str:
    """El workflow publicado, tal cual lo pega el consumer."""
    return canonical_job()


PR_CONTRACT_TEMPLATE = SKILL_ROOT / "assets" / "templates" / ".github" / "workflows" / "pr-contract.yml.tmpl"
TRIGGER_TYPES = re.compile(r"^\s*types:\s*\[(.*?)\]", re.M)
CONCURRENCY_GROUP = re.compile(r"^\s*group:\s*(\S+)", re.M)


def trigger_types(workflow: str) -> list[str]:
    """Los `types` del disparador de PR de un workflow."""
    match = TRIGGER_TYPES.search(workflow)
    assert match, "el workflow no declara `types` en su disparador"
    return [t.strip() for t in match.group(1).split(",")]


def concurrency_group(workflow: str) -> str:
    match = CONCURRENCY_GROUP.search(workflow)
    assert match, "el workflow no declara grupo de `concurrency`"
    return match.group(1)


class TreeOnlyTriggerTests(unittest.TestCase):
    """#1440 (C2): un job que solo mide el arbol no corre al editar el cuerpo o las etiquetas."""

    def test_the_fragment_runs_only_when_the_tree_can_change(self):
        self.assertEqual(trigger_types(consumer_workflow()), ["opened", "synchronize", "reopened"])

    def test_the_fragment_does_not_share_the_pr_contract_queue(self):
        template = PR_CONTRACT_TEMPLATE.read_text(encoding="utf-8")
        self.assertNotEqual(concurrency_group(consumer_workflow()), concurrency_group(template))


class CanonicalJobTests(unittest.TestCase):
    def test_fragment_passes_the_workflow_policy_gate(self):
        """La forma publicada pasa el gate de política."""
        with tempfile.TemporaryDirectory() as tmp:
            rendered = Path(tmp) / "wf.yml"
            rendered.write_text(consumer_workflow(), encoding="utf-8")
            policy = Path(tmp) / "policy.json"
            policy.write_text(json.dumps({"required_jobs": [JOB]}), encoding="utf-8")
            proc = subprocess.run([sys.executable, str(GATE), "--workflow", str(rendered),
                                   "--required-jobs", str(policy), "--base-branch", "main"],
                                  capture_output=True, text=True, cwd=REPO)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("WORKFLOW-POLICY OK", proc.stdout)

    def test_fragment_is_the_job_the_schema_declares(self):
        """La clave que publica el porting-guide existe en el esquema y pide job."""
        props = json.loads(ADOPTION_CONTRACT.read_text(encoding="utf-8"))["properties"]
        self.assertIn("governance_seals", props)
        self.assertEqual(props["governance_seals"]["required"], ["job"])
        self.assertIn('"job": "governance-seals"', PORTING_GUIDE.read_text(encoding="utf-8"))

    def test_phase_three_accepts_the_published_form(self):
        """Con el job cableado y en `required_jobs`, la fase 3 sale 0."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp)
            (repo / ".github" / "workflows").mkdir(parents=True)
            (repo / ".github" / "workflows" / "ci.yml").write_text(
                consumer_workflow(), encoding="utf-8")
            (repo / ".github" / "required-jobs-policy.json").write_text(
                json.dumps({"required_jobs": ["unit", JOB]}), encoding="utf-8")
            contract = {"governance_seals": {"job": JOB}}
            self.assertEqual(CLI._check_governance_seals(repo, contract), [])

    def test_phase_three_without_the_declaration_is_still_green(self):
        """Fail-open: un consumer que no declara conserva su adopción previa."""
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp)
            self.assertEqual(CLI._check_governance_seals(repo, {}), [])

    def test_job_publishes_its_check_only_from_pull_request(self):
        """HR-30: el workflow evalúa el PR y solo en ese evento."""
        workflow = consumer_workflow()
        events = policy_gate.workflow_events(policy_gate.parse_workflow(workflow, "wf.yml"))
        self.assertEqual(sorted(events), ["pull_request_target"], events)

    def test_job_invokes_the_cli_as_a_top_level_command(self):
        """HR-46: la CLI es el comando del paso, no el argumento de otro gate."""
        tree = policy_gate.parse_workflow(consumer_workflow(), "ci.yml")
        commands = "\n".join(
            str(step.get("run") or "") for step in tree["jobs"][JOB]["steps"])
        self.assertIn("adoption check --phase 3", commands)
        # Un gate que invoca la CLI como subproceso es el ciclo que produjo
        # 1.300 procesos y el OOM del 2026-10-05: ni `ci-pattern verify` ni otro
        # script del patrón pueden aparecer en la línea de comando.
        self.assertNotIn("verify", commands)
        for line in commands.splitlines():
            if "adoption check" in line:
                self.assertNotRegex(line, r"check_[a-z_]+\.py|ci-pattern verify")

    def test_job_does_not_swallow_the_verdict(self):
        """Un `|| true` en el paso dejaría el job verde sin haber medido."""
        tree = policy_gate.parse_workflow(consumer_workflow(), "ci.yml")
        commands = "\n".join(
            str(step.get("run") or "") for step in tree["jobs"][JOB]["steps"])
        self.assertNotIn("|| true", commands)
        self.assertNotIn("continue-on-error", str(tree["jobs"][JOB]))

    # --- #1062: el commit del catálogo se lee del sello, no se escribe ---

    def test_fragment_does_not_declare_the_catalog_commit(self):
        """El fragmento no lleva el sha del catálogo escrito a mano."""
        self.assertEqual(hand_written_shas(canonical_job()), [])

    def test_the_detector_sees_a_written_pin(self):
        """Caso negativo: el mismo detector sobre un `pin:` literal."""
        found = hand_written_shas(f"           with:\n             pin: {LITERAL_PIN}\n")
        self.assertEqual(len(found), 1, found)
        self.assertTrue(found[0].endswith(LITERAL_PIN), found)

    def test_fragment_reads_the_commit_from_the_manifest_seal(self):
        """El commit del catálogo es el de la copia vendorizada en la base (#1062, #1396)."""
        tree = policy_gate.parse_workflow(consumer_workflow(), "ci.yml")
        steps = [s for s in tree["jobs"][JOB]["steps"] if "GITHUB_OUTPUT" in str(s.get("run") or "")]
        self.assertEqual(steps, [], "el job no resuelve el sello: lo fija la propagación en la base")
        self.assertEqual(tree["jobs"][JOB]["env"]["PATTERN_DIR"], "base/.agents/skills/ci-pattern")

    # --- #1396: un solo job, sin token, sin artefacto de entrada; el espejo se queda ---

    def test_the_fragment_has_no_canon_job_nor_token(self):
        tree = policy_gate.parse_workflow(consumer_workflow(), "ci.yml")
        self.assertEqual(list(tree["jobs"]), [JOB])
        self.assertNotIn("needs", tree["jobs"][JOB])
        self.assertNotIn("TEAM_SKILLS", consumer_workflow())
        self.assertNotIn("pattern-at-pin", consumer_workflow())

    def test_the_verdict_does_not_depend_on_an_artifact(self):
        """Una cuota de artefactos no salta el check requerido (cadete#752)."""
        self.assertNotIn("download-artifact", consumer_workflow())

    def test_the_verdict_is_mirrored_on_the_pr_head_with_the_github_token(self):
        """El check run nativo de `pull_request_target` queda en la base; el merge
        gobernado lee la cabeza, así que el espejo se queda y su motivo está escrito (#1396, S9)."""
        tree = policy_gate.parse_workflow(consumer_workflow(), "ci.yml")
        mirror = [s for s in tree["jobs"][JOB]["steps"] if "check-runs" in str(s.get("run") or "")]
        self.assertEqual(len(mirror), 1)
        self.assertEqual(mirror[0]["env"]["CONTEXT"], "governance-seals")
        self.assertEqual(mirror[0]["env"]["HEAD_SHA"], "${{ github.event.pull_request.head.sha }}")
        self.assertEqual(mirror[0]["env"]["GH_TOKEN"], "${{ github.token }}")
        guide = " ".join(PORTING_GUIDE.read_text(encoding="utf-8").split())
        self.assertIn("Por qué el gate aparece dos veces entre los checks del PR (#1396)", guide)

    # --- #1172: el manifiesto re-sellado como artefacto, sin push ---

    def test_the_fragment_reseals_the_manifest_without_pushing(self):
        """El job deja el manifiesto re-sellado como artefacto; nada se empuja."""
        tree = policy_gate.parse_workflow(consumer_workflow(), "ci.yml")
        commands = "\n".join(
            str(step.get("run") or "") for step in tree["jobs"][JOB]["steps"])
        self.assertIn("manifest seal pr-head", commands)
        self.assertIn("--out", commands)
        # Re-sellar es escribir el manifiesto: el job no lo commitea ni lo empuja.
        for forbidden in ("git commit", "git push", "git add"):
            self.assertNotIn(forbidden, commands)

    def test_the_reseal_steps_only_run_when_the_gate_failed(self):
        """`if: failure()`: el artefacto se publica cuando el sello no coincide."""
        tree = policy_gate.parse_workflow(consumer_workflow(), "ci.yml")
        steps = tree["jobs"][JOB]["steps"]
        reseal = [s for s in steps if "manifest seal" in str(s.get("run") or "")]
        self.assertEqual(len(reseal), 1, "el job no re-sella el manifiesto")
        self.assertIn("failure()", str(reseal[0].get("if") or ""))
        upload = [s for s in steps
                  if "upload-artifact" in str(s.get("uses") or "")]
        self.assertEqual(len(upload), 1, "el job no publica el manifiesto re-sellado")
        self.assertIn("failure()", str(upload[0].get("if") or ""))
        # HR-26: la acción va pineada a un sha de 40 hex.
        self.assertRegex(str(upload[0].get("uses")), r"@[0-9a-f]{40}$")

    def test_the_reseal_steps_do_not_swallow_the_verdict(self):
        """Ni `|| true` ni `continue-on-error`: el rojo del gate sigue siendo rojo."""
        tree = policy_gate.parse_workflow(consumer_workflow(), "ci.yml")
        for step in tree["jobs"][JOB]["steps"]:
            command = str(step.get("run") or "")
            self.assertNotIn("|| true", command)
            self.assertNotIn("|| exit 0", command)
            self.assertNotIn("continue-on-error", str(step.get("continue-on-error") or ""))


if __name__ == "__main__":
    unittest.main()
