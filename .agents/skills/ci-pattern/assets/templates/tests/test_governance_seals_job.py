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
# Esqueleto mínimo de un `ci.yml` de consumer con el agregador `required`: el
# fragmento del porting-guide se inyecta como job más, con su entrada en el
# `needs` del agregador, que es como lo cablea un adopter.
SKELETON_HEAD = """name: ci
on:
  pull_request:
  push:
    branches: [main]
concurrency:
  group: wf-${{ github.ref }}
jobs:
"""
SKELETON_TAIL = """  required:
    name: required
    needs: [unit, JOB_NAME]
    if: always()
    timeout-minutes: 5
    steps:
      - run: echo gate
"""


def canonical_job() -> str:
    """El fragmento que publica el porting-guide, extraído del documento.

    Se busca el bloque cercado que empieza con la clave `governance-seals:`: si
    el documento deja de.publisharlo, esta función revienta y el fallo nombra
    el documento, no una constante copiada que nadie lee.
    """
    text = PORTING_GUIDE.read_text(encoding="utf-8")
    for block in re.findall(r"```yaml\n(.*?)```", text, re.S):
        if f"{JOB}:" in block:
            # El bloque vive dentro de una lista numerada del documento, con la
            # sangria de esa lista encima; sin `dedent` el fragmento no seria un
            # workflow valido. `dedent` se lleva tambien los dos espacios del
            # nivel `jobs:`, asi que se devuelven: la linea del job tiene que
            # quedar en `  governance-seals:` o el parser ve contenido suelto.
            flat = textwrap.dedent(block).rstrip("\n")
            return "".join("  " + line + "\n" if line.strip() else line + "\n"
                           for line in flat.splitlines())
    raise AssertionError(
        f"el porting-guide ya no publica el fragmento del job `{JOB}`: "
        f"la forma canónica de HR-72 desapareció de {PORTING_GUIDE}"
    )


def consumer_workflow() -> str:
    return (SKELETON_HEAD + canonical_job()
            + SKELETON_TAIL.replace("JOB_NAME", JOB))


class CanonicalJobTests(unittest.TestCase):
    def test_fragment_passes_the_workflow_policy_gate(self):
        """La forma publicada pasa el gate de política con la política real."""
        with tempfile.TemporaryDirectory() as tmp:
            rendered = Path(tmp) / "ci.yml"
            rendered.write_text(consumer_workflow(), encoding="utf-8")
            cmd = [sys.executable, str(GATE), "--required-jobs", str(WORKFLOW_POLICY),
                   "--base-branch", "main", "--workflow", str(rendered)]
            for w in (REPO / ".github" / "workflows").glob("*.yml"):
                cmd += ["--workflow", str(w)]
            proc = subprocess.run(cmd, capture_output=True, text=True, cwd=REPO)
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
        """HR-30: el nombre se publica solo desde el evento que evalúa el PR."""
        tree = policy_gate.parse_workflow(consumer_workflow(), "ci.yml")
        job = tree["jobs"][JOB]
        cond = str(job.get("if") or "")
        self.assertIn("pull_request", cond, cond)
        for event in ("schedule", "workflow_dispatch"):
            self.assertNotIn(event, cond, cond)

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


if __name__ == "__main__":
    unittest.main()
