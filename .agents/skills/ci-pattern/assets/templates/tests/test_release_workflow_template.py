#!/usr/bin/env python3
# ci-pattern asset tests — plantilla de workflow de release (#1015)
"""Suite de la plantilla de release: pasa el `workflow-policy` del canon.

La plantilla vive en `assets/templates/.github/workflows/release.yml.tmpl` y
un consumer la materializa en `.github/workflows/release.yml`. El test la
renderiza en un temporal y la somete al gate de política de workflows del
propio canon, junto con los workflows reales de este repositorio y su
política de jobs requeridos: así se mide como lo mide el CI de un consumer
(todos los workflows a la vez, con `required_jobs` = checks de PR), no con una
política inventada para el caso.

Sobre el árbol ya parseado comprueba además lo que el gate no puede saber:
que el disparo es solo de tag estable, que la publicación va en el entorno
protegido y sin `--admin`, y que el job `verify` comprueba la salida.

Ejecuta sin red y sin consumer.
"""
import importlib
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
REPO = ASSET.parents[4]
WORKFLOW_TMPL = ASSET / ".github" / "workflows" / "release.yml.tmpl"
VERIFY_TMPL = ASSET / "scripts" / "verify-release-output.sh.tmpl"
WORKFLOW_POLICY = REPO / ".github" / "workflow-policy.json"
GATE = ASSET.parent / "workflow-policy" / "check_workflow_policy.py"

if str(ASSET.parent / "workflow-policy") not in sys.path:
    sys.path.insert(0, str(ASSET.parent / "workflow-policy"))
policy_gate = importlib.import_module("check_workflow_policy")


def parse(text: str) -> dict:
    """El árbol que ve el gate, leído con su propio lector estructural."""
    return policy_gate.parse_workflow(text, "release.yml")


def run_commands(job: dict) -> str:
    return "\n".join(str(step.get("run") or "") for step in job.get("steps") or [])


class TestReleaseWorkflowTemplate(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.text = WORKFLOW_TMPL.read_text(encoding="utf-8")
        cls.doc = parse(cls.text)

    def rendered(self, tmp: str) -> Path:
        target = Path(tmp) / "release.yml"
        target.write_text(self.text, encoding="utf-8")
        return target

    # --- el gate del canon, con la configuración real de este repositorio ---

    def test_la_plantilla_pasa_el_workflow_policy_del_canon(self):
        with tempfile.TemporaryDirectory() as tmp:
            rendered = self.rendered(tmp)
            canon_workflows = [w for w in (REPO / ".github" / "workflows").glob("*.yml")
                               if w.name != "release.yml"]
            cmd = [sys.executable, str(GATE), "--required-jobs", str(WORKFLOW_POLICY),
                   "--base-branch", "main", "--workflow", str(rendered)]
            for w in canon_workflows:
                cmd += ["--workflow", str(w)]
            proc = subprocess.run(cmd, capture_output=True, text=True, cwd=REPO)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("WORKFLOW-POLICY OK", proc.stdout)

    # --- punto 2: disparo, gates previos y publicación ---

    def test_se_dispara_solo_con_tags_estables(self):
        events = policy_gate.workflow_events(self.doc)
        self.assertEqual(list(events), ["push"])
        self.assertEqual(events["push"].get("tags"), ["v*", "!v*-*"])

    def test_el_job_previo_ejecuta_los_dos_gates(self):
        commands = run_commands(self.doc["jobs"]["preflight"])
        self.assertIn("check_release_tag.py", commands)
        self.assertIn("check_release_ci.py", commands)
        self.assertIn("--workflow", commands)

    def test_la_publicacion_va_en_entorno_protegido_y_sin_admin(self):
        job = self.doc["jobs"]["release"]
        self.assertEqual(job["needs"], "preflight")
        self.assertEqual(job["environment"], "release")
        self.assertEqual(job["permissions"], {"contents": "write"})
        self.assertIn("--latest", run_commands(job))
        self.assertNotIn("--admin", run_commands(job))
        # El token con escritura solo existe en el job que publica.
        self.assertEqual(self.doc["permissions"], {"contents": "read"})
        self.assertEqual(self.doc["jobs"]["preflight"]["permissions"]["contents"], "read")
        self.assertEqual(self.doc["jobs"]["verify"]["permissions"]["contents"], "read")

    # --- punto 3: la salida publicada corresponde al tag ---

    def test_el_job_verify_comprueba_la_salida_y_no_repite_tests(self):
        job = self.doc["jobs"]["verify"]
        self.assertEqual(job["needs"], "release")
        commands = run_commands(job)
        self.assertIn("releases/tags/", commands)
        self.assertIn("releases/latest", commands)
        self.assertIn('RELEASE_SHA=', commands)
        self.assertIn("$RELEASE_VERIFY_SCRIPT", commands)
        self.assertNotIn("pytest", commands)
        self.assertNotIn("run-unit-suites", commands)

    def test_el_script_de_verificacion_compara_lo_desplegado_con_el_sha_del_tag(self):
        script = VERIFY_TMPL.read_text(encoding="utf-8")
        proc = subprocess.run(["bash", "-n", str(VERIFY_TMPL)], capture_output=True, text=True)
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertIn('RELEASE_SHA', script)
        self.assertIn('RELEASE_HEALTH_URL', script)
        self.assertIn('[[ "$revision" == "$RELEASE_SHA" ]]', script)


if __name__ == "__main__":
    unittest.main()
