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
import os
import shutil
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

    def ci_gate_argv(self, drop=(), **env) -> list[str]:
        """Argumentos con los que el paso del gate de CI llama a Python, según
        el bloque `env`: el paso se ejecuta con bash y un `python3` falso."""
        if shutil.which("bash") is None:
            self.skipTest("bash no disponible")
        step = next(s for s in self.doc["jobs"]["preflight"]["steps"]
                    if "check_release_ci.py" in str(s.get("run") or ""))
        with tempfile.TemporaryDirectory() as tmp:
            stub = Path(tmp) / "python3"
            stub.write_text('#!/bin/sh\nprintf "%s\\n" "$@"\n', encoding="utf-8")
            stub.chmod(0o755)
            script = Path(tmp) / "step.sh"
            script.write_text(step["run"], encoding="utf-8")
            environ = {**os.environ, **self.doc["env"], "GITHUB_REF_NAME": "v1.0.1", **env,
                       "PATH": tmp + os.pathsep + os.environ.get("PATH", "")}
            for name in drop:
                environ.pop(name, None)
            out = subprocess.run(["bash", str(script)], env=environ, capture_output=True,
                                 text=True, check=True).stdout
        return out.splitlines()

    def test_por_defecto_el_gate_de_ci_no_espera_ni_filtra_el_run_del_tag(self):
        self.assertEqual(self.doc["env"]["CI_TAG_RUN"], "false")
        self.assertEqual(self.doc["env"]["CI_WAIT_MINUTES"], "0")
        argv = self.ci_gate_argv()
        self.assertNotIn("--tag-run", argv)
        self.assertEqual(argv[argv.index("--wait-minutes") + 1], "0")

    def test_ci_tag_run_espera_al_run_del_push_del_tag(self):
        # Consumers cuyo CI solo ejecuta una suite en el push del tag (#1451).
        argv = self.ci_gate_argv(CI_TAG_RUN="true", CI_WAIT_MINUTES="60")
        self.assertIn("--tag-run", argv)
        self.assertEqual(argv[argv.index("--wait-minutes") + 1], "60")
        self.assertEqual(argv[argv.index("--tag") + 1], "v1.0.1")

    def test_un_bloque_env_antiguo_sin_los_parametros_no_espera(self):
        # Un consumer que actualiza el paso y conserva su bloque `env` anterior.
        argv = self.ci_gate_argv(drop=("CI_TAG_RUN", "CI_WAIT_MINUTES"))
        self.assertNotIn("--tag-run", argv)
        self.assertEqual(argv[argv.index("--wait-minutes") + 1], "0")

    def test_ci_tag_run_con_otro_valor_sale_dos_en_vez_de_medir_main(self):
        with self.assertRaises(subprocess.CalledProcessError) as ctx:
            self.ci_gate_argv(CI_TAG_RUN="True")
        self.assertEqual(ctx.exception.returncode, 2)
        self.assertIn("CI_TAG_RUN", ctx.exception.stderr)

    def run_step(self, job: str, marker: str, cwd: str, drop=(), **env):
        """Ejecuta con bash el paso de `job` cuyo `run` contiene `marker`, con el
        bloque `env` de la plantilla, y devuelve el proceso terminado."""
        if shutil.which("bash") is None:
            self.skipTest("bash no disponible")
        step = next(s for s in self.doc["jobs"][job]["steps"]
                    if marker in str(s.get("run") or ""))
        script = Path(cwd) / "step.sh"
        script.write_text(step["run"], encoding="utf-8")
        environ = {**os.environ, **self.doc["env"], "GITHUB_REF_NAME": "v1.0.1",
                   "GITHUB_ENV": str(Path(cwd) / "github_env"),
                   "GITHUB_STEP_SUMMARY": str(Path(cwd) / "summary"), **env}
        for name in drop:
            environ.pop(name, None)
        return subprocess.run(["bash", str(script)], env=environ, cwd=cwd,
                              capture_output=True, text=True)

    # --- #1456: ruta de las notas, entorno, runner y verificación aplazable ---

    def notes_file(self, files, drop=(), **env):
        with tempfile.TemporaryDirectory() as tmp:
            for name in files:
                (Path(tmp) / name).parent.mkdir(parents=True, exist_ok=True)
                (Path(tmp) / name).write_text("notas\n", encoding="utf-8")
            proc = self.run_step("release", "RELEASE_NOTES_PATH", tmp, drop=drop, **env)
            github_env = Path(tmp) / "github_env"
            exported = github_env.read_text(encoding="utf-8") if github_env.exists() else ""
        return proc, exported

    def test_por_defecto_las_notas_son_release_tag_en_la_raiz(self):
        self.assertEqual(self.doc["env"]["RELEASE_NOTES_PATH"], "RELEASE-{tag}.md")
        proc, exported = self.notes_file(["RELEASE-v1.0.1.md"])
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertIn("RELEASE_NOTES_FILE=RELEASE-v1.0.1.md", exported)

    def test_release_notes_path_lleva_las_notas_a_otra_ruta(self):
        proc, exported = self.notes_file(["docs/releases/v1.0.1.md"],
                                         RELEASE_NOTES_PATH="docs/releases/{tag}.md")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertIn("RELEASE_NOTES_FILE=docs/releases/v1.0.1.md", exported)

    def test_sin_el_fichero_de_notas_el_paso_falla_y_lo_nombra(self):
        proc, exported = self.notes_file([], RELEASE_NOTES_PATH="docs/releases/{tag}.md")
        self.assertEqual(proc.returncode, 1)
        self.assertIn("docs/releases/v1.0.1.md", proc.stderr)
        self.assertEqual(exported, "")

    def test_un_bloque_env_antiguo_sin_ruta_de_notas_usa_la_raiz(self):
        proc, exported = self.notes_file(["RELEASE-v1.0.1.md"], drop=("RELEASE_NOTES_PATH",))
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertIn("RELEASE_NOTES_FILE=RELEASE-v1.0.1.md", exported)

    def test_la_publicacion_lee_las_notas_del_checkout_del_tag(self):
        steps = self.doc["jobs"]["release"]["steps"]
        self.assertIn("actions/checkout@", str(steps[0].get("uses")))
        self.assertIn('--notes-file "$RELEASE_NOTES_FILE"', run_commands(self.doc["jobs"]["release"]))

    def test_el_entorno_es_release_salvo_que_la_variable_diga_none(self):
        # `environment` no admite el contexto `env`: el parámetro es la variable
        # del repositorio RELEASE_ENVIRONMENT (sin definir = `release`).
        expr = self.doc["jobs"]["release"]["environment"]
        self.assertEqual(expr, "${{ (vars.RELEASE_ENVIRONMENT || 'release') != 'none'"
                               " && (vars.RELEASE_ENVIRONMENT || 'release') || '' }}")

    def test_los_cuatro_jobs_corren_en_el_mismo_runner_literal(self):
        # HR-61: un runs-on dinámico es hallazgo con perfil self-hosted, así que
        # el runner es un literal que el consumer cambia igual en los cuatro jobs.
        runners = {name: job["runs-on"] for name, job in self.doc["jobs"].items()}
        self.assertEqual(set(runners), {"preflight", "release", "verify", "provenance"})
        self.assertEqual(set(map(str, runners.values())), {"ubuntu-24.04"})
        self.assertEqual(self.text.count("runs-on: ubuntu-24.04  # RELEASE_RUNS_ON"), 4)

    def verify_output(self, drop=(), **env):
        with tempfile.TemporaryDirectory() as tmp:
            script = Path(tmp) / "verify.sh"
            script.write_text("#!/bin/sh\necho SCRIPT-RAN\n", encoding="utf-8")
            script.chmod(0o755)
            proc = self.run_step("verify", "RELEASE_VERIFY_SCRIPT", tmp, drop=drop,
                                 RELEASE_VERIFY_SCRIPT=str(script), **env)
            summary = Path(tmp) / "summary"
            proc.summary = summary.read_text(encoding="utf-8") if summary.exists() else ""
        return proc

    def test_por_defecto_verify_ejecuta_el_script(self):
        self.assertEqual(self.doc["env"]["RELEASE_VERIFY"], "run")
        proc = self.verify_output()
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertIn("SCRIPT-RAN", proc.stdout)

    def test_un_bloque_env_antiguo_sin_release_verify_ejecuta_el_script(self):
        proc = self.verify_output(drop=("RELEASE_VERIFY", "RELEASE_VERIFY_DEFERRED_REASON"))
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertIn("SCRIPT-RAN", proc.stdout)

    def test_verify_aplazado_no_ejecuta_el_script_y_lo_dice(self):
        proc = self.verify_output(RELEASE_VERIFY="deferred",
                                  RELEASE_VERIFY_DEFERRED_REASON="producción tras VPN")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertNotIn("SCRIPT-RAN", proc.stdout)
        self.assertIn("::warning", proc.stdout)
        self.assertIn("producción tras VPN", proc.stdout)
        self.assertIn("producción tras VPN", proc.summary)

    def test_verify_aplazado_sin_motivo_sale_dos(self):
        proc = self.verify_output(RELEASE_VERIFY="deferred", RELEASE_VERIFY_DEFERRED_REASON="")
        self.assertEqual(proc.returncode, 2)
        self.assertIn("RELEASE_VERIFY_DEFERRED_REASON", proc.stderr)
        self.assertNotIn("SCRIPT-RAN", proc.stdout)

    def test_verify_aplazado_con_motivo_en_blanco_sale_dos(self):
        proc = self.verify_output(RELEASE_VERIFY="deferred", RELEASE_VERIFY_DEFERRED_REASON="   ")
        self.assertEqual(proc.returncode, 2)
        self.assertIn("RELEASE_VERIFY_DEFERRED_REASON", proc.stderr)
        self.assertNotIn("SCRIPT-RAN", proc.stdout)

    def test_release_verify_con_otro_valor_sale_dos(self):
        proc = self.verify_output(RELEASE_VERIFY="skip")
        self.assertEqual(proc.returncode, 2)
        self.assertIn("RELEASE_VERIFY", proc.stderr)
        self.assertNotIn("SCRIPT-RAN", proc.stdout)

    def test_el_preflight_deja_tiempo_a_la_espera_maxima(self):
        # Espera máxima documentada (100) más una consulta y los pasos previos.
        timeout = int(self.doc["jobs"]["preflight"]["timeout-minutes"])
        self.assertGreaterEqual(timeout, 100 + 10)
        self.assertIn("como mucho 100", self.text)

    def test_la_publicacion_va_en_entorno_protegido_y_sin_admin(self):
        job = self.doc["jobs"]["release"]
        self.assertEqual(job["needs"], "preflight")
        # Sin RELEASE_ENVIRONMENT, el entorno protegido `release`.
        self.assertIn("vars.RELEASE_ENVIRONMENT || 'release'", job["environment"])
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

    # --- HR-83: manifiesto de procedencia de imagen como asset de la release ---

    def test_el_job_provenance_sube_el_manifiesto_del_tag_si_hay_config(self):
        job = self.doc["jobs"]["provenance"]
        self.assertEqual(job["needs"], "release")
        self.assertEqual(job["permissions"], {"contents": "write"})
        commands = run_commands(job)
        self.assertIn('if [ ! -f "$IMAGE_PROVENANCE_CONFIG" ]', commands)
        self.assertIn("image-provenance/image_provenance.py\" manifest", commands)
        self.assertIn('--tag "$GITHUB_REF_NAME"', commands)
        self.assertIn('gh release upload "$GITHUB_REF_NAME" image-provenance-manifest.json',
                      commands)
        self.assertNotIn("actions/upload-artifact", self.text)

    def test_el_script_de_verificacion_compara_lo_desplegado_con_el_sha_del_tag(self):
        script = VERIFY_TMPL.read_text(encoding="utf-8")
        proc = subprocess.run(["bash", "-n", str(VERIFY_TMPL)], capture_output=True, text=True)
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertIn('RELEASE_SHA', script)
        self.assertIn('RELEASE_HEALTH_URL', script)
        self.assertIn('[[ "$revision" == "$RELEASE_SHA" ]]', script)


if __name__ == "__main__":
    unittest.main()
