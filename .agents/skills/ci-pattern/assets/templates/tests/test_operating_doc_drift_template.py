#!/usr/bin/env python3
# ci-pattern asset tests — plantilla del workflow operating-doc-drift (#1062)
"""Suite de la plantilla que deja de declarar el commit del catálogo dos veces.

`assets/templates/.github/workflows/operating-doc-drift.yml.tmpl` es el workflow
de deriva del documento operativo que el canon publica para que un consumer no
lleve el commit del catálogo escrito a mano en el `ref:` de su checkout. La fuente
única es el sello `# Auto-generado por propagate-team-skills.ps1 (catalog @
<sha40>).` de `.team-skills.yaml`, que el propagador reescribe en cada corrida:
la plantilla lo resuelve en runtime y exporta a `$GITHUB_OUTPUT` (#1062).

Lo que mide, con los mecanismos reales y no con una política inventada:

1. La plantilla pasa el `workflow-policy` del canon junto con los workflows reales
   de este repositorio y su política real de jobs requeridos.
2. No lleva ningún sha de 40 hex escrito a mano: el único 40 hex que puede
   aparecer es el ref de una action en `uses:` (HR-26). El detector se demuestra
   en un caso negativo sobre una copia con el `ref:` literal inyectado.
3. El paso que resuelve el sello existe, escribe en `$GITHUB_OUTPUT` y su output
   es el `ref:` del checkout del catálogo.
4. Ese paso se ejecuta de verdad: con el sello bien formado exporta el commit, y
   sale distinto de 0 nombrando el dato que falta cuando `.team-skills.yaml` no
   existe, no trae el sello `catalog @` o el sha no es de 40 hex (HR-3).
5. Solo se dispara en `schedule` y `workflow_dispatch` sin publicar check
   requerido (HR-30), y los dos gates los corre la CLI del checkout del catálogo.

Ejecuta sin red y sin consumer.
"""
from __future__ import annotations

import importlib
import os
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
REPO = ASSET.parents[4]
WORKFLOW_TMPL = ASSET / ".github" / "workflows" / "operating-doc-drift.yml.tmpl"
WORKFLOW_POLICY = REPO / ".github" / "workflow-policy.json"
GATE = ASSET.parent / "workflow-policy" / "check_workflow_policy.py"

if str(ASSET.parent / "workflow-policy") not in sys.path:
    sys.path.insert(0, str(ASSET.parent / "workflow-policy"))
policy_gate = importlib.import_module("check_workflow_policy")

SHA40 = re.compile(r"(?<![0-9a-fA-F])[0-9a-f]{40}(?![0-9a-fA-F])")
#: El ref de una action en `uses:` es el único sha de 40 hex que un workflow
#: puede llevar escrito a mano: lo exige HR-26 y no lo escribe el propagador.
USES_REF = re.compile(r"\buses:\s*[A-Za-z0-9._/-]+@[0-9a-f]{40}")
#: `ref: ${{ steps.<id>.outputs.<name> }}`: el pin se resuelve, no se escribe.
REF_OUTPUT = re.compile(r"^\$\{\{\s*steps\.([A-Za-z0-9_-]+)\.outputs\.([A-Za-z0-9_-]+)\s*\}\}$")
#: El pin que el porting-guide publicaba escrito a mano, reutilizado como fixture
#: del caso negativo: si el detector no lo ve, no está mirando lo que dice mirar.
LITERAL_PIN = "bc915526af7920e459afd2532e93e96e0479c0e6"
GOOD_COMMIT = "1234567890abcdef1234567890abcdef12345678"
SEAL = f"# Auto-generado por propagate-team-skills.ps1 (catalog @ {GOOD_COMMIT})."
MANIFEST = ".team-skills.yaml"


def hand_written_shas(text: str) -> list[str]:
    """Los sha de 40 hex escritos a mano, fuera del ref de una action (`uses:`)."""
    found = []
    for n, line in enumerate(text.splitlines(), 1):
        for match in SHA40.finditer(USES_REF.sub("", line)):
            found.append(f"linea {n}: {match.group(0)}")
    return found


def inject_literal_ref(text: str) -> str:
    """La plantilla tal como estaba antes de #1062: el `ref:` del catálogo escrito."""
    pinned, count = re.subn(r"(?m)^(\s*)ref: \S.*$", rf"\g<1>ref: {LITERAL_PIN}", text, count=1)
    if count != 1:
        raise AssertionError(
            "la plantilla ya no trae un `ref:` de expression que sustituir, o trae mas de uno: "
            "el fixture del caso negativo ya no aplica"
        )
    return pinned


def resolve_step(doc: dict) -> dict:
    """El paso que resuelve el sello: el único que escribe en `$GITHUB_OUTPUT`."""
    steps = [s for job in (doc.get("jobs") or {}).values()
             for s in (job.get("steps") or [])
             if isinstance(s, dict) and "GITHUB_OUTPUT" in str(s.get("run") or "")]
    if len(steps) != 1:
        raise AssertionError(
            f"se esperaba un paso que exporte el commit a $GITHUB_OUTPUT, hay {len(steps)}: "
            f"sin el, el `ref:` del checkout del catálogo no tiene de donde leer (#1062)"
        )
    return steps[0]


def checkout_using(doc: dict, step_id: str) -> dict:
    """El `with:` del checkout cuyo `ref:` es el output del paso `step_id`."""
    for job in (doc.get("jobs") or {}).values():
        for step in job.get("steps") or []:
            uses = step.get("uses") if isinstance(step, dict) else None
            with_ = step.get("with") if isinstance(step, dict) else None
            if not isinstance(uses, str) or "actions/checkout" not in uses or not isinstance(with_, dict):
                continue
            found = REF_OUTPUT.fullmatch(str(with_.get("ref") or ""))
            if found and found.group(1) == step_id:
                return with_
    raise AssertionError(
        f"ningun checkout usa `ref: ${{{{ steps.{step_id}.outputs.* }}}}`: "
        f"el commit del catálogo se seguiria escribiendo a mano (#1062)"
    )


class TestOperatingDocDriftTemplate(unittest.TestCase):
    def template(self) -> str:
        self.assertTrue(WORKFLOW_TMPL.is_file(), f"no existe la plantilla {WORKFLOW_TMPL} (#1062)")
        return WORKFLOW_TMPL.read_text(encoding="utf-8")

    def tree(self) -> dict:
        """El árbol que ve el gate, leído con su propio lector estructural."""
        return policy_gate.parse_workflow(self.template(), "operating-doc-drift.yml")

    def declared_env(self, doc: dict | None = None) -> dict:
        env = (doc or self.tree()).get("env")
        self.assertIsInstance(env, dict, "los valores parametrizables van en el bloque env de arriba")
        return {str(k): str(v) for k, v in env.items()}

    def commands(self, doc: dict | None = None) -> str:
        tree = doc or self.tree()
        return "\n".join(str(step.get("run") or "")
                         for job in tree["jobs"].values() for step in job.get("steps") or [])

    # --- punto 1: el gate del canon, con la configuración real de este repositorio ---

    def test_la_plantilla_pasa_el_workflow_policy_del_canon(self):
        with tempfile.TemporaryDirectory() as tmp:
            rendered = Path(tmp) / "operating-doc-drift.yml"
            rendered.write_text(self.template(), encoding="utf-8")
            cmd = [sys.executable, str(GATE), "--required-jobs", str(WORKFLOW_POLICY),
                   "--base-branch", "main", "--workflow", str(rendered)]
            for w in (REPO / ".github" / "workflows").glob("*.yml"):
                cmd += ["--workflow", str(w)]
            proc = subprocess.run(cmd, capture_output=True, text=True, cwd=REPO)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("WORKFLOW-POLICY OK", proc.stdout)

    # --- punto 2: ni un sha escrito a mano, y el detector que lo vigila ---

    def test_no_declara_el_commit_del_catalogo_por_escrito(self):
        self.assertEqual(hand_written_shas(self.template()), [])

    def test_el_detector_ve_un_ref_literal_inyectado(self):
        with tempfile.TemporaryDirectory() as tmp:
            copy = Path(tmp) / "operating-doc-drift.yml"
            copy.write_text(inject_literal_ref(self.template()), encoding="utf-8")
            found = hand_written_shas(copy.read_text(encoding="utf-8"))
        self.assertEqual(len(found), 1, f"el detector no vio el `ref:` inyectado: {found}")
        self.assertTrue(found[0].endswith(LITERAL_PIN), found)

    # --- punto 3: el sello se resuelve y su output es el ref del checkout ---

    def test_el_pin_se_resuelve_desde_el_sello_y_alimenta_el_checkout(self):
        tree = self.tree()
        step = resolve_step(tree)
        self.assertTrue(str(step.get("id") or "").strip(),
                        "el paso que resuelve el sello necesita `id:` para que el checkout lo lea")
        script = str(step["run"])
        self.assertIn("$GITHUB_OUTPUT", script)
        declared = self.declared_env(tree)
        manifest = [k for k, v in declared.items() if v.strip() == MANIFEST]
        self.assertEqual(len(manifest), 1, f"el manifiesto del catalogo se declara en `env`: {declared}")
        self.assertIn(f"${manifest[0]}", script)
        self.assertIn("repository", checkout_using(tree, str(step["id"])))

    # --- punto 4: el paso se ejecuta, y falla en voz alta cuando falta el dato ---

    def run_resolver(self, manifest: str | None, api: dict[str, str] | None = None) -> tuple[int, str, str, str]:
        """Ejecuta el paso de la plantilla con bash sobre un consumer de mentira.

        `api` simula la API del catálogo: ruta relativa bajo `CATALOG_API` -> cuerpo.
        Con `file://` curl la lee del disco, así que el caso no necesita red.
        """
        script = str(resolve_step(self.tree())["run"])
        env_values = {k: ("" if "${{" in v else v) for k, v in self.declared_env().items()}
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp)
            if manifest is not None:
                (repo / MANIFEST).write_text(manifest, encoding="utf-8")
            if api is not None:
                for rel, body in api.items():
                    (repo / "api" / rel).parent.mkdir(parents=True, exist_ok=True)
                    (repo / "api" / rel).write_text(body, encoding="utf-8")
                env_values["CATALOG_API"] = f"file://{repo / 'api'}"
            (repo / "resolve.sh").write_text(script, encoding="utf-8")
            outputs = repo / "github-output"
            outputs.touch()
            proc = subprocess.run(["bash", "resolve.sh"], cwd=repo,
                                  env={**os.environ, **env_values, "GITHUB_OUTPUT": str(outputs)},
                                  capture_output=True, text=True)
            return proc.returncode, proc.stdout, proc.stderr, outputs.read_text(encoding="utf-8")

    def test_con_el_sello_bien_formado_exporta_el_commit(self):
        rc, out, err, exported = self.run_resolver(f"{SEAL}\nproject_type: web\n")
        self.assertEqual(rc, 0, err)
        self.assertIn(GOOD_COMMIT, exported)
        self.assertIn(GOOD_COMMIT, out)

    def test_sin_manifesto_falla_nombrando_el_fichero(self):
        rc, _, err, exported = self.run_resolver(None)
        self.assertNotEqual(rc, 0, "sin .team-skills.yaml el job debe fallar, no salir verde (HR-3)")
        self.assertIn(MANIFEST, err)
        self.assertNotIn(GOOD_COMMIT, exported)

    def test_sin_sello_falla_nombrando_el_sello(self):
        rc, _, err, exported = self.run_resolver("project_type: web\ntiers:\n  - universal\n")
        self.assertNotEqual(rc, 0, "un manifiesto sin sello `catalog @` no dice qué commit medir (HR-3)")
        self.assertIn("catalog @", err)
        self.assertNotIn(GOOD_COMMIT, exported)

    def test_con_un_sha_demasiado_corto_falla_nombrando_el_sha(self):
        rc, _, err, exported = self.run_resolver(
            "# Auto-generado por propagate-team-skills.ps1 (catalog @ 123456).\nproject_type: web\n")
        self.assertNotEqual(rc, 0, "un prefijo de menos de 7 hex no identifica un commit (HR-3)")
        self.assertIn("123456", err)
        self.assertNotIn(GOOD_COMMIT, exported)

    def test_sello_abreviado_se_expande_con_la_api_del_catalogo(self):
        # #1157: el sello que ya está en el campo es `git rev-parse --short`.
        # La API resuelve el prefijo al oid completo y el paso exporta ese oid.
        rc, out, err, exported = self.run_resolver(
            "# Auto-generado por propagate-team-skills.ps1 (catalog @ 1234567).\nproject_type: web\n",
            api={"repos/DysTelefonica/team-skills/commits/1234567": f'{{"sha": "{GOOD_COMMIT}"}}'})
        self.assertEqual(rc, 0, err)
        self.assertIn(GOOD_COMMIT, exported)

    def test_sello_abreviado_que_la_api_no_conoce_falla_nombrando_el_sha(self):
        rc, _, err, exported = self.run_resolver(
            "# Auto-generado por propagate-team-skills.ps1 (catalog @ 1234567).\nproject_type: web\n",
            api={})
        self.assertNotEqual(rc, 0, "un sello que el catálogo no resuelve no puede medirse (HR-3)")
        self.assertIn("1234567", err)
        self.assertNotIn(GOOD_COMMIT, exported)

    # --- punto 5: disparo, check requerido y los dos gates con la CLI del checkout ---

    def test_solo_schedule_y_dispatch_y_no_publica_check_requerido(self):
        tree = self.tree()
        events = policy_gate.workflow_events(tree)
        self.assertEqual(list(events), ["schedule", "workflow_dispatch"])
        self.assertNotIn("pull_request", events)
        # HR-30: aquí no hay PR que evaluar, así que el nombre no se exige como check.
        self.assertIn("HR-30", self.template())
        for jid, job in tree["jobs"].items():
            self.assertNotIn("continue-on-error", str(job), jid)
        self.assertNotIn("|| true", self.commands(tree))

    def test_los_dos_gates_coren_con_la_cli_del_checkout_del_catalogo(self):
        tree = self.tree()
        commands = self.commands(tree)
        cli = [k for k, v in self.declared_env(tree).items() if v.strip().endswith("assets/bin/ci-pattern")]
        self.assertEqual(len(cli), 1, self.declared_env(tree))
        for needle in ("params validate", "adoption check --phase 5"):
            lines = [ln for ln in commands.splitlines() if needle in ln]
            self.assertEqual(len(lines), 1, f"`{needle}` deberia ser el comando de su propio paso")
            self.assertIn(f"${cli[0]}", lines[0])


if __name__ == "__main__":
    unittest.main()
