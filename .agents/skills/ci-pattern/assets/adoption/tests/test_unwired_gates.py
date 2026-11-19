"""#1742: aviso cuando un gate distribuido no lo invoca ningún workflow del consumer.

El canon de gates es `assets/*/check_*.py` de la skill (una sola fuente). La
invariante es la línea HR de SKILL.md que cita el asset. Es detectivo: avisa por
stderr y no añade hallazgos, así que no cambia el veredicto de la fase 4.
"""

import contextlib
import importlib.machinery
import importlib.util
import io
import tempfile
import unittest
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[3]
CLI_PATH = SKILL_ROOT / "assets" / "bin" / "ci-pattern"
E2E = "assets/e2e-evidence/check_e2e_evidence.py"


def _load_cli():
    spec = importlib.util.spec_from_loader(
        "ci_pattern_cli_unwired_gates",
        importlib.machinery.SourceFileLoader("ci_pattern_cli_unwired_gates", str(CLI_PATH)),
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


CLI = _load_cli()


def _write(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _consumer(root, gates=(E2E,), workflow=""):
    """Consumer con los assets de `gates` vendorizados y un workflow dado."""
    repo = Path(root)
    for rel in gates:
        _write(repo / ".agents" / "skills" / "ci-pattern" / rel, "# vendored\n")
    _write(repo / ".github" / "workflows" / "ci.yml", workflow or "name: ci\non: pull_request\njobs: {}\n")
    return repo


WIRED = ("name: ci\non: pull_request\njobs:\n  gates:\n    steps:\n"
         "      - run: python3 .agents/skills/ci-pattern/" + E2E + " --repo .\n")


def _warn(repo):
    buf = io.StringIO()
    with contextlib.redirect_stderr(buf):
        CLI._warn_unwired_gates(repo)
    return buf.getvalue()


class UnwiredGatesTests(unittest.TestCase):
    def test_fuente_unica_es_el_canon_de_assets(self):
        esperado = sorted(p.relative_to(SKILL_ROOT).as_posix()
                          for p in (SKILL_ROOT / "assets").glob("*/check_*.py"))
        self.assertEqual(CLI._gate_assets(SKILL_ROOT), esperado)
        self.assertIn(E2E, CLI._gate_assets(SKILL_ROOT))

    def test_gate_distribuido_y_no_invocado_avisa_con_su_invariante(self):
        with tempfile.TemporaryDirectory() as root:
            repo = _consumer(root)
            aviso = _warn(repo)
            self.assertIn("check_e2e_evidence.py", aviso)
            self.assertIn("HR-58", aviso)

    def test_gate_cableado_en_un_workflow_no_avisa(self):
        with tempfile.TemporaryDirectory() as root:
            repo = _consumer(root, workflow=WIRED)
            self.assertEqual(_warn(repo), "")

    def test_gate_no_distribuido_no_avisa(self):
        with tempfile.TemporaryDirectory() as root:
            repo = _consumer(root, gates=())
            self.assertEqual(_warn(repo), "")

    def test_wrapper_de_la_misma_carpeta_cuenta_como_cableado(self):
        # APAP_WEB (#1742): el workflow llama a run_docs_links.py, que ejecuta
        # check_docs_links.py de la misma carpeta; el nombre del gate no aparece.
        wired = ("name: ci\non: pull_request\njobs:\n  docs:\n    steps:\n"
                 "      - run: python3 .agents/skills/ci-pattern/assets/docs-links/"
                 "run_docs_links.py\n")
        with tempfile.TemporaryDirectory() as root:
            repo = _consumer(root, gates=("assets/docs-links/check_docs_links.py",),
                             workflow=wired)
            self.assertEqual(_warn(repo), "")

    def test_copia_adoptada_en_scripts_cuenta_como_cableada(self):
        # APAP_WEB (#1742): ci.yml ejecuta scripts/docs-decisions/run_docs_decisions.py,
        # una copia adoptada del gate en scripts/, no el asset vendorizado.
        wired = ("name: ci\non: pull_request\njobs:\n  docs:\n    steps:\n"
                 "      - run: python scripts/docs-decisions/run_docs_decisions.py .\n")
        with tempfile.TemporaryDirectory() as root:
            repo = _consumer(root, gates=("assets/docs-decisions/check_docs_decisions.py",),
                             workflow=wired)
            self.assertEqual(_warn(repo), "")

    def test_copia_renombrada_que_cita_el_asset_cuenta_como_cableada(self):
        # APAP_WEB (#1742, APAP_WEB#4215): ci.yml ejecuta
        # scripts/check_e2e_evidence_gate.py, un adaptador renombrado que importa
        # el asset vendorizado; el nombre del gate no aparece en el workflow.
        adaptador = ("#!/usr/bin/env python3\n"
                     "from pathlib import Path\n"
                     "ASSET = Path('.agents/skills/ci-pattern/assets/e2e-evidence/"
                     "check_e2e_evidence.py')\n")
        wired = ("name: ci\non: pull_request\njobs:\n  gates:\n    steps:\n"
                 "      - run: python scripts/check_e2e_evidence_gate.py --repo .\n")
        with tempfile.TemporaryDirectory() as root:
            repo = _consumer(root, workflow=wired)
            _write(repo / "scripts" / "check_e2e_evidence_gate.py", adaptador)
            self.assertEqual(_warn(repo), "")

    def test_copia_renombrada_del_contenido_del_asset_cuenta_como_cableada(self):
        # El asset copiado byte a byte con otro nombre tambien cablea el gate.
        lineas = [f"linea_{i} = {i}" for i in range(20)]
        copia = ["# copia renombrada"] + lineas[1:]
        wired = ("name: ci\non: pull_request\njobs:\n  gates:\n    steps:\n"
                 "      - run: python scripts/check_e2e_evidence_gate.py\n")
        with tempfile.TemporaryDirectory() as root:
            repo = _consumer(root, workflow=wired)
            _write(repo / ".agents" / "skills" / "ci-pattern" / E2E,
                   "\n".join(lineas) + "\n")
            _write(repo / "scripts" / "check_e2e_evidence_gate.py",
                   "\n".join(copia) + "\n")
            self.assertEqual(_warn(repo), "")

    def test_fichero_ajeno_con_nombre_parecido_no_cuenta_como_cableado(self):
        # Un script cuyo nombre se parece al del gate pero que no lo cita no
        # cablea el gate: el aviso sigue.
        wired = ("name: ci\non: pull_request\njobs:\n  gates:\n    steps:\n"
                 "      - run: python scripts/check_e2e_evidence_helper.py\n")
        with tempfile.TemporaryDirectory() as root:
            repo = _consumer(root, workflow=wired)
            _write(repo / "scripts" / "check_e2e_evidence_helper.py",
                   "#!/usr/bin/env python3\nprint('otro script')\n")
            aviso = _warn(repo)
            self.assertIn("check_e2e_evidence.py", aviso)
            self.assertIn("HR-58", aviso)

    def test_gate_que_invoca_la_propia_cli_no_avisa(self):
        # host-readback lo ejecuta `ci-pattern adoption check` (ADOPTION_GATES).
        with tempfile.TemporaryDirectory() as root:
            repo = _consumer(root, gates=("assets/host-readback/check_host_drift.py",))
            self.assertEqual(_warn(repo), "")


class HrCitingGateTests(unittest.TestCase):
    """#1756: la invariante de un gate esta en su bloque HR, no solo en su primera linea.

    La cita del asset puede caer en la continuacion de una entrada HR larga o
    en una sub-regla indentada (HR-54.5): el resolvedor tiene que leer el
    bloque entero. Lo que ya resolvia la primera linea no cambia.
    """

    def test_hr_citada_en_la_continuacion_del_bloque_se_resuelve(self):
        self.assertEqual(
            CLI._hr_citing_gate(SKILL_ROOT, "assets/exemptions/check_exemptions.py"),
            "HR-33",
        )
        self.assertEqual(
            CLI._hr_citing_gate(SKILL_ROOT, "assets/local-ci-parity/check_local_ci_parity.py"),
            "HR-4",
        )
        self.assertEqual(
            CLI._hr_citing_gate(SKILL_ROOT, "assets/ratchet/check_ratchet.py"),
            "HR-15",
        )

    def test_hr_de_una_subregla_indentada_se_resuelve(self):
        self.assertEqual(
            CLI._hr_citing_gate(SKILL_ROOT, "assets/issue-spec/check_issue_spec.py"),
            "HR-54.5",
        )
        self.assertEqual(
            CLI._hr_citing_gate(SKILL_ROOT, "assets/pr-size/check_pr_size.py"),
            "HR-54.11",
        )
        self.assertEqual(
            CLI._hr_citing_gate(SKILL_ROOT, "assets/runner-controls/check_runner_binding.py"),
            "HR-54.15",
        )

    def test_lo_que_resolvia_la_primera_linea_no_cambia(self):
        self.assertEqual(
            CLI._hr_citing_gate(SKILL_ROOT, "assets/workflow-policy/check_workflow_policy.py"),
            "HR-73",
        )
        self.assertEqual(
            CLI._hr_citing_gate(SKILL_ROOT, "assets/e2e-evidence/check_e2e_evidence.py"),
            "HR-58",
        )
        self.assertEqual(
            CLI._hr_citing_gate(SKILL_ROOT, "assets/pr-contract/check_pr_contract.py"),
            "HR-53",
        )

    def test_gate_sin_ninguna_cita_no_inventa_hr(self):
        # Ninguna entrada HR cita este asset: el aviso lo dice sin inventar
        # invariante (las HR que faltan son el paso 4 de #1755).
        self.assertIsNone(
            CLI._hr_citing_gate(SKILL_ROOT, "assets/where-to-work/check_where_to_work.py")
        )


if __name__ == "__main__":
    unittest.main()
