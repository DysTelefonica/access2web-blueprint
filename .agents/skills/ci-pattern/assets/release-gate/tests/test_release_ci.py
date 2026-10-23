#!/usr/bin/env python3
# ci-pattern asset tests — release gate: CI en verde sobre el commit del tag (#1015)
"""Suite del gate de CI que acompaña al corte de release (HR-19).

El workflow de release y el de CI son independientes: los dos se disparan
sobre el mismo commit y ninguno espera al otro, así que un tag puede
publicar con el CI en rojo sobre los bytes que se firman. Este gate mide esa
condición; el de ``check_release_tag.py`` mide el tag y no el CI.

Ejecuta el gate contra un repositorio git temporal y un **doble del cliente
de la API** (sin red): el doble recibe ``(repo, workflow_ref, sha)`` y
devuelve la lista de runs que el host devolvería, o lanza el error que el
cliente real lanzaría.

Códigos de salida: 0 CI en verde · 1 el último run del SHA del tag no terminó
en ``success`` · 2 sujeto no medible (no hay run para ese SHA, la API no
responde, el tag no existe).
"""
import contextlib
import importlib
import io
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
GATE = ASSET / "check_release_ci.py"
if str(ASSET) not in sys.path:
    sys.path.insert(0, str(ASSET))
gate = importlib.import_module("check_release_ci")


def git(cwd, *args):
    return subprocess.run(["git", "-C", str(cwd), *args],
                          capture_output=True, text=True, check=True)


class Repo:
    """Repositorio git local mínimo; sin red."""

    def __init__(self, d):
        self.d = Path(d)
        git(self.d, "init", "-q", "-b", "main")
        git(self.d, "config", "user.name", "fixture")
        git(self.d, "config", "user.email", "fixture@example.test")
        (self.d / "f.txt").write_text("c", encoding="utf-8")
        git(self.d, "add", "-A")
        git(self.d, "commit", "-q", "-m", "primer commit")
        git(self.d, "tag", "-a", "v1.0.0", "-m", "v1.0.0")
        (self.d / "f.txt").write_text("c2", encoding="utf-8")
        git(self.d, "add", "-A")
        git(self.d, "commit", "-q", "-m", "hotfix")
        git(self.d, "tag", "-a", "v1.0.1", "-m", "v1.0.1")
        self.sha = git(self.d, "rev-parse", "v1.0.1^{}").stdout.strip()


class FakeRuns:
    """Doble del cliente de la API: registra las llamadas y devuelve runs fijas."""

    def __init__(self, runs=(), error=None):
        self.runs = list(runs)
        self.error = error
        self.calls = []

    def __call__(self, repo, workflow_ref, sha):
        self.calls.append((repo, workflow_ref, sha))
        if self.error is not None:
            raise self.error
        return list(self.runs)


def run(ident, created_at, status="completed", conclusion="success"):
    return {"id": ident, "created_at": created_at, "status": status,
            "conclusion": conclusion,
            "html_url": f"https://example.test/runs/{ident}"}


class ReleaseCiCase(unittest.TestCase):
    """Fixture compartida: repositorio git temporal y llamada al gate."""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.repo = Repo(self.tmp.name)

    def run_gate(self, fake, tag="v1.0.1", workflow="ci.yml"):
        buf = io.StringIO()
        argv = ["--repo", str(self.repo.d), "--tag", tag,
                "--workflow", workflow]
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            code = gate.main(argv, fetch_runs=fake)
        return code, buf.getvalue()


class TestCiEnVerde(ReleaseCiCase):
    def test_ci_en_verde_sobre_el_sha_del_tag_pasa(self):
        fake = FakeRuns([run(42, "2026-10-22T10:00:00Z")])
        code, out = self.run_gate(fake)
        self.assertEqual(code, 0, out)
        self.assertIn("42", out)
        self.assertEqual(fake.calls, [(str(self.repo.d), "ci.yml", self.repo.sha)])

    def test_toma_el_run_mas_nuevo_por_fecha(self):
        fake = FakeRuns([run(1, "2026-10-21T09:00:00Z", conclusion="failure"),
                         run(2, "2026-10-22T09:00:00Z")])
        code, out = self.run_gate(fake)
        self.assertEqual(code, 0, out)
        self.assertIn("2", out)

    def test_a_fecha_igual_desempata_por_id(self):
        fake = FakeRuns([run(3, "2026-10-22T09:00:00Z", conclusion="failure"),
                         run(4, "2026-10-22T09:00:00Z")])
        code, out = self.run_gate(fake)
        self.assertEqual(code, 0, out)
        self.assertIn("4", out)


class TestCiNoVerde(ReleaseCiCase):
    def test_conclusion_rojo_es_hallazgo(self):
        fake = FakeRuns([run(7, "2026-10-22T10:00:00Z", conclusion="failure")])
        code, out = self.run_gate(fake)
        self.assertEqual(code, 1, out)
        self.assertIn("failure", out)
        self.assertIn("https://example.test/runs/7", out)

    def test_run_en_curso_es_hallazgo(self):
        fake = FakeRuns([run(8, "2026-10-22T10:00:00Z", status="in_progress",
                             conclusion=None)])
        code, out = self.run_gate(fake)
        self.assertEqual(code, 1, out)
        self.assertIn("in_progress", out)


class TestSujetoNoMedible(ReleaseCiCase):
    def test_sin_runs_para_el_sha_del_tag_sale_dos(self):
        code, out = self.run_gate(FakeRuns([]))
        self.assertEqual(code, 2, out)
        self.assertIn("v1.0.1", out)

    def test_api_que_no_responde_sale_dos(self):
        code, out = self.run_gate(FakeRuns(error=gate.ApiUnavailableError("500")))
        self.assertEqual(code, 2, out)
        self.assertIn("500", out)

    def test_tag_inexistente_sale_dos_sin_llamar_a_la_api(self):
        fake = FakeRuns([run(9, "2026-10-22T10:00:00Z")])
        code, out = self.run_gate(fake, tag="v9.9.9")
        self.assertEqual(code, 2, out)
        self.assertEqual(fake.calls, [])


if __name__ == "__main__":
    unittest.main()
