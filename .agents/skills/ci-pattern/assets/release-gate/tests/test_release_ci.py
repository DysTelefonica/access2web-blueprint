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
import os
import subprocess
import sys
import tempfile
import unittest
import unittest.mock
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

    def run_gate(self, fake, tag="v1.0.1", workflow="ci.yml", argv=None,
                 env=None):
        buf = io.StringIO()
        if argv is None:
            argv = ["--repo", str(self.repo.d), "--tag", tag,
                    "--workflow", workflow]
        environ = {"GITHUB_REPOSITORY": "acme/app"} if env is None else env
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf), \
                unittest.mock.patch.dict(os.environ, environ, clear=False):
            if env is not None and "GITHUB_REPOSITORY" not in env:
                os.environ.pop("GITHUB_REPOSITORY", None)
            code = gate.main(argv, fetch_runs=fake)
        return code, buf.getvalue()


class TestSujetoDelGate(ReleaseCiCase):
    """`--repo` nombra dos cosas: el repositorio git local (para el commit del
    tag) y el OWNER/NAME del host (para la API). Con un solo valor, una de las
    dos llamadas era imposible y el gate salía siempre con 2."""

    def test_repo_owner_name_pregunta_a_la_api_por_ese_repo_y_lee_git_de_git_dir(self):
        fake = FakeRuns([run(42, "2026-10-22T10:00:00Z")])
        code, out = self.run_gate(fake, argv=["--repo", "acme/app", "--git-dir", str(self.repo.d),
                                              "--tag", "v1.0.1"], env={})
        self.assertEqual(code, 0, out)
        self.assertEqual(fake.calls, [("acme/app", "ci.yml", self.repo.sha)])

    def test_repo_directorio_toma_owner_name_de_github_repository(self):
        fake = FakeRuns([run(42, "2026-10-22T10:00:00Z")])
        code, out = self.run_gate(fake, env={"GITHUB_REPOSITORY": "acme/app"})
        self.assertEqual(code, 0, out)
        self.assertEqual(fake.calls[0][0], "acme/app")

    def test_repo_directorio_sin_github_repository_sale_dos_sin_llamar_a_la_api(self):
        fake = FakeRuns([run(42, "2026-10-22T10:00:00Z")])
        code, out = self.run_gate(fake, env={})
        self.assertEqual(code, 2, out)
        self.assertIn("GITHUB_REPOSITORY", out)
        self.assertEqual(fake.calls, [])


class TestRunDelPushDelTag(ReleaseCiCase):
    """`--tag-run`: solo cuenta el run que lanzó el push del propio tag. Otro run
    sobre el mismo commit (el push a main, un dispatch) puede haber saltado las
    suites y no prueba lo que se publica."""

    def tag_run(self, ident, created_at, conclusion="success"):
        r = run(ident, created_at, conclusion=conclusion)
        r.update({"event": "push", "head_branch": "v1.0.1"})
        return r

    def main_run(self, ident, created_at):
        r = run(ident, created_at)
        r.update({"event": "push", "head_branch": "main"})
        return r

    def argv(self):
        return ["--repo", str(self.repo.d), "--tag", "v1.0.1", "--tag-run"]

    def test_el_run_del_tag_en_verde_pasa_aunque_haya_uno_mas_nuevo_de_main(self):
        fake = FakeRuns([self.tag_run(1, "2026-10-22T10:00:00Z"),
                         self.main_run(2, "2026-10-22T11:00:00Z")])
        code, out = self.run_gate(fake, argv=self.argv())
        self.assertEqual(code, 0, out)
        self.assertIn("id 1", out)

    def test_sin_run_del_push_del_tag_sale_dos(self):
        fake = FakeRuns([self.main_run(2, "2026-10-22T11:00:00Z")])
        code, out = self.run_gate(fake, argv=self.argv())
        self.assertEqual(code, 2, out)
        self.assertIn("push del tag", out)

    def test_el_run_del_tag_en_rojo_es_hallazgo(self):
        fake = FakeRuns([self.tag_run(1, "2026-10-22T10:00:00Z", conclusion="failure")])
        code, out = self.run_gate(fake, argv=self.argv())
        self.assertEqual(code, 1, out)

    def test_sin_tag_run_se_mide_el_run_mas_nuevo_como_antes(self):
        fake = FakeRuns([self.tag_run(1, "2026-10-22T10:00:00Z", conclusion="failure"),
                         self.main_run(2, "2026-10-22T11:00:00Z")])
        code, out = self.run_gate(fake)
        self.assertEqual(code, 0, out)
        self.assertIn("id 2", out)


class SequenceRuns:
    """Doble que devuelve una lista distinta en cada consulta: el host visto
    mientras el run del CI avanza."""

    def __init__(self, *answers):
        self.answers = list(answers)
        self.calls = 0

    def __call__(self, repo, workflow_ref, sha):
        answer = self.answers[min(self.calls, len(self.answers) - 1)]
        self.calls += 1
        if isinstance(answer, Exception):
            raise answer
        return list(answer)


class FakeClock:
    """Reloj y `sleep` falsos: dormir avanza el reloj, sin esperar de verdad."""

    def __init__(self):
        self.now = 0.0
        self.sleeps = []

    def clock(self):
        return self.now

    def sleep(self, seconds):
        self.sleeps.append(seconds)
        self.now += seconds


class TestEsperaAlRunDelTag(ReleaseCiCase):
    """`--wait-minutes`: el release y el CI del tag arrancan con el mismo push,
    así que el preflight encuentra el CI en curso. El gate espera a que termine
    en vez de dar rojo por llegar antes (#1451)."""

    def tag_run(self, ident, status="completed", conclusion="success"):
        r = run(ident, "2026-10-22T10:00:00Z", status=status, conclusion=conclusion)
        r.update({"event": "push", "head_branch": "v1.0.1"})
        return r

    def run_waiting(self, fake, *extra):
        clock = FakeClock()
        argv = ["--repo", str(self.repo.d), "--tag", "v1.0.1", *extra]
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf), \
                unittest.mock.patch.dict(os.environ, {"GITHUB_REPOSITORY": "acme/app"}):
            code = gate.main(argv, fetch_runs=fake, sleep=clock.sleep, clock=clock.clock)
        return code, buf.getvalue(), clock

    def test_espera_a_que_el_run_en_curso_termine_en_verde(self):
        fake = SequenceRuns([self.tag_run(1, status="in_progress", conclusion=None)],
                            [self.tag_run(1)])
        code, out, clock = self.run_waiting(fake, "--tag-run", "--wait-minutes", "30")
        self.assertEqual(code, 0, out)
        self.assertEqual(fake.calls, 2)
        self.assertEqual(len(clock.sleeps), 1)

    def test_espera_a_que_aparezca_el_run_del_push_del_tag(self):
        fake = SequenceRuns([], [self.tag_run(1, status="queued", conclusion=None)],
                            [self.tag_run(1)])
        code, out, _clock = self.run_waiting(fake, "--tag-run", "--wait-minutes", "30")
        self.assertEqual(code, 0, out)
        self.assertEqual(fake.calls, 3)

    def test_plazo_vencido_con_el_run_en_curso_es_hallazgo(self):
        fake = SequenceRuns([self.tag_run(1, status="in_progress", conclusion=None)])
        code, out, clock = self.run_waiting(fake, "--tag-run", "--wait-minutes", "5")
        self.assertEqual(code, 1, out)
        self.assertIn("in_progress", out)
        self.assertEqual(clock.now, 5 * 60)

    def test_plazo_vencido_sin_run_del_push_del_tag_sale_dos(self):
        fake = SequenceRuns([])
        code, out, clock = self.run_waiting(fake, "--tag-run", "--wait-minutes", "5")
        self.assertEqual(code, 2, out)
        self.assertIn("push del tag", out)
        self.assertEqual(clock.now, 5 * 60)

    def test_api_caida_no_espera(self):
        fake = SequenceRuns(gate.ApiUnavailableError("500"))
        code, out, clock = self.run_waiting(fake, "--tag-run", "--wait-minutes", "30")
        self.assertEqual(code, 2, out)
        self.assertIn("500", out)
        self.assertEqual(clock.sleeps, [])

    def test_sin_tag_run_espera_al_run_mas_nuevo_en_curso(self):
        fake = SequenceRuns([run(1, "2026-10-22T10:00:00Z", status="in_progress",
                                 conclusion=None)],
                            [run(1, "2026-10-22T10:00:00Z")])
        code, out, _clock = self.run_waiting(fake, "--wait-minutes", "30")
        self.assertEqual(code, 0, out)
        self.assertEqual(fake.calls, 2)

    def test_un_run_en_rojo_no_espera(self):
        fake = SequenceRuns([self.tag_run(1, conclusion="failure")])
        code, out, clock = self.run_waiting(fake, "--tag-run", "--wait-minutes", "30")
        self.assertEqual(code, 1, out)
        self.assertEqual(fake.calls, 1)
        self.assertEqual(clock.sleeps, [])

    def test_sin_wait_minutes_el_run_en_curso_no_espera(self):
        fake = SequenceRuns([self.tag_run(1, status="in_progress", conclusion=None)],
                            [self.tag_run(1)])
        code, out, clock = self.run_waiting(fake, "--tag-run")
        self.assertEqual(code, 1, out)
        self.assertEqual(fake.calls, 1)
        self.assertEqual(clock.sleeps, [])

    def test_wait_minutes_negativo_o_no_finito_sale_dos(self):
        for value in ("-1", "inf", "nan"):
            with self.subTest(value=value):
                fake = SequenceRuns([self.tag_run(1)])
                buf = io.StringIO()
                with self.assertRaises(SystemExit) as ctx,                         contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
                    gate.main(["--repo", str(self.repo.d), "--tag", "v1.0.1",
                               "--wait-minutes", value], fetch_runs=fake)
                self.assertEqual(ctx.exception.code, 2)
                self.assertEqual(fake.calls, 0)


class TestCiEnVerde(ReleaseCiCase):
    def test_ci_en_verde_sobre_el_sha_del_tag_pasa(self):
        fake = FakeRuns([run(42, "2026-10-22T10:00:00Z")])
        code, out = self.run_gate(fake)
        self.assertEqual(code, 0, out)
        self.assertIn("42", out)
        # La API recibe el OWNER/NAME del host (GITHUB_REPOSITORY), no la ruta
        # del checkout: con la ruta, `/repos/<ruta>/...` no resuelve.
        self.assertEqual(fake.calls, [("acme/app", "ci.yml", self.repo.sha)])

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
