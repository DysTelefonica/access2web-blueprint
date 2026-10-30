#!/usr/bin/env python3
"""Suite RED/GREEN del gate portátil de presupuesto de revisión (#623).

Cubre la regla de BP portada como activo de ci-pattern:
``VENDORED_ROOTS = (".agents/skills/",)`` queda fuera del presupuesto de
400 líneas, pero un renombrado desde una ruta de autoría hacia una
vendorizada sí cuenta, porque el trabajo fue authored.

``skills/`` **no** está en la lista desde #1088: en el canon es el catálogo
de autoría (64 skills escritas aquí), no contenido vendorizado. Excluirlo
medía 0 líneas donde el revisor tenía 569 delante.

Convenciones: stdlib exclusivamente, ``unittest``, fixtures con
``subprocess`` real sobre repos ``git init`` en ``tmp_path``. Sin
dependencias de terceros; el gate se importa directamente para ejercitar
sus funciones puras, y se invoca por CLI (``subprocess``) en un test
de extremo a extremo.
"""

from __future__ import annotations

import importlib.util
import os
import subprocess
import sys
import unittest
from pathlib import Path

GATE = Path(__file__).resolve().parents[1] / "check_pr_size.py"


def _load_gate():
    spec = importlib.util.spec_from_file_location("check_pr_size", GATE)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _clean_env() -> dict[str, str]:
    """Entorno sin variables ``GIT_*`` heredadas (``GIT_DIR``, ``GIT_INDEX_FILE``...).

    Un hook o un worktree puede exportarlas y redirigir git a otro repo: el
    fixture tiene que ser hermético respecto al entorno del proceso padre.
    """
    return {k: v for k, v in os.environ.items() if not k.startswith("GIT_")}


def _git(*args: str, cwd: Path) -> str:
    env = _clean_env()
    env.update({
        "GIT_AUTHOR_NAME": "t",
        "GIT_AUTHOR_EMAIL": "t@t",
        "GIT_COMMITTER_NAME": "t",
        "GIT_COMMITTER_EMAIL": "t@t",
    })
    out = subprocess.run(
        ["git", *args],
        cwd=cwd, capture_output=True, text=True, check=True, env=env,
    )
    return out.stdout.strip()


def _make_repo(tmp: Path, base_branch: str = "main") -> Path:
    repo = tmp / "repo"
    repo.mkdir(parents=True)
    _git("init", "-q", "-b", base_branch, cwd=repo)
    _git("config", "user.email", "t@t", cwd=repo)
    _git("config", "user.name", "t", cwd=repo)
    (repo / "README.md").write_text("base\n", encoding="utf-8")
    _git("add", "-A", cwd=repo)
    _git("commit", "-qm", "chore: base", cwd=repo)
    # Park subsequent commits on a feature branch so ``--base-ref main``
    # still resolves to the base commit even after test commits land.
    _git("checkout", "-q", "-b", "feature/test", cwd=repo)
    return repo


def _commit_file(repo: Path, path: str, content: str, msg: str) -> None:
    target = repo / path
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(content, encoding="utf-8")
    _git("add", "-A", cwd=repo)
    _git("commit", "-qm", msg, cwd=repo)


def _commit_files(repo: Path, files: dict[str, str], msg: str) -> None:
    """Varios ficheros en el MISMO commit: un bump mueve el lockfile y el manifiesto."""
    for path, content in files.items():
        target = repo / path
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(content, encoding="utf-8")
    _git("add", "-A", cwd=repo)
    _git("commit", "-qm", msg, cwd=repo)


def _rename_and_modify(repo: Path, src: str, dst: str, new_content: str, msg: str) -> None:
    dst_p = repo / dst
    dst_p.parent.mkdir(parents=True, exist_ok=True)
    _git("mv", src, dst, cwd=repo)
    dst_p.write_text(new_content, encoding="utf-8")
    _git("add", "-A", cwd=repo)
    _git("commit", "-qm", msg, cwd=repo)


# ---------------------------------------------------------------------------
# Tests unitarios de las funciones puras del gate
# ---------------------------------------------------------------------------


class TestVendoredRootsRule(unittest.TestCase):
    """RED: la regla VENDORED_ROOTS debe excluir el subtree vendorizado."""

    def test_vendored_path_under_agents_skills_is_recognised(self) -> None:
        gate = _load_gate()
        self.assertTrue(gate.is_under_vendored(".agents/skills/foo/auth.py"))
        # Con prefijos relativos al repo
        self.assertTrue(gate.is_under_vendored(".agents/skills/x.py"))

    def test_the_catalog_is_not_vendored(self) -> None:
        """RED (#1088): ``skills/`` es el catálogo de autoría del canon, cuenta."""
        gate = _load_gate()
        self.assertFalse(gate.is_under_vendored("skills/legacy/skill/SKILL.md"))
        self.assertFalse(gate.is_under_vendored("skills/x.py"))

    def test_non_vendored_path_is_not_vendored(self) -> None:
        gate = _load_gate()
        self.assertFalse(gate.is_under_vendored("app/auth.py"))
        self.assertFalse(gate.is_under_vendored("src/foo.py"))
        # Substring que parece vendorizado pero no lo es (prefijo "skillsx/")
        self.assertFalse(gate.is_under_vendored("skillsx/foo.py"))
        # Substring que parece vendorizado pero no lo es (prefijo ".agents/")
        self.assertFalse(gate.is_under_vendored(".agents/other/auth.py"))

    def test_vendored_constant_matches_issue(self) -> None:
        gate = _load_gate()
        # #1088: el canon no excluye su propio catálogo. Solo queda fuera el
        # árbol vendorizado de los consumers (.agents/skills/).
        self.assertEqual(gate.VENDORED_ROOTS, (".agents/skills/",))


class TestCountsToBudgetRule(unittest.TestCase):
    """RED: el conteo aplica la regla de exclusión + rename-cuenta."""

    def test_authored_path_counts(self) -> None:
        gate = _load_gate()
        self.assertTrue(
            gate.counts_to_budget("app/auth.py", "app/auth.py"),
            "un fichero authored sin renombrar debe contar",
        )

    def test_vendored_new_path_does_not_count_when_no_rename(self) -> None:
        gate = _load_gate()
        self.assertFalse(
            gate.counts_to_budget(".agents/skills/foo/x.py", ".agents/skills/foo/x.py"),
            "un fichero vendorizado nuevo NO debe contar",
        )

    def test_catalog_path_counts_as_authored(self) -> None:
        """RED (#1088): bajo ``skills/`` es autoría y cuenta para el presupuesto."""
        gate = _load_gate()
        self.assertTrue(
            gate.counts_to_budget("skills/legacy/x.py", "skills/legacy/x.py"),
            "skills/<x> sin renombrar DEBE contar: es el catálogo de autoría",
        )

    def test_rename_from_authored_to_vendored_still_counts(self) -> None:
        """REGLA BP: renombrar authored a vendored sigue contando."""
        gate = _load_gate()
        self.assertTrue(
            gate.counts_to_budget("app/auth.py", ".agents/skills/foo/auth.py"),
            "renombrar app/auth.py -> .agents/skills/... debe CONTAR "
            "(el trabajo fue authored fuera del subtree vendorizado)",
        )
        self.assertTrue(
            gate.counts_to_budget("app/auth.py", "skills/legacy/auth.py"),
            "renombrar app/auth.py -> skills/... debe CONTAR",
        )

    def test_rename_within_vendored_does_not_count(self) -> None:
        gate = _load_gate()
        self.assertFalse(
            gate.counts_to_budget(
                ".agents/skills/old/x.py", ".agents/skills/new/x.py",
            ),
            "renombrar dentro de vendored NO debe contar",
        )


# ---------------------------------------------------------------------------
# Tests de integración sobre un repo git real
# ---------------------------------------------------------------------------


class TestGeneratedLockfileRule(unittest.TestCase):
    """RED (#1177): el gate publica la regla de lockfile generado."""

    def test_the_gate_exposes_the_lockfile_rule(self) -> None:
        gate = _load_gate()
        self.assertTrue(gate.is_generated_lockfile("uv.lock"))
        self.assertTrue(gate.is_generated_lockfile("app/requirements-dev.lock"))
        self.assertFalse(gate.is_generated_lockfile("app/requirements.txt"))

    def test_the_documented_patterns_are_the_ones_in_the_gate(self) -> None:
        gate = _load_gate()
        self.assertEqual(
            gate.LOCKFILE_PATTERNS,
            ("uv.lock", "package-lock.json", "pnpm-lock.yaml", "poetry.lock",
             "requirements*.lock"),
        )


class TestChangedLinesIntegration(unittest.TestCase):
    """RED: la integración con ``git diff --numstat --find-renames``
    implementa la regla end-to-end sobre un repo de fixture."""

    def setUp(self) -> None:
        import tempfile
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.repo = _make_repo(self.tmp)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _run_gate(self, *args: str, env_extra: dict[str, str] | None = None) -> subprocess.CompletedProcess:
        env = _clean_env()
        if env_extra:
            env.update(env_extra)
        return subprocess.run(
            [sys.executable, str(GATE), *args],
            cwd=self.repo, capture_output=True, text=True, env=env,
        )

    def test_vendored_change_is_excluded_from_budget(self) -> None:
        """RED: 1000 líneas en .agents/skills/ NO cuentan → PASS."""
        body = "\n".join(f"# line {i}" for i in range(1000)) + "\n"
        _commit_file(
            self.repo, ".agents/skills/foo/x.py", body,
            "feat: vendor a thousand lines of skill under .agents/skills/",
        )
        result = self._run_gate("--base-ref", "HEAD~1")
        self.assertEqual(
            result.returncode, 0,
            f"el gate debe pasar (excluye vendored), pero devolvió "
            f"{result.returncode}.\nstdout: {result.stdout}\n"
            f"stderr: {result.stderr}",
        )
        self.assertIn("0 changed line(s)", result.stdout)

    def test_catalog_change_counts_toward_the_budget(self) -> None:
        """RED (#1088): N líneas añadidas bajo ``skills/`` cuentan como N.

        El catálogo del canon es el producto del repo: excluirlo del
        presupuesto medía 0 líneas donde el revisor tenía 569 delante.
        """
        body = "\n".join(f"# line {i}" for i in range(569)) + "\n"
        _commit_file(
            self.repo, "skills/acme/SKILL.md", body,
            "feat: write a 569-line skill",
        )
        result = self._run_gate("--base-ref", "HEAD~1")
        self.assertIn("569 changed line(s)", result.stdout)
        self.assertEqual(
            result.returncode, 1,
            f"569 líneas de catálogo superan el techo de 400: el gate debe "
            f"fallar.\nstdout: {result.stdout}\nstderr: {result.stderr}",
        )

    def test_catalog_and_vendored_are_counted_apart(self) -> None:
        """RED (#1088): el catálogo cuenta; lo vendorizado sigue fuera."""
        _commit_file(
            self.repo, "skills/acme/SKILL.md",
            "\n".join(f"# s {i}" for i in range(300)) + "\n",
            "feat: write a skill",
        )
        _commit_file(
            self.repo, ".agents/skills/foo/x.py",
            "\n".join(f"# v {i}" for i in range(1000)) + "\n",
            "feat: vendor a thousand lines",
        )
        result = self._run_gate("--base-ref", "HEAD~2")
        self.assertIn("300 changed line(s)", result.stdout)
        self.assertIn("1000 vendored line(s) excluded", result.stdout)
        self.assertEqual(
            result.returncode, 0,
            f"300 líneas de catálogo están bajo el techo: el gate debe pasar.\n"
            f"stdout: {result.stdout}\nstderr: {result.stderr}",
        )

    def test_authored_change_over_budget_fails(self) -> None:
        """RED: 500 líneas authored en app/ → FAIL (500 > 400)."""
        body = "\n".join(f"# line {i}" for i in range(500)) + "\n"
        _commit_file(self.repo, "app/x.py", body, "feat: add 500 authored lines")
        result = self._run_gate("--base-ref", "HEAD~1")
        self.assertEqual(
            result.returncode, 1,
            f"el gate debe fallar (500 > 400), pero devolvió "
            f"{result.returncode}.\nstdout: {result.stdout}\n"
            f"stderr: {result.stderr}",
        )
        self.assertIn("500 changed line(s)", result.stdout)

    def test_rename_authored_to_vendored_still_counts(self) -> None:
        """RED: un rename authored→vendored con >400 líneas nuevas debe FALLAR.

        El propagador mueve ``skills/<legacy>/<file>`` a
        ``.agents/skills/<new>/<file>`` con contenido casi idéntico;
        git detecta el rename (``{old => new}/path`` en numstat) y
        la regla BP exige que ese trabajo cuente.
        """
        # 1) Commit el archivo authored en main para que el diff
        #    cumulativo ``main...HEAD`` pueda detectar el rename
        #    (sin esto, el archivo solo existe en HEAD y git no
        #    tiene con qué emparejarlo para emitir ``{old => new}``).
        _git("checkout", "main", cwd=self.repo)
        body_main = "\n".join(["shared"] * 1000) + "\n"
        _commit_file(self.repo, "app/auth.py", body_main, "feat: introduce auth")
        # 2) Rama feature desde main
        _git("checkout", "-q", "-b", "feature/rename", cwd=self.repo)
        # 3) Rename + 500 líneas nuevas (1500 total, similaridad ~67%)
        body_new = "\n".join(["shared"] * 1000 + [f"new {i}" for i in range(500)]) + "\n"
        _rename_and_modify(
            self.repo, "app/auth.py", ".agents/skills/foo/auth.py",
            body_new, "refactor: move auth under vendored skill",
        )
        # Diff cumulativo main..HEAD con --find-renames=50%:
        #   500	0	{app => .agents/skills/foo}/auth.py
        # (similaridad 67% > 50%, git detecta el rename)
        # Mi regla: rename authored→vendored CUENTA → 500 > 400 → FAIL
        result = self._run_gate("--base-ref", "main")
        self.assertEqual(
            result.returncode, 1,
            f"el gate debe fallar (rename authored→vendored con 500 cuenta), "
            f"pero devolvió {result.returncode}.\nstdout: {result.stdout}\n"
            f"stderr: {result.stderr}",
        )
        self.assertIn("500 changed line(s)", result.stdout)

    def test_rename_within_vendored_is_excluded(self) -> None:
        """RED: rename entre dos rutas vendorizadas no cuenta."""
        # 1) En main, crear una skill vendorizada con 300 líneas
        body_main = "\n".join(f"# main {i}" for i in range(300)) + "\n"
        _commit_file(
            self.repo, ".agents/skills/old/x.py", body_main,
            "feat: introduce vendored skill",
        )
        # 2) En feature, renombrarla a otra ruta vendorizada y cambiar
        #    el contenido a 500 líneas
        body_new = "\n".join(f"# new {i}" for i in range(500)) + "\n"
        _rename_and_modify(
            self.repo, ".agents/skills/old/x.py", ".agents/skills/new/x.py",
            body_new, "refactor: rename vendored skill",
        )
        # Con la regla: rename dentro de vendored NO cuenta → PASS
        result = self._run_gate("--base-ref", "main")
        self.assertEqual(
            result.returncode, 0,
            f"el gate debe pasar (rename dentro de vendored no cuenta), "
            f"pero devolvió {result.returncode}.\nstdout: {result.stdout}\n"
            f"stderr: {result.stderr}",
        )

    def test_mixed_changes_count_only_authored(self) -> None:
        """RED: mix de vendored y authored → solo authored cuenta."""
        # 300 líneas authored en app/
        _commit_file(
            self.repo, "app/foo.py",
            "\n".join(f"# a {i}" for i in range(300)) + "\n",
            "feat: add 300 authored",
        )
        # 1000 líneas en .agents/skills/ (no cuentan)
        _commit_file(
            self.repo, ".agents/skills/big/x.py",
            "\n".join(f"# v {i}" for i in range(1000)) + "\n",
            "feat: vendor a big skill",
        )
        # Esperado: solo 300 cuentan → PASS
        result = self._run_gate("--base-ref", "HEAD~2")
        self.assertEqual(
            result.returncode, 0,
            f"el gate debe pasar (300 < 400), pero devolvió "
            f"{result.returncode}.\nstdout: {result.stdout}\n"
            f"stderr: {result.stderr}",
        )
        self.assertIn("300 changed line(s)", result.stdout)


class TestCliExitCodes(unittest.TestCase):
    """RED: contrato de exit codes y de la salida estructurada."""

    def setUp(self) -> None:
        import tempfile
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.repo = _make_repo(self.tmp)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_json_envelope_within_budget(self) -> None:
        _commit_file(
            self.repo, "app/small.py", "x = 1\n",
            "feat: tiny change",
        )
        result = subprocess.run(
            [sys.executable, str(GATE), "--base-ref", "HEAD~1", "--json"],
            cwd=self.repo, capture_output=True, text=True, env=_clean_env(),
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        import json
        env = json.loads(result.stdout)
        self.assertEqual(env["gate"], "pr_size")
        self.assertEqual(env["status"], "pass")
        self.assertEqual(env["indicators"]["changed_lines"], 1)
        self.assertEqual(env["ceilings"]["changed_lines"], 400)

    def test_json_envelope_over_budget(self) -> None:
        body = "\n".join(f"# line {i}" for i in range(500)) + "\n"
        _commit_file(self.repo, "app/x.py", body, "feat: add 500 authored")
        result = subprocess.run(
            [sys.executable, str(GATE), "--base-ref", "HEAD~1", "--json"],
            cwd=self.repo, capture_output=True, text=True, env=_clean_env(),
        )
        self.assertEqual(result.returncode, 1, result.stderr)
        import json
        env = json.loads(result.stdout)
        self.assertEqual(env["status"], "fail")
        self.assertEqual(env["indicators"]["changed_lines"], 500)

    def test_vendored_indicator_in_envelope(self) -> None:
        """La salida indica cuántas líneas fueron excluidas por vendored."""
        # 1000 vendored + 50 authored
        _commit_file(
            self.repo, ".agents/skills/big/x.py",
            "\n".join(f"# v {i}" for i in range(1000)) + "\n",
            "feat: vendor big",
        )
        _commit_file(
            self.repo, "app/small.py",
            "\n".join(f"# a {i}" for i in range(50)) + "\n",
            "feat: add small",
        )
        result = subprocess.run(
            [sys.executable, str(GATE), "--base-ref", "HEAD~2", "--json"],
            cwd=self.repo, capture_output=True, text=True, env=_clean_env(),
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        import json
        env = json.loads(result.stdout)
        # changed_lines es SOLO lo que cuenta al presupuesto (excluye vendored)
        self.assertEqual(env["indicators"]["changed_lines"], 50)
        self.assertEqual(env["indicators"]["vendored_lines"], 1000)


# ---------------------------------------------------------------------------
# #806 C3: la excepcion es solo la etiqueta; el modo API cuenta sin historial
# ---------------------------------------------------------------------------


def _over_budget_repo(tmp: Path, path: str = "app/x.py") -> Path:
    repo = _make_repo(tmp)
    _commit_file(repo, path, "\n".join(f"# line {i}" for i in range(500)) + "\n", "feat: 500 lines")
    return repo


def _run_cli(repo: Path, *extra: str, env: dict[str, str] | None = None):
    full = _clean_env()
    full.update(env or {})
    return subprocess.run([sys.executable, str(GATE), "--base-ref", "HEAD~1", *extra],
                          cwd=repo, capture_output=True, text=True, env=full)


class TestLabelOverride(unittest.TestCase):
    """La unica excepcion es la etiqueta ``size:exception`` que aplica el mantenedor."""

    def setUp(self) -> None:
        import tempfile
        self._tmp = tempfile.TemporaryDirectory()
        self.repo = _over_budget_repo(Path(self._tmp.name))

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_over_budget_without_label_fails_and_names_the_maintainer(self) -> None:
        result = _run_cli(self.repo)
        self.assertEqual(result.returncode, 1, result.stdout)
        self.assertIn("maintainer", result.stdout)
        self.assertIn("size:exception", result.stdout)

    def test_label_flag_overrides(self) -> None:
        result = _run_cli(self.repo, "--labels", "type:feature,size:exception")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("overridden by the 'size:exception' label", result.stdout)

    def test_label_env_overrides(self) -> None:
        result = _run_cli(self.repo, env={"PR_LABELS": "size:exception"})
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_another_label_does_not_override(self) -> None:
        result = _run_cli(self.repo, "--labels", "size:small,type:feature")
        self.assertEqual(result.returncode, 1, result.stdout)

    def test_body_field_is_retired(self) -> None:
        body = "size:exception\nsize-exception-reason: it is fine"
        result = _run_cli(self.repo, env={"PR_BODY": body})
        self.assertEqual(result.returncode, 1, result.stdout)

    def test_json_envelope_reports_the_override(self) -> None:
        import json
        result = _run_cli(self.repo, "--labels", "size:exception", "--json")
        self.assertEqual(result.returncode, 0, result.stderr)
        envelope = json.loads(result.stdout)
        self.assertTrue(envelope["indicators"]["overridden"])
        self.assertEqual(envelope["status"], "pass")


class TestGeneratedLockfilesAreExcluded(unittest.TestCase):
    """RED (#1177): un lockfile generado no es trabajo de revision y no cuenta.

    El lockfile lo escribe el gestor de paquetes a partir del manifiesto
    (``ardelperal/APAP_WEB#2424``: 444 de las 448 lineas del diff eran
    ``uv.lock`` y la PR solo pudo entrar con la etiqueta ``size:exception``).
    """

    def test_a_lockfile_does_not_count_toward_the_budget(self) -> None:
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            repo = _over_budget_repo(Path(tmp), "uv.lock")
            result = _run_cli(repo)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_a_lockfile_is_reported_apart_from_the_vendored_lines(self) -> None:
        """La cifra excluida se dice por separado: no es arbol vendorizado."""
        import json
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            repo = _over_budget_repo(Path(tmp), "uv.lock")
            result = _run_cli(repo, "--json")
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            env = json.loads(result.stdout)
            self.assertEqual(env["status"], "pass")
            self.assertEqual(env["indicators"]["changed_lines"], 0)
            self.assertEqual(env["indicators"]["generated_lines"], 500)
            self.assertEqual(env["indicators"]["vendored_lines"], 0)

    def test_the_lockfile_is_named_in_the_text_output(self) -> None:
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            repo = _over_budget_repo(Path(tmp), "uv.lock")
            result = _run_cli(repo)
            self.assertIn("500 generated line(s) excluded", result.stdout)

    def test_authored_code_still_fails(self) -> None:
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            repo = _over_budget_repo(Path(tmp), "app/x.py")
            result = _run_cli(repo)
            self.assertEqual(result.returncode, 1, result.stdout)
            self.assertIn("maintainer", result.stdout)

    def test_requirements_txt_still_counts(self) -> None:
        """Solo la extension ``.lock`` queda fuera: ``requirements.txt`` es de autoria."""
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            repo = _over_budget_repo(Path(tmp), "app/requirements.txt")
            result = _run_cli(repo)
            self.assertEqual(result.returncode, 1, result.stdout)

    def test_requirements_lock_is_excluded(self) -> None:
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            repo = _over_budget_repo(Path(tmp), "app/requirements-dev.lock")
            result = _run_cli(repo)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)


class TestLockfileBumpRealCase(unittest.TestCase):
    """El caso real de APAP_WEB#2424: un bump que mueve 444 líneas de ``uv.lock``
    y 4 de ``pyproject.toml`` (448 en crudo). El presupuesto cuenta 4, no 448
    (#1177): el lockfile lo escribió el gestor de paquetes, no quien abre la PR.
    """

    def _bump_repo(self, tmp: Path) -> Path:
        repo = _make_repo(tmp)
        lock = "\n".join(f"# lock {i}" for i in range(222)) + "\n"
        manifest = "\n".join(f"# dep {i}" for i in range(2)) + "\n"
        _commit_files(repo, {"uv.lock": lock, "pyproject.toml": manifest},
                      "chore: bump")
        # El bump: el lockfile se reescribe entero (222 altas + 222 bajas) y el
        # manifiesto cambia dos líneas (2 altas + 2 bajas). Los dos ficheros
        # entran en el MISMO commit: es un commit, no dos.
        _commit_files(repo,
                      {"uv.lock": "\n".join(f"# lock {i} v2" for i in range(222)) + "\n",
                       "pyproject.toml": "\n".join(f"# dep {i} v2" for i in range(2)) + "\n"},
                      "chore: bump v2")
        return repo

    def test_the_budget_counts_the_manifest_and_not_the_lockfile(self) -> None:
        import json
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            repo = self._bump_repo(Path(tmp))
            result = _run_cli(repo, "--json")
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            env = json.loads(result.stdout)
            self.assertEqual(env["status"], "pass")
            self.assertEqual(env["indicators"]["changed_lines"], 4)
            self.assertEqual(env["indicators"]["generated_lines"], 444)
            self.assertEqual(env["indicators"]["vendored_lines"], 0)

    def test_the_text_output_says_what_it_left_out(self) -> None:
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            repo = self._bump_repo(Path(tmp))
            result = _run_cli(repo)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertIn("OK    4 changed line(s)", result.stdout)
            self.assertIn("444 generated line(s) excluded", result.stdout)


def _file(name: str, added: int, deleted: int = 0, previous: str | None = None) -> dict:
    item = {"filename": name, "additions": added, "deletions": deleted,
            "status": "renamed" if previous else "modified"}
    if previous:
        item["previous_filename"] = previous
    return item


class _Seq:
    """Respuestas en orden: cada llamada toma la siguiente; la última se repite."""

    def __init__(self, items: list) -> None:
        self.items, self.i = items, 0

    def take(self):
        value = self.items[min(self.i, len(self.items) - 1)]
        self.i += 1
        return value


class _Api:
    """Servidor HTTP minimo en loopback para ``/repos/o/r/pulls/N`` y ``.../files``.

    Aislamiento por instancia: cada ``_Api`` levanta su propio ``HTTPServer`` en puerto
    efímero (``bind 0``) y guarda su estado en la propia instancia (#941, #511, #585).

    - ``pulls``: secuencia de cuerpos de ``GET pulls/N`` (``head.sha``, ``changed_files``).
    - ``listings``: secuencia de listados de ficheros; cada primera página toma el
      siguiente listado y las siguientes páginas siguen el mismo.
    - ``paths``: cada petición de la API, en orden (la sonda ajena no cuenta).
    - ``live_labels``: etiquetas que devuelve ``GET issues/N/labels`` (#1359).
    - ``labels_status``: código de respuesta de ``GET issues/N/labels`` (#1359).
    """

    def __init__(self, files: list[dict], status: int = 200, pulls: list | None = None,
                 listings: list | None = None, live_labels: list | None = None,
                 labels_status: int = 200) -> None:
        import threading
        from http.server import HTTPServer
        self.files, self.status, self.auth, self.paths = files, status, [], []
        self.live_labels, self.labels_status = list(live_labels or []), labels_status
        self.pulls = _Seq(pulls if pulls is not None
                          else [{"head": {"sha": "HEAD1"}, "changed_files": len(files)}])
        self.listings = _Seq(listings if listings is not None else [files])
        self.current: list[dict] = files
        self.server = HTTPServer(("127.0.0.1", 0), _make_handler(self))
        self.url = f"http://127.0.0.1:{self.server.server_address[1]}"
        threading.Thread(target=self.server.serve_forever, daemon=True).start()

    def close(self) -> None:
        self.server.shutdown()
        self.server.server_close()


def _make_handler(api: _Api):
    """Clase Handler ligada a una ``_Api``: enruta ``pulls/N`` y ``pulls/N/files``.

    Cualquier otra URL devuelve 404 explicito (#941)."""
    import json as _json
    import re as _re
    import urllib.parse as _urlparse
    from http.server import BaseHTTPRequestHandler

    class Handler(BaseHTTPRequestHandler):
        def _send(self, status: int, payload: object) -> None:
            body = _json.dumps(payload).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self) -> None:  # noqa: N802
            # Un sondeo ajeno (`GET /` sin Authorization, #941) recibe el 404 de abajo
            # sin entrar en `auth`; el resto de peticiones, tengan o no token, cuentan.
            probe = _urlparse.urlsplit(self.path).path == "/" and self.headers.get("Authorization") is None
            if not probe:
                api.auth.append(self.headers.get("Authorization"))
                api.paths.append(self.path)
            url = _urlparse.urlsplit(self.path)
            query = _urlparse.parse_qs(url.query)
            if api.status != 200:
                self._send(api.status, {"message": "error"})
                return
            if _re.search(r"/issues/\d+/labels$", url.path):
                if api.labels_status != 200:
                    self._send(api.labels_status, {"message": "error"})
                    return
                per_i = int(query.get("per_page", ["100"])[0])
                page_i = int(query.get("page", ["1"])[0])
                self._send(200, [{"name": name} for name in api.live_labels][(page_i - 1) * per_i:page_i * per_i])
                return
            if _re.search(r"/pulls/\d+$", url.path):
                self._send(200, api.pulls.take())
                return
            page, per = query.get("page"), query.get("per_page")
            if not url.path.endswith("/files") or not page or not per:
                self._send(404, {"message": "Not Found"})
                return
            page_i, per_i = int(page[0]), int(per[0])
            if page_i == 1:
                api.current = api.listings.take()
            self._send(200, api.current[(page_i - 1) * per_i:page_i * per_i])

        def log_message(self, *args: object) -> None:
            return

    return Handler


class TestApiMode(unittest.TestCase):
    """``--event``: cuenta con la API de GitHub, sin historial de git."""

    def setUp(self) -> None:
        import tempfile
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.apis: list[_Api] = []

    def tearDown(self) -> None:
        for api in self.apis:
            api.close()
        self._tmp.cleanup()

    def run_api(self, files, labels=(), changed=None, token="tok", status=200, url=None, extra=(),
                head="HEAD1", pulls=None, listings=None, live_labels=None, labels_status=200):
        import json
        # Por defecto las etiquetas vivas son las del payload; un caso de etiqueta añadida
        # o quitada tras el evento pasa `live_labels` distinto (#1359).
        api = _Api(files, status, pulls=pulls, listings=listings,
                   live_labels=labels if live_labels is None else live_labels,
                   labels_status=labels_status)
        self.apis.append(api)
        event = self.tmp / "event.json"
        event.write_text(json.dumps({
            "repository": {"full_name": "o/r"},
            "pull_request": {"number": 7, "labels": [{"name": n} for n in labels],
                             "head": {"sha": head},
                             "changed_files": len(files) if changed is None else changed},
        }), encoding="utf-8")
        env = _clean_env()
        env.pop("GITHUB_TOKEN", None)
        env["GITHUB_API_URL"] = url or api.url
        if token:
            env["GITHUB_TOKEN"] = token
        result = subprocess.run([sys.executable, str(GATE), "--event", str(event), *extra],
                                cwd=self.tmp, capture_output=True, text=True, env=env)
        return result, api

    def test_within_budget_passes_and_sends_the_token(self) -> None:
        result, api = self.run_api([_file("app/a.py", 10, 5)])
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("15 changed line(s)", result.stdout)
        self.assertEqual(set(api.auth), {"Bearer tok"})

    def test_additions_plus_deletions_over_budget_fails(self) -> None:
        result, _ = self.run_api([_file("app/a.py", 300, 150)])
        self.assertEqual(result.returncode, 1, result.stdout)
        self.assertIn("app/a.py", result.stdout)

    def test_vendored_root_is_excluded_and_reported(self) -> None:
        result, _ = self.run_api([_file(".agents/skills/x/a.py", 1000), _file("app/a.py", 50)])
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("1000 vendored line(s) excluded", result.stdout)

    def test_catalog_file_counts_and_is_not_reported_as_vendored(self) -> None:
        """RED (#1088): en modo API el catálogo cuenta como autoría."""
        result, _ = self.run_api([_file("skills/acme/SKILL.md", 500)])
        self.assertEqual(result.returncode, 1, result.stdout)
        self.assertIn("500 changed line(s)", result.stdout)
        self.assertNotIn("vendored", result.stdout)

    def test_rename_from_authored_into_vendored_still_counts(self) -> None:
        result, _ = self.run_api([_file(".agents/skills/x/a.py", 500, previous="app/a.py")])
        self.assertEqual(result.returncode, 1, result.stdout)

    def test_lockfile_does_not_count(self) -> None:
        """RED (#1177): en modo API el lockfile generado tampoco cuenta."""
        result, _ = self.run_api([_file("uv.lock", 500)])
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("500 generated line(s) excluded", result.stdout)

    def test_a_rename_from_authored_into_a_lockfile_still_counts(self) -> None:
        result, _ = self.run_api([_file("uv.lock", 500, previous="app/a.py")])
        self.assertEqual(result.returncode, 1, result.stdout)

    def test_a_lockfile_bump_counts_only_the_manifest(self) -> None:
        """APAP_WEB#2424: 444 líneas de ``uv.lock`` y 4 de ``pyproject.toml`` → 4."""
        result, _ = self.run_api([_file("uv.lock", 222, 222),
                                  _file("pyproject.toml", 2, 2)])
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("4 changed line(s)", result.stdout)
        self.assertIn("444 generated line(s) excluded", result.stdout)

    def test_label_overrides(self) -> None:
        result, _ = self.run_api([_file("app/a.py", 900)], labels=("size:exception",))
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("overridden", result.stdout)

    def test_label_added_after_the_event_overrides(self) -> None:
        """RED (#1359): el payload del `opened` no trae la etiqueta; la API ya la tiene.

        cadete#715: `size:exception` se puso 7 s después de abrir el PR. El run del
        `opened` leyó el snapshot del payload, sin etiqueta, y falló; el run siguiente
        pasó. El gate lee las etiquetas en vivo, no del snapshot del evento.
        """
        result, api = self.run_api([_file("app/a.py", 900)], labels=(),
                                   live_labels=("size:exception",))
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("overridden", result.stdout)
        self.assertTrue(any("/issues/7/labels" in path for path in api.paths))

    def test_a_label_removed_after_the_event_does_not_override(self) -> None:
        """RED (#1359): la etiqueta del payload ya no está en la API: no exime."""
        result, _ = self.run_api([_file("app/a.py", 900)], labels=("size:exception",),
                                 live_labels=())
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)

    def test_labels_api_failure_fails_closed(self) -> None:
        """RED (#1359): sin etiquetas vivas no hay veredicto: falla cerrado, aunque quepa."""
        result, _ = self.run_api([_file("app/a.py", 10)], labels_status=500)
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        self.assertIn("labels", result.stdout + result.stderr)

    def test_a_foreign_probe_is_not_counted(self) -> None:
        """RED (#941): un `GET /` ajeno al gate (sin Authorization) no entra en ``auth``."""
        import urllib.error
        import urllib.request
        api = _Api([_file("app/a.py", 1)])
        self.apis.append(api)
        with self.assertRaises(urllib.error.HTTPError):
            urllib.request.urlopen(f"{api.url}/", timeout=5)
        self.assertEqual(api.auth, [])

    def test_an_unauthenticated_api_request_is_counted(self) -> None:
        """Una peticion sin token a una ruta de API SI cuenta en ``auth`` (#941)."""
        import urllib.request
        api = _Api([_file("app/a.py", 1)])
        self.apis.append(api)
        with urllib.request.urlopen(f"{api.url}/repos/o/r/pulls/7/files?page=1&per_page=100", timeout=5) as resp:
            resp.read()
        self.assertEqual(api.auth, [None])

    def test_pagination_counts_every_page(self) -> None:
        files = [_file(f"app/f{i}.py", 3) for i in range(150)]
        result, api = self.run_api(files)
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        self.assertIn("450 changed line(s)", result.stdout)
        self.assertEqual(len([p for p in api.paths if "/files" in p]), 2)

    def test_breakdown_that_does_not_match_changed_files_fails_closed(self) -> None:
        """Un desfase persistente entre el listado y GET pulls, leidos a la vez, falla cerrado."""
        result, _ = self.run_api([_file("app/a.py", 1)],
                                 pulls=[{"head": {"sha": "HEAD1"}, "changed_files": 3}])
        self.assertEqual(result.returncode, 1)
        self.assertIn("changed_files", result.stderr)

    def test_stale_changed_files_in_the_payload_alone_does_not_fail(self) -> None:
        """#1113: el payload del evento es una foto anterior; no decide por si solo."""
        result, _ = self.run_api([_file("app/a.py", 1)], changed=3)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_api_error_fails_closed(self) -> None:
        result, _ = self.run_api([_file("app/a.py", 1)], status=500)
        self.assertEqual(result.returncode, 1)
        self.assertIn("failed", result.stderr)

    def test_missing_token_fails_closed(self) -> None:
        result, _ = self.run_api([_file("app/a.py", 1)], token="")
        self.assertEqual(result.returncode, 1)
        self.assertIn("GITHUB_TOKEN", result.stderr)

    def test_plain_http_to_a_non_loopback_host_is_refused(self) -> None:
        result, _ = self.run_api([_file("app/a.py", 1)], url="http://example.invalid")
        self.assertEqual(result.returncode, 1)
        self.assertIn("https", result.stderr)

    def test_head_moved_after_the_event_is_superseded_not_an_error(self) -> None:
        """RED (#1113): si el head de la PR ya no es el del evento, sale 0 y lo dice."""
        result, _ = self.run_api(
            [_file("app/a.py", 1)], head="HEAD1",
            pulls=[{"head": {"sha": "HEAD2"}, "changed_files": 1}])
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("superado por un push más reciente", result.stdout)

    def test_breakdown_that_settles_on_the_retry_passes(self) -> None:
        """RED (#1113): un desglose desfasado en el primer intento se reintenta una vez."""
        result, _ = self.run_api(
            [_file("app/a.py", 1), _file("app/b.py", 2)],
            pulls=[{"head": {"sha": "HEAD1"}, "changed_files": 2}],
            listings=[[_file("app/a.py", 1)], [_file("app/a.py", 1), _file("app/b.py", 2)]])
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("3 changed line(s)", result.stdout)

    def test_json_error_envelope_in_api_mode(self) -> None:
        import json
        result, _ = self.run_api([_file("app/a.py", 1)], token="", extra=("--json",))
        self.assertEqual(result.returncode, 1)
        self.assertEqual(json.loads(result.stdout)["status"], "error")


if __name__ == "__main__":
    unittest.main()


# ---------------------------------------------------------------------------
# RED: el modo local mide contra la base real del PR (issue #658)
# ---------------------------------------------------------------------------


class TestLocalBaseResolution(unittest.TestCase):
    """RED: sin ``--base-ref`` ni ``BASE_REF`` el gate local pregunta a
    ``gh`` la base del PR de la rama actual y mide contra ella, como hace
    el ``pr-gate`` de CI con ``github.base_ref``. Sin PR cae a
    ``origin/main``. Una base explícita siempre gana."""

    def setUp(self) -> None:
        import tempfile
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.repo = _make_repo(self.tmp)
        # Cadena: main <- feature/base (300 líneas) <- feature/test (150 líneas).
        _commit_file(self.repo, "app/base_work.py", "x = 1\n" * 300, "feat: base slice")
        base_sha = _git("rev-parse", "HEAD", cwd=self.repo)
        _git("update-ref", "refs/remotes/origin/feature/base", base_sha, cwd=self.repo)
        _commit_file(self.repo, "app/tip_work.py", "y = 2\n" * 150, "feat: tip slice")
        main_sha = _git("rev-parse", "HEAD~2", cwd=self.repo)
        _git("update-ref", "refs/remotes/origin/main", main_sha, cwd=self.repo)
        self.bin = self.tmp / "bin"
        self.bin.mkdir()

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _fake_gh(self, stdout: str, exit_code: int = 0) -> None:
        gh = self.bin / "gh"
        gh.write_text(f"#!/bin/sh\nprintf '%s' '{stdout}'\nexit {exit_code}\n", encoding="utf-8")
        gh.chmod(0o755)

    def _run_gate(self, *args: str, env_extra: dict[str, str] | None = None) -> subprocess.CompletedProcess:
        env = _clean_env()
        env["PATH"] = f"{self.bin}{os.pathsep}{env.get('PATH', '')}"
        env.pop("BASE_REF", None)
        if env_extra:
            env.update(env_extra)
        return subprocess.run(
            [sys.executable, str(GATE), *args],
            cwd=self.repo, capture_output=True, text=True, env=env,
        )

    def test_pr_base_from_gh_counts_only_the_stacked_slice(self) -> None:
        self._fake_gh('{"baseRefName":"feature/base"}')
        result = self._run_gate()
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("origin/feature/base", result.stdout)
        self.assertIn("150 changed line(s)", result.stdout)

    def test_no_pr_falls_back_to_origin_main(self) -> None:
        self._fake_gh("", exit_code=1)
        result = self._run_gate()
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        self.assertIn("origin/main", result.stdout)
        self.assertIn("450 changed line(s)", result.stdout)

    def test_explicit_base_ref_wins_over_gh(self) -> None:
        self._fake_gh('{"baseRefName":"feature/base"}')
        result = self._run_gate("--base-ref", "origin/main")
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        self.assertIn("450 changed line(s)", result.stdout)

    def test_base_ref_env_wins_over_gh(self) -> None:
        self._fake_gh('{"baseRefName":"feature/base"}')
        result = self._run_gate(env_extra={"BASE_REF": "origin/main"})
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        self.assertIn("450 changed line(s)", result.stdout)
