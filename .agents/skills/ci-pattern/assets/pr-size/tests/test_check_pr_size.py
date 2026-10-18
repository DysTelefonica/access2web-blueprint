#!/usr/bin/env python3
"""Suite RED/GREEN del gate portátil de presupuesto de revisión (#623).

Cubre la regla de BP portada como activo de ci-pattern:
``VENDORED_ROOTS = (".agents/skills/", "skills/")`` queda fuera del
presupuesto de 400 líneas, pero un renombrado desde una ruta de autoría
hacia una vendorizada sí cuenta, porque el trabajo fue authored.

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


class VendoredRootsRule(unittest.TestCase):
    """RED: la regla VENDORED_ROOTS debe excluir el subtree vendorizado."""

    def test_vendored_path_under_agents_skills_is_recognised(self) -> None:
        gate = _load_gate()
        self.assertTrue(gate.is_under_vendored(".agents/skills/foo/auth.py"))
        self.assertTrue(gate.is_under_vendored("skills/legacy/skill/SKILL.md"))
        # Igual con prefijos relativos al repo
        self.assertTrue(gate.is_under_vendored("skills/x.py"))

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
        # El issue #623 fija explícitamente estos dos prefijos
        self.assertEqual(gate.VENDORED_ROOTS, (".agents/skills/", "skills/"))


class CountsToBudgetRule(unittest.TestCase):
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
        self.assertFalse(
            gate.counts_to_budget("skills/legacy/x.py", "skills/legacy/x.py"),
            "skills/<x> sin renombrar NO debe contar",
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
        self.assertFalse(
            gate.counts_to_budget(
                "skills/old/x.py", "skills/new/x.py",
            ),
            "renombrar dentro de skills/ NO debe contar",
        )


# ---------------------------------------------------------------------------
# Tests de integración sobre un repo git real
# ---------------------------------------------------------------------------


class ChangedLinesIntegration(unittest.TestCase):
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


class CliExitCodes(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()


# ---------------------------------------------------------------------------
# RED: el modo local mide contra la base real del PR (issue #658)
# ---------------------------------------------------------------------------


class LocalBaseResolution(unittest.TestCase):
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
