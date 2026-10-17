#!/usr/bin/env python3
# ci-pattern asset tests — safe-cleanup core (DysTelefonica/team-skills#436, chain tramo a)
"""Suite del nucleo de safe_cleanup.py: criterios de preservacion
(3 tests), motivos de rechazo (3 tests), fail-closed de entorno
(1 test) y contrato de no-mutacion remota (1 test). El camino del
bundle y el CLI llegan en los tramos (b) y (c) de la cadena.

Solo stdlib; repos temporales con git real; sin red."""
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET_DIR = Path(__file__).resolve().parent.parent
ASSET = ASSET_DIR / "safe_cleanup.py"

GIT_ENV = {
    "GIT_AUTHOR_NAME": "tester",
    "GIT_AUTHOR_EMAIL": "tester@local",
    "GIT_COMMITTER_NAME": "tester",
    "GIT_COMMITTER_EMAIL": "tester@local",
}


def _run(args: list[str], cwd: Path, env_extra: dict | None = None) -> subprocess.CompletedProcess:
    env = {**os.environ, **GIT_ENV, **(env_extra or {})}
    return subprocess.run(args, capture_output=True, text=True, cwd=cwd, env=env, timeout=20)


def make_repo(tmp: Path, default_branch: str = "main") -> Path:
    """Crea un repo con origin remoto (bare) y rama main con un commit."""
    repo = tmp / "repo"
    remote = tmp / "remote.git"
    repo.mkdir()
    remote.mkdir()
    _run(["git", "init", "-q", "-b", default_branch], cwd=repo)
    _run(["git", "init", "-q", "--bare"], cwd=remote)
    _run(["git", "remote", "add", "origin", str(remote)], cwd=repo)
    _run(["git", "config", "user.email", "tester@local"], cwd=repo)
    _run(["git", "config", "user.name", "tester"], cwd=repo)
    (repo / "README.md").write_text("init\n", encoding="utf-8")
    _run(["git", "add", "README.md"], cwd=repo)
    _run(["git", "commit", "-q", "-m", "init"], cwd=repo)
    _run(["git", "push", "-q", "-u", "origin", default_branch], cwd=repo)
    _run(["git", "remote", "set-head", "origin", "--auto"], cwd=repo)
    return repo


def make_branch(repo: Path, name: str, default_branch: str = "main") -> str:
    """Crea una rama con un commit y la sube. Devuelve el SHA del tip."""
    _run(["git", "checkout", "-q", default_branch], cwd=repo)
    _run(["git", "checkout", "-q", "-b", name], cwd=repo)
    (repo / f"{name}.txt").write_text(f"branch {name}\n", encoding="utf-8")
    _run(["git", "add", f"{name}.txt"], cwd=repo)
    _run(["git", "commit", "-q", "-m", f"work on {name}"], cwd=repo)
    _run(["git", "push", "-q", "-u", "origin", name], cwd=repo)
    tip = _run(["git", "rev-parse", "HEAD"], cwd=repo).stdout.strip()
    return tip


def setup_worktree(repo: Path, branch: str, wt_path: Path) -> None:
    """Crea un worktree real para la rama en wt_path. Sale del checkout
    de esa rama en el main repo primero; git rechaza ``worktree add``
    si la rama ya esta usada por otro worktree (incluido el principal)."""
    _run(["git", "checkout", "-q", "main"], cwd=repo)
    _run(["git", "worktree", "add", "-q", str(wt_path), branch], cwd=repo)


def _import():
    sys.path.insert(0, str(ASSET_DIR))
    for m in list(sys.modules):
        if m == "safe_cleanup":
            del sys.modules[m]
    import safe_cleanup  # noqa: E402
    return safe_cleanup


class PreservationTests(unittest.TestCase):
    """Cubre las tres ramas del OR de preservacion."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def tearDown(self) -> None:
        for d in self.tmp.glob("*"):
            if d.is_dir() and (d / ".git").exists():
                _run(["git", "worktree", "prune"], cwd=d)

    def test_tip_matches_remote_ref_is_preserved(self):
        """Caso 1: la punta local coincide con origin/<branch>."""
        repo = make_repo(self.tmp)
        tip = make_branch(repo, "feat-tip")
        sc = _import()
        result = sc.cleanup(str(repo), "feat-tip")
        self.assertTrue(result.ok, result)
        self.assertEqual(result.exit_code, 0)
        self.assertFalse((repo / ".git" / "refs" / "heads" / "feat-tip").exists())
        remote_tip = _run(
            ["git", "rev-parse", "origin/feat-tip"], cwd=repo
        ).stdout.strip()
        self.assertEqual(remote_tip, tip)

    def test_tip_is_ancestor_of_active_branch_is_preserved(self):
        """Caso 2: la rama fue mergeada a main, asi que su punta es
        ancestro de origin/main."""
        repo = make_repo(self.tmp)
        make_branch(repo, "feat-anc")
        _run(["git", "checkout", "-q", "main"], cwd=repo)
        merge = _run(["git", "merge", "--ff-only", "feat-anc"], cwd=repo)
        self.assertEqual(merge.returncode, 0, merge.stderr)
        _run(["git", "push", "-q", "origin", "main"], cwd=repo)
        _run(["git", "push", "origin", ":feat-anc"], cwd=repo)
        sc = _import()
        result = sc.cleanup(str(repo), "feat-anc")
        self.assertTrue(result.ok, result.reason)

    def test_tip_equals_pr_head_is_preserved(self):
        """Caso 3: la punta coincide con pr_head."""
        repo = make_repo(self.tmp)
        tip = make_branch(repo, "feat-pr")
        _run(["git", "push", "origin", ":feat-pr"], cwd=repo)
        sc = _import()
        result = sc.cleanup(str(repo), "feat-pr", pr_head=tip)
        self.assertTrue(result.ok, result.reason)


class RefusalTests(unittest.TestCase):
    """Cubre cada motivo de rechazo (exit 3)."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def tearDown(self) -> None:
        for d in self.tmp.glob("*"):
            if d.is_dir() and (d / ".git").exists():
                _run(["git", "worktree", "prune"], cwd=d)

    def test_branch_not_preserved_refuses(self):
        """Sin remote, sin pr_head, sin ancestro: rc=3, nada se borra."""
        repo = make_repo(self.tmp)
        make_branch(repo, "feat-orphan")
        _run(["git", "push", "origin", ":feat-orphan"], cwd=repo)
        _run(["git", "checkout", "-q", "main"], cwd=repo)
        (repo / "main-advance.txt").write_text("adv\n", encoding="utf-8")
        _run(["git", "add", "main-advance.txt"], cwd=repo)
        _run(["git", "commit", "-q", "-m", "advance main"], cwd=repo)
        _run(["git", "push", "-q", "origin", "main"], cwd=repo)
        sc = _import()
        result = sc.cleanup(str(repo), "feat-orphan")
        self.assertFalse(result.ok)
        self.assertEqual(result.exit_code, 3)
        self.assertTrue((repo / ".git" / "refs" / "heads" / "feat-orphan").exists())
        self.assertIn("not preserved", result.reason.lower())

    def test_dirty_worktree_refuses(self):
        """Worktree con cambios sin commit: rc=3, nada se borra."""
        repo = make_repo(self.tmp)
        make_branch(repo, "feat-dirty")
        wt = self.tmp / "wt"
        setup_worktree(repo, "feat-dirty", wt)
        (wt / "main-advance.txt").write_text("dirty\n", encoding="utf-8")
        sc = _import()
        result = sc.cleanup(str(repo), "feat-dirty", worktree=str(wt))
        self.assertFalse(result.ok)
        self.assertEqual(result.exit_code, 3)
        self.assertIn("dirty", result.reason.lower())
        self.assertTrue(wt.exists())

    def test_process_in_worktree_refuses(self):
        """Si un proceso tiene cwd dentro del worktree: rc=3."""
        repo = make_repo(self.tmp)
        make_branch(repo, "feat-lock")
        wt = self.tmp / "wt"
        setup_worktree(repo, "feat-lock", wt)
        proc = subprocess.Popen(
            [sys.executable, "-c", "import time; time.sleep(30)"],
            cwd=str(wt),
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        try:
            sc = _import()
            result = sc.cleanup(str(repo), "feat-lock", worktree=str(wt))
            self.assertFalse(result.ok)
            self.assertEqual(result.exit_code, 3)
            self.assertIn("process", result.reason.lower())
        finally:
            proc.terminate()
            try:
                proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                proc.kill()


class EnvironmentFailTests(unittest.TestCase):
    """Errores de entorno: rc=2, fail-closed."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def test_nonexistent_repo_fails_closed(self):
        sc = _import()
        result = sc.cleanup(str(self.tmp / "no-such-repo"), "feat-x")
        self.assertFalse(result.ok)
        self.assertEqual(result.exit_code, 2)


class NoRemoteMutationTests(unittest.TestCase):
    """Garantia dura: nunca se invoca `git push origin --delete` ni similar."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def test_remote_branch_unchanged_after_cleanup(self):
        repo = make_repo(self.tmp)
        tip = make_branch(repo, "feat-keep-remote")
        remote_tip_before = _run(
            ["git", "rev-parse", "origin/feat-keep-remote"], cwd=repo
        ).stdout.strip()
        sc = _import()
        result = sc.cleanup(str(repo), "feat-keep-remote")
        self.assertTrue(result.ok, result.reason)
        remote_tip_after = _run(
            ["git", "rev-parse", "origin/feat-keep-remote"], cwd=repo
        ).stdout.strip()
        self.assertEqual(remote_tip_before, tip)
        self.assertEqual(remote_tip_after, tip)


class BundleTests(unittest.TestCase):
    """Cubre los 2 tests del bundle (tramo 2/3): creacion exitosa y
    fallo de creacion. Tambien la semantica de bundled_to en rechazos."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def tearDown(self) -> None:
        for d in self.tmp.glob("*"):
            if d.is_dir() and (d / ".git").exists():
                _run(["git", "worktree", "prune"], cwd=d)

    def test_bundle_created_on_successful_cleanup(self):
        """Cuando cleanup succeeds y bundle_dir se pasa, bundled_to
        apunta al bundle creado."""
        repo = make_repo(self.tmp)
        make_branch(repo, "feat-bundle-ok")
        bd = self.tmp / "backups"
        sc = _import()
        result = sc.cleanup(str(repo), "feat-bundle-ok", bundle_dir=str(bd))
        self.assertTrue(result.ok, result)
        self.assertIsNotNone(result.bundled_to)
        self.assertTrue(result.bundled_to.exists())
        self.assertIn("feat-bundle-ok", str(result.bundled_to))

    def test_bundle_created_on_rejected_cleanup(self):
        """Cuando cleanup se rechaza pero bundle_dir se pasa, bundled_to
        se llena igual: el bundle es un respaldo que el operador puede
        reutilizar aunque no se borre."""
        repo = make_repo(self.tmp)
        make_branch(repo, "feat-bundle-reject")
        # Hacer que no sea preservada: borrar remote y avanzar main.
        _run(["git", "push", "origin", ":feat-bundle-reject"], cwd=repo)
        _run(["git", "checkout", "-q", "main"], cwd=repo)
        (repo / "main-advance.txt").write_text("adv\n", encoding="utf-8")
        _run(["git", "add", "main-advance.txt"], cwd=repo)
        _run(["git", "commit", "-q", "-m", "advance main"], cwd=repo)
        _run(["git", "push", "-q", "origin", "main"], cwd=repo)
        bd = self.tmp / "backups"
        sc = _import()
        result = sc.cleanup(str(repo), "feat-bundle-reject", bundle_dir=str(bd))
        self.assertFalse(result.ok)
        self.assertEqual(result.exit_code, 3)
        self.assertIsNotNone(result.bundled_to)
        self.assertTrue(result.bundled_to.exists())


class CLITests(unittest.TestCase):
    """Cubre los 2 tests del CLI (tramo 2/3): exito y rechazo por
    stdout, y salida de BUNDLE en ambos caminos."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def tearDown(self) -> None:
        for d in self.tmp.glob("*"):
            if d.is_dir() and (d / ".git").exists():
                _run(["git", "worktree", "prune"], cwd=d)

    def test_cli_success_returns_0_and_prints_clean(self):
        """El CLI con exito retorna 0, imprime CLEAN y BUNDLE."""
        repo = make_repo(self.tmp)
        make_branch(repo, "feat-cli-ok")
        bd = self.tmp / "backups"
        proc = _run(
            [sys.executable, str(ASSET),
             "--repo-dir", str(repo),
             "--branch", "feat-cli-ok",
             "--bundle-dir", str(bd)],
            cwd=self.tmp,
        )
        self.assertEqual(proc.returncode, 0, proc.stderr)
        out = proc.stdout
        self.assertIn("CLEAN: branch:feat-cli-ok", out)
        self.assertTrue(any("OK:" in line for line in out.strip().splitlines()))
        self.assertIn("BUNDLE:", out)

    def test_cli_refusal_returns_3_and_prints_refused(self):
        """El CLI con rechazo retorna 3, imprime REFUSED y BUNDLE."""
        repo = make_repo(self.tmp)
        make_branch(repo, "feat-cli-ref")
        _run(["git", "push", "origin", ":feat-cli-ref"], cwd=repo)
        _run(["git", "checkout", "-q", "main"], cwd=repo)
        (repo / "main-advance.txt").write_text("adv\n", encoding="utf-8")
        _run(["git", "add", "main-advance.txt"], cwd=repo)
        _run(["git", "commit", "-q", "-m", "advance main"], cwd=repo)
        _run(["git", "push", "-q", "origin", "main"], cwd=repo)
        bd = self.tmp / "backups"
        proc = _run(
            [sys.executable, str(ASSET),
             "--repo-dir", str(repo),
             "--branch", "feat-cli-ref",
             "--bundle-dir", str(bd)],
            cwd=self.tmp,
        )
        self.assertEqual(proc.returncode, 3, proc.stdout)
        self.assertTrue(any("REFUSED:" in line for line in proc.stdout.strip().splitlines()))
        self.assertIn("BUNDLE:", proc.stdout)


class IntegrationWorktreeTests(unittest.TestCase):
    """Prueba de integracion: safe_cleanup con un worktree real
    conectado a una rama que fue mergeada a main. Verifica que
    safe_cleanup borra tanto el worktree como la rama local, y
    que el bundle se guarda."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def tearDown(self) -> None:
        for d in self.tmp.glob("*"):
            if d.is_dir() and (d / ".git").exists():
                _run(["git", "worktree", "prune"], cwd=d)

    def test_cleanup_worktree_and_branch_after_merge(self):
        """La rama fue mergeada a main, el worktree se limpia junto
        con la rama local. El bundle se guarda."""
        repo = make_repo(self.tmp)
        tip = make_branch(repo, "feat-integ")
        # Merge a main
        _run(["git", "checkout", "-q", "main"], cwd=repo)
        _run(["git", "merge", "--ff-only", "feat-integ"], cwd=repo)
        _run(["git", "push", "-q", "origin", "main"], cwd=repo)
        # Crea un worktree para la rama (ya mergeada)
        wt = self.tmp / "wt"
        setup_worktree(repo, "feat-integ", wt)
        # Borra la rama remota (ya esta en main)
        _run(["git", "push", "origin", ":feat-integ"], cwd=repo)
        # Vuelve a main en el repo principal
        _run(["git", "checkout", "-q", "main"], cwd=repo)
        bd = self.tmp / "backups"
        sc = _import()
        result = sc.cleanup(
            str(repo), "feat-integ",
            worktree=str(wt), bundle_dir=str(bd),
        )
        self.assertTrue(result.ok, result)
        self.assertEqual(result.exit_code, 0)
        self.assertIsNotNone(result.bundled_to)
        self.assertTrue(result.bundled_to.exists())
        self.assertIn("worktree", result.deleted[0])
        self.assertIn("branch", result.deleted[1])
        self.assertFalse(wt.exists(), "worktree debe haber sido borrado")
        remote_tip = _run(
            ["git", "rev-parse", "origin/main"], cwd=repo
        ).stdout.strip()
        self.assertEqual(remote_tip, tip, "main debe retener la punta de feat-integ")

if __name__ == "__main__":
    unittest.main()
