#!/usr/bin/env python3
# ci-pattern asset tests -- safe merge, limpieza tras el merge (DysTelefonica/team-skills#1491)
"""Tras el merge, safe-merge conserva la rama remota y deja que
``safe_cleanup`` decida si la rama local se puede borrar: solo se borra si su
punta esta preservada (la cabeza de la PR fusionada, ``origin/<rama>`` o un
ancestro de la rama activa). La salida de ``gh`` se lee como UTF-8 con
cualquier locale.
"""
import importlib.util
import json
import subprocess
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _shim import ASSET, BaseSafeMergeTest, base_state, run_gate

BRANCH = "chore/test"


def git(repo: Path, *args: str) -> str:
    return subprocess.run(["git", "-C", str(repo), *args], check=True,
                          capture_output=True, text=True).stdout.strip()


def local_repo(root: Path) -> tuple[Path, str]:
    """Repo local con la rama de la PR y un commit que no esta en ningun remoto."""
    repo = root / "work"
    repo.mkdir()
    git(repo, "init", "-q", "-b", "main")
    git(repo, "-c", "user.email=t@x", "-c", "user.name=t", "commit", "-q", "--allow-empty", "-m", "base")
    git(repo, "checkout", "-q", "-b", BRANCH)
    git(repo, "-c", "user.email=t@x", "-c", "user.name=t", "commit", "-q", "--allow-empty", "-m", "work")
    tip = git(repo, "rev-parse", "HEAD")
    git(repo, "checkout", "-q", "main")
    return repo, tip


def headed_at(sha: str) -> dict:
    """El estado del stub con la cabeza de la PR en ``sha``."""
    state = base_state()
    old = state["pr_details"]["7"]["head"]["sha"]
    state["pr_details"]["7"]["head"]["sha"] = sha
    state["check_runs"] = {sha: state["check_runs"][old]}
    return state


class CleanupAfterMergeTest(BaseSafeMergeTest):
    def stored(self) -> dict:
        return json.loads((self.root / "state.json").read_text(encoding="utf-8"))

    def test_local_branch_whose_tip_is_not_preserved_is_kept(self) -> None:
        repo, _tip = local_repo(self.root)
        proc = run_gate(self.root, base_state(), cwd=repo)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(self.stored()["deleted_branches"], [])
        self.assertEqual(git(repo, "branch", "--list", BRANCH).strip("* "), BRANCH)
        self.assertIn(f"KEEP: local branch '{BRANCH}' is kept: branch '{BRANCH}' is not preserved",
                      proc.stdout)

    def test_local_branch_at_the_merged_head_is_deleted(self) -> None:
        repo, tip = local_repo(self.root)
        proc = run_gate(self.root, headed_at(tip), cwd=repo)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(self.stored()["deleted_branches"], [])
        self.assertEqual(git(repo, "branch", "--list", BRANCH), "")
        self.assertIn(f"CLEAN: local branch '{BRANCH}' cleaned (branch tip matches pr_head)", proc.stdout)


class OutputEncodingTest(unittest.TestCase):
    def test_gh_output_is_read_as_utf8_whatever_the_locale(self) -> None:
        # El UTF-8 de este emoji lleva 0x81 y 0x8D, que cp1252 no define (#1491).
        spec = importlib.util.spec_from_file_location("safe_merge", ASSET)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        proc = module.run_command(
            [Path(sys.executable).name, "-c",
             "import sys; sys.stdout.buffer.write('{\\\"t\\\": \\\"🍁\\\"}'.encode('utf-8'))"], 30)
        self.assertEqual(json.loads(proc.stdout), {"t": "🍁"})


if __name__ == "__main__":
    unittest.main()
