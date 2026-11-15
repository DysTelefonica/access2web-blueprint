#!/usr/bin/env python3
# ci-pattern asset tests — vendored tree and seal checked against the base (#1261)
"""A PR job without secrets must not let the PR change the gate it runs.

`vendored_seal_check.py` compares the PR head with its base commit: the
vendored pattern tree and the `catalog @` seal of `.team-skills.yaml` must be
identical to what the base declares. The check runs from the base copy of the
script, so a PR cannot edit the checker either.

Runs offline in temporary git repositories; no network, no secret.
"""
import subprocess
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
CHECK = ASSET / "vendored_seal_check.py"
VENDORED = ".agents/skills/ci-pattern"
MANIFEST = ".team-skills.yaml"
SEAL = "# Auto-generado por propagate-team-skills.ps1 (catalog @ 1d34d1c76ab3c0e8f0a1b2c3d4e5f60718293a4b).\n"
GIT = ["git", "-c", "user.name=t", "-c", "user.email=t@example.invalid",
       "-c", "commit.gpgsign=false"]


def git(repo: Path, *args: str) -> str:
    return subprocess.run(GIT + list(args), cwd=repo, check=True,
                          capture_output=True, text=True).stdout.strip()


def write(repo: Path, rel: str, text: str) -> None:
    path = repo / rel
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def base_repo(tmp: Path, seal: str = SEAL) -> Path:
    """A consumer at its base commit: manifest, vendored gate and one product file."""
    repo = tmp / "consumer"
    repo.mkdir()
    git(repo, "init", "-q", "-b", "main")
    write(repo, MANIFEST, seal)
    write(repo, f"{VENDORED}/assets/bin/gate.py", "print('gate')\n")
    write(repo, "src/app.py", "x = 1\n")
    git(repo, "add", "-A")
    git(repo, "commit", "-q", "-m", "base")
    return repo


def run_check(repo: Path, base: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        ["python3", "-I", str(CHECK), "--base", base, "--vendored", VENDORED,
         "--manifest", MANIFEST],
        cwd=repo, capture_output=True, text=True)


class VendoredSealCheckTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self):
        self._tmp.cleanup()

    def test_pr_identical_to_base_passes(self):
        repo = base_repo(self.tmp)
        base = git(repo, "rev-parse", "HEAD")
        proc = run_check(repo, base)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_pr_that_changes_only_product_code_passes(self):
        """The PR can change its own code freely; only the gate tree is pinned."""
        repo = base_repo(self.tmp)
        base = git(repo, "rev-parse", "HEAD")
        write(repo, "src/app.py", "x = 2\n")
        git(repo, "commit", "-qam", "pr")
        proc = run_check(repo, base)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_altered_vendored_file_fails_against_base(self):
        """Un vendorizado alterado falla contra el sello de base (#1261)."""
        repo = base_repo(self.tmp)
        base = git(repo, "rev-parse", "HEAD")
        write(repo, f"{VENDORED}/assets/bin/gate.py", "print('always green')\n")
        git(repo, "commit", "-qam", "pr: weaken the gate")
        proc = run_check(repo, base)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn(VENDORED, proc.stdout + proc.stderr)

    def test_vendored_file_added_by_pr_fails(self):
        repo = base_repo(self.tmp)
        base = git(repo, "rev-parse", "HEAD")
        write(repo, f"{VENDORED}/assets/bin/extra.py", "print('injected')\n")
        git(repo, "add", "-A")
        git(repo, "commit", "-q", "-m", "pr: add a gate file")
        proc = run_check(repo, base)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)

    def test_untracked_vendored_file_fails(self):
        """A file the PR leaves uncommitted in the vendored tree still counts."""
        repo = base_repo(self.tmp)
        base = git(repo, "rev-parse", "HEAD")
        write(repo, f"{VENDORED}/assets/bin/extra.py", "print('untracked')\n")
        proc = run_check(repo, base)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)

    def test_seal_repointed_by_pr_fails(self):
        """A PR cannot point the gate at another catalog commit."""
        repo = base_repo(self.tmp)
        base = git(repo, "rev-parse", "HEAD")
        write(repo, MANIFEST, SEAL.replace("1d34d1c76ab3", "ffffffffffff"))
        git(repo, "commit", "-qam", "pr: re-seal")
        proc = run_check(repo, base)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn(MANIFEST, proc.stdout + proc.stderr)

    def test_base_without_seal_fails_closed(self):
        """HR-3: sin sello en la base no hay commit del patrón que medir."""
        repo = base_repo(self.tmp, seal="# sin sello\n")
        base = git(repo, "rev-parse", "HEAD")
        proc = run_check(repo, base)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)

    def test_base_seal_with_non_hex_sha_fails_closed(self):
        repo = base_repo(self.tmp, seal="# (catalog @ xyzxyzxyz).\n")
        base = git(repo, "rev-parse", "HEAD")
        proc = run_check(repo, base)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)

    def test_unknown_base_commit_fails_closed(self):
        repo = base_repo(self.tmp)
        proc = run_check(repo, "0123456789abcdef0123456789abcdef01234567")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)

    def test_abbreviated_seal_matching_base_passes(self):
        """#1157: the seal may be short; the check compares the line, not a resolved sha."""
        repo = base_repo(self.tmp, seal="# Auto (catalog @ 1d34d1c).\n")
        base = git(repo, "rev-parse", "HEAD")
        proc = run_check(repo, base)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)


if __name__ == "__main__":
    unittest.main()
