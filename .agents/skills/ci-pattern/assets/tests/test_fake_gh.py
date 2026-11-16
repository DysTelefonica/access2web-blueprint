#!/usr/bin/env python3
# ci-pattern asset tests -- the fake gh is reached by explicit path, never through PATH (#1609)
"""The shared fake ``gh`` and its fail-closed guard.

A test that uses ``_fake_gh`` runs the asset only when the ``gh`` the asset
would run (``CI_PATTERN_GH`` or, unset, ``gh`` on the PATH) lies inside the
test's fixture dir. With a real ``gh`` first on the PATH the guard refuses and
no process is started. The "real gh" here is a decoy that records any call.
"""
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _fake_gh  # noqa: E402
from _fake_gh import GH_ENV, RealGhRefusedError, install_fake_gh, run_guarded  # noqa: E402

RECORD = ("import pathlib, sys\n"
          "pathlib.Path(sys.argv[0]).with_name('called').write_text(' '.join(sys.argv[1:]), encoding='utf-8')\n")
# An asset that runs gh, with the rule the assets apply (#1609).
ASK_GH = ("import os, shutil, subprocess, sys\n"
          "gh = os.environ.get('CI_PATTERN_GH')\n"
          "cmd = [sys.executable, gh] if gh and gh.endswith('.py') else [shutil.which(gh or 'gh')]\n"
          "sys.exit(subprocess.run([*cmd, 'api', 'repos/x/y?per_page=100&page=1']).returncode)\n")


class FakeGhGuardTests(unittest.TestCase):
    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.fixture = Path(tmp.name) / "fixture"
        real_dir = Path(tmp.name) / "real"
        self.fake = install_fake_gh(self.fixture / "bin", RECORD)
        self.real = install_fake_gh(real_dir, RECORD)
        self.env = {k: v for k, v in os.environ.items() if k != GH_ENV}
        self.env["PATH"] = f"{real_dir}{os.pathsep}{self.fixture / 'bin'}{os.pathsep}{os.environ.get('PATH', '')}"

    def run_asset(self, env: dict) -> subprocess.CompletedProcess:
        return run_guarded([sys.executable, "-c", ASK_GH], env=env, fixture_dir=self.fixture,
                           capture_output=True, text=True, timeout=60)

    def assert_refused(self, env: dict) -> None:
        with mock.patch.object(_fake_gh.subprocess, "run") as spawn:
            with self.assertRaises(RealGhRefusedError) as caught:
                self.run_asset(env)
        spawn.assert_not_called()
        self.assertIn("refused: the gh this test would run", str(caught.exception))
        self.assertIn("no call was made (#1609)", str(caught.exception))
        self.assertFalse((self.real.parent / "called").exists())

    def test_real_gh_first_on_path_refuses_without_any_call(self) -> None:
        self.assert_refused(self.env)

    def test_bare_name_in_the_variable_refuses_on_windows(self) -> None:
        env = {**self.env, GH_ENV: "gh",
               "PATH": f"{self.fake.parent}{os.pathsep}{os.environ.get('PATH', '')}"}
        if os.name == "nt":
            self.assert_refused(env)
        else:
            self.run_asset(env)  # POSIX resolves the bare name on PATH: the fake one

    @unittest.skipIf(os.name == "nt", "symlinks need privileges on Windows")
    def test_symlink_in_fixture_to_outside_gh_refuses(self) -> None:
        link = self.fixture / "bin" / "link"
        link.symlink_to(self.real)
        self.assert_refused({**self.env, GH_ENV: str(link)})

    def test_trailing_separator_on_real_gh_dir_refuses(self) -> None:
        real_dir = self.real.parent
        self.assert_refused({**self.env, "PATH": f"{real_dir}{os.sep}{os.pathsep}{self.env['PATH']}"})

    def test_explicit_gh_outside_the_fixture_refuses(self) -> None:
        self.assert_refused({**self.env, GH_ENV: str(self.real)})

    def test_ambient_gh_of_this_host_is_refused(self) -> None:
        env = {k: v for k, v in os.environ.items() if k != GH_ENV}
        self.assert_refused(env)

    def test_fake_gh_by_explicit_path_runs_despite_the_real_gh_on_path(self) -> None:
        proc = self.run_asset({**self.env, GH_ENV: str(self.fake)})
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual((self.fake.parent / "called").read_text(encoding="utf-8"),
                         "api repos/x/y?per_page=100&page=1")
        self.assertFalse((self.real.parent / "called").exists())

    def test_the_fake_is_an_absolute_python_script_with_a_launcher_for_the_host(self) -> None:
        self.assertTrue(self.fake.is_absolute())
        self.assertEqual(self.fake.suffix, ".py")
        self.assertTrue((self.fake.parent / ("gh.cmd" if os.name == "nt" else "gh")).is_file())


if __name__ == "__main__":
    unittest.main()
