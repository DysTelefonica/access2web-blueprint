#!/usr/bin/env python3
# ci-pattern asset tests — release image build from a fresh clone of the tag (#1389)
"""Black-box suite of ``build_release_image.py``.

The container engine is a stand-in that records its argv, so the suite needs
no podman or docker. Rules under test: the build clones the tag fresh,
refuses a missing tag, a tag off the default branch and a checkout whose
deployable tree is not the git content of the tag, and labels the image with
the tag commit and the tag.

Exit codes: 0 built · 1 refused · 2 tool failure."""
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _fixtures import ASSET, GIT_ENV, Repo, git, write_config  # noqa: E402

BUILD = ASSET / "build_release_image.py"
RECORDER = ("import json, sys; "
            "open(sys.argv[1], 'a', encoding='utf-8').write(json.dumps(sys.argv[2:]) + '\\n')")


class BuildTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.repo = Repo(self.tmp / "repo")
        self.repo.write("www/index.php", b"<?php echo 1;\n")
        self.repo.write("Dockerfile", b"FROM scratch\nCOPY www /var/www/html\n")
        self.commit = self.repo.commit()
        self.repo.tag("v1.0.0")
        self.calls = self.tmp / "engine.jsonl"
        self.cfg = write_config(self.tmp / "cfg.json",
                                engine_cmd=[sys.executable, "-c", RECORDER, str(self.calls)])

    def tearDown(self):
        self._tmp.cleanup()

    def build(self, tag="v1.0.0", *extra):
        return subprocess.run([sys.executable, str(BUILD), "--config", str(self.cfg),
                               "--tag", tag, "--git-remote", str(self.repo.d), *extra],
                              capture_output=True, text=True, encoding="utf-8", env=GIT_ENV)

    def engine_calls(self):
        if not self.calls.exists():
            return []
        return [json.loads(x) for x in self.calls.read_text(encoding="utf-8").splitlines()]

    def test_builds_the_tag_with_its_commit_and_version_labels(self):
        r = self.build()
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertIn(f"RELEASE-BUILD OK: registry.example.test/team/app:v1.0.0 built from "
                      f"tag 'v1.0.0' ({self.commit[:12]})", r.stdout)
        (call,) = self.engine_calls()
        self.assertEqual(call[0], "build")
        self.assertIn(f"org.opencontainers.image.revision={self.commit}", call)
        self.assertIn("org.opencontainers.image.version=v1.0.0", call)
        self.assertEqual(call[call.index("-t") + 1], "registry.example.test/team/app:v1.0.0")
        self.assertTrue(call[call.index("-f") + 1].endswith("Dockerfile"))
        self.assertNotEqual(Path(call[-1]).resolve(), self.repo.d.resolve())

    def test_push_runs_after_the_build(self):
        r = self.build("v1.0.0", "--push")
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertEqual([c[0] for c in self.engine_calls()], ["build", "push"])
        self.assertIn("and pushed", r.stdout)

    def test_missing_tag_is_refused_without_building(self):
        r = self.build("v2.0.0")
        self.assertEqual(r.returncode, 1)
        self.assertIn("tag 'v2.0.0' does not exist on the remote", r.stdout)
        self.assertEqual(self.engine_calls(), [])

    def test_tag_off_the_default_branch_is_refused(self):
        git(self.repo.d, "switch", "-q", "-c", "staging")
        self.repo.write("www/index.php", b"staging\n")
        side = self.repo.commit("side")
        self.repo.tag("v1.1.0")
        r = self.build("v1.1.0")
        self.assertEqual(r.returncode, 1)
        self.assertIn(f"tag 'v1.1.0' ({side[:12]}) is not on 'main'", r.stdout)
        self.assertEqual(self.engine_calls(), [])

    def test_checkout_that_differs_from_the_git_content_is_refused(self):
        self.repo.write(".gitattributes", b"*.php text eol=crlf\n")
        self.repo.write("www/index.php", b"<?php echo 3;\n")
        self.repo.commit("crlf on checkout")
        self.repo.tag("v1.2.0")
        r = self.build("v1.2.0")
        self.assertEqual(r.returncode, 1, r.stdout + r.stderr)
        self.assertIn("checkout differs from the git content of the tag (changed: index.php)",
                      r.stdout)
        self.assertEqual(self.engine_calls(), [])

    def test_engine_failure_is_a_tool_failure(self):
        write_config(self.cfg, engine_cmd=[sys.executable, "-c", "import sys; sys.exit(3)"])
        r = self.build()
        self.assertEqual(r.returncode, 2)
        self.assertIn("RELEASE-BUILD ERROR: building registry.example.test/team/app:v1.0.0",
                      r.stderr)


if __name__ == "__main__":
    unittest.main()
