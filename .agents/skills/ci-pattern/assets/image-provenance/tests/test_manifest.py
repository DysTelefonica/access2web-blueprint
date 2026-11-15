#!/usr/bin/env python3
# ci-pattern asset tests — image-provenance release manifest and comparer (#1389)
"""Black-box suite of ``image_provenance.py manifest`` and of the comparer.

Manifest: sha256 of the deployable tree from the git objects of the tag commit,
no line-ending normalisation. Exit: 0 written, 2 unmeasurable."""
import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import image_provenance as ip  # noqa: E402
from _fixtures import Repo, run_tool, sha, write_config  # noqa: E402


class ManifestTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.repo = Repo(self.tmp / "repo")
        self.cfg = write_config(self.tmp / "cfg.json", exclude=["vendor/"])
        self.out = self.tmp / "manifest.json"

    def tearDown(self):
        self._tmp.cleanup()

    def manifest(self, tag="v1.0.0"):
        r = run_tool("manifest", "--config", self.cfg, "--repo", self.repo.d,
                     "--tag", tag, "--out", self.out)
        return r, (json.loads(self.out.read_text(encoding="utf-8")) if r.returncode == 0 else None)

    def test_hashes_every_file_of_the_tree_at_the_tag_commit(self):
        self.repo.write("www/index.php", b"<?php echo 1;\n")
        self.repo.write("www/css/a.css", b"body{}\n")
        self.repo.write("Dockerfile", b"FROM scratch\n")
        commit = self.repo.commit()
        self.repo.tag("v1.0.0")
        self.repo.write("www/index.php", b"<?php echo 2;\n")
        self.repo.commit("after the tag")
        r, man = self.manifest()
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn("MANIFEST OK: 2 file(s)", r.stdout)
        self.assertEqual(man["commit"], commit)
        self.assertEqual(man["tag"], "v1.0.0")
        self.assertEqual(man["schema"], ip.MANIFEST_SCHEMA)
        self.assertEqual(man["files"], {"index.php": sha(b"<?php echo 1;\n"),
                                        "css/a.css": sha(b"body{}\n")})

    def test_excluded_directories_stay_out(self):
        self.repo.write("www/index.php", b"x\n")
        self.repo.write("www/vendor/lib.php", b"v\n")
        self.repo.commit()
        self.repo.tag("v1.0.0")
        r, man = self.manifest()
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertEqual(sorted(man["files"]), ["index.php"])
        self.assertEqual(man["exclude"], ["vendor/"])

    def test_blob_bytes_are_hashed_without_normalisation(self):
        self.repo.write("www/lf.php", b"a\nb\n")
        self.repo.write("www/crlf.bat", b"a\r\nb\r\n")
        self.repo.commit()
        self.repo.tag("v1.0.0")
        _, man = self.manifest()
        self.assertEqual(man["files"]["lf.php"], sha(b"a\nb\n"))
        self.assertEqual(man["files"]["crlf.bat"], sha(b"a\r\nb\r\n"))
        self.assertTrue(man["normalisation"].startswith("none"))

    def test_symlinks_are_recorded_by_target(self):
        self.repo.write("www/index.php", b"x\n")
        self.repo.symlink("www/latest", "index.php")
        self.repo.commit()
        self.repo.tag("v1.0.0")
        _, man = self.manifest()
        self.assertEqual(man["files"]["latest"], "symlink:index.php")

    def test_missing_tag_is_unmeasurable(self):
        self.repo.write("www/index.php", b"x\n")
        self.repo.commit()
        r, _ = self.manifest("v9.9.9")
        self.assertEqual(r.returncode, 2)
        self.assertIn("tag 'v9.9.9' does not exist", r.stderr)
        self.assertFalse(self.out.exists())

    def test_empty_tree_is_unmeasurable(self):
        self.repo.write("src/app.py", b"x\n")
        self.repo.commit()
        self.repo.tag("v1.0.0")
        r, _ = self.manifest()
        self.assertEqual(r.returncode, 2)
        self.assertIn("has no files", r.stderr)

    def test_config_without_required_keys_is_invalid(self):
        self.cfg.write_text(json.dumps({"tree": "www"}), encoding="utf-8")
        r, _ = self.manifest()
        self.assertEqual(r.returncode, 2)
        self.assertIn("image_path, image, github_repo", r.stderr)


class CompareTests(unittest.TestCase):
    EXPECTED = {"a.php": sha(b"a\n"), "b.php": sha(b"b\n")}

    def test_match(self):
        self.assertEqual(ip.compare(self.EXPECTED, dict(self.EXPECTED), {}), [])

    def test_extra_missing_and_changed_files_are_named(self):
        actual = {"a.php": sha(b"A\n"), "c.php": sha(b"c\n")}
        self.assertEqual(ip.compare(self.EXPECTED, actual, {}),
                         ["changed: a.php", "missing in image: b.php",
                          "extra in image: c.php"])

    def test_line_ending_only_change_is_a_finding_with_its_cause(self):
        actual = {"a.php": sha(b"a\r\n"), "b.php": sha(b"b\n")}
        lf = {"a.php": ip.lf_sha256(b"a\r\n")}
        self.assertEqual(ip.compare(self.EXPECTED, actual, lf),
                         ["changed in line endings only (CRLF in image): a.php"])

    def test_exclusion_patterns(self):
        pats = ["vendor/", "cache/**", "*.log"]
        self.assertTrue(ip.is_excluded("vendor", pats))
        self.assertTrue(ip.is_excluded("vendor/x/y.php", pats))
        self.assertTrue(ip.is_excluded("cache/a", pats))
        self.assertTrue(ip.is_excluded("logs/app.log", pats))
        self.assertFalse(ip.is_excluded("vendored.php", pats))


if __name__ == "__main__":
    unittest.main()
