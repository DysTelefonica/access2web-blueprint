#!/usr/bin/env python3
# ci-pattern asset tests — image-provenance deploy gate (#1389)
"""Black-box suite of ``image_provenance.py gate``.

Each test builds a local git repository (the "remote"), the release manifest
of a tag with the tool itself, and an image tarball (docker-archive or OCI)
whose layers carry the deployable tree. Rules under test:

- the image revision label is the commit of its release tag, that commit was
  pushed and the tag is on the default branch;
- the deployable tree of the image is the manifest, byte for byte (a
  line-ending-only change is a finding, named as such);
- the diff against the running image is printed and does not change the verdict.

Exit codes: 0 clean · 1 findings (files named) · 2 unmeasurable (fail closed)."""
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _fixtures import (  # noqa: E402
    Repo,
    docker_archive,
    git,
    layer,
    run_tool,
    write_config,
)

ROOT = "var/www/html/"
FILES = {"index.php": b"<?php echo 1;\n", "css/a.css": b"body{}\n"}


def tree(files: dict, root=ROOT) -> dict:
    return {root + k: v for k, v in files.items()}


class GateTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.repo = Repo(self.tmp / "repo")
        for rel, data in FILES.items():
            self.repo.write("www/" + rel, data)
        self.commit = self.repo.commit()
        self.repo.tag("v1.0.0")
        self.cfg = write_config(self.tmp / "cfg.json", exclude=["vendor/"])
        self.manifest = self.tmp / "manifest.json"
        r = run_tool("manifest", "--config", self.cfg, "--repo", self.repo.d,
                     "--tag", "v1.0.0", "--out", self.manifest)
        self.assertEqual(r.returncode, 0, r.stderr)

    def tearDown(self):
        self._tmp.cleanup()

    def labels(self, revision=None, version="v1.0.0"):
        out = {"org.opencontainers.image.revision": revision or self.commit}
        if version:
            out["org.opencontainers.image.version"] = version
        return out

    def image(self, layers, labels=None, name="img.tar", fmt=docker_archive):
        return fmt(self.tmp / name, [layer(x) for x in layers],
                   self.labels() if labels is None else labels)

    def gate(self, archive, *extra, manifest=True):
        args = ["gate", "--config", self.cfg, "--image-archive", archive,
                "--git-remote", self.repo.d]
        if manifest:
            args += ["--manifest", self.manifest]
        return run_tool(*args, *extra)

    # --- accepted ---

    def test_release_image_matching_its_manifest_passes(self):
        r = self.gate(self.image([tree(FILES)]))
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertIn("IMAGE-PROVENANCE OK", r.stdout)
        self.assertIn("2 file(s) match", r.stdout)

    # --- content findings ---

    def test_extra_missing_and_changed_files_are_named(self):
        files = {"index.php": b"<?php echo 2;\n", "new.php": b"n\n"}
        r = self.gate(self.image([tree(files)]))
        self.assertEqual(r.returncode, 1, r.stdout + r.stderr)
        self.assertIn("IMAGE-PROVENANCE FAIL: 3 finding(s)", r.stdout)
        self.assertIn("summary: 1 missing in image, 1 changed, 1 extra in image", r.stdout)
        self.assertIn("deployment refused", r.stdout)
        for line in ("content: changed: index.php", "content: missing in image: css/a.css",
                     "content: extra in image: new.php"):
            self.assertIn(line, r.stdout)

    def test_crlf_copy_of_the_release_is_refused_and_explained(self):
        files = {k: v.replace(b"\n", b"\r\n") for k, v in FILES.items()}
        r = self.gate(self.image([tree(files)]))
        self.assertEqual(r.returncode, 1)
        self.assertIn("changed in line endings only (CRLF in image): index.php", r.stdout)

    # --- provenance findings ---

    def test_revision_never_pushed_is_refused(self):
        r = self.gate(self.image([tree(FILES)], self.labels(revision="30981a52" * 5)))
        self.assertEqual(r.returncode, 1)
        self.assertIn("is not a commit of the remote (built from a commit that was never "
                      "pushed)", r.stdout)

    def test_revision_other_than_the_tag_commit_is_refused(self):
        self.repo.write("www/later.php", b"l\n")
        later = self.repo.commit("later")
        r = self.gate(self.image([tree(FILES)], self.labels(revision=later)))
        self.assertEqual(r.returncode, 1)
        self.assertIn(f"image revision {later[:12]} is not the commit of tag 'v1.0.0'",
                      r.stdout)

    def test_missing_revision_label_is_refused(self):
        r = self.gate(self.image([tree(FILES)], {"org.opencontainers.image.version": "v1.0.0"}))
        self.assertEqual(r.returncode, 1)
        self.assertIn("has no org.opencontainers.image.revision label", r.stdout)

    def test_tag_off_the_default_branch_is_refused(self):
        git(self.repo.d, "switch", "-q", "-c", "staging")
        self.repo.write("www/index.php", b"staging\n")
        side = self.repo.commit("side")
        self.repo.tag("v1.1.0")
        git(self.repo.d, "switch", "-q", "main")
        files = {**FILES, "index.php": b"staging\n"}
        man = self.tmp / "side.json"
        run_tool("manifest", "--config", self.cfg, "--repo", self.repo.d, "--tag", "v1.1.0",
                 "--out", man)
        archive = self.image([tree(files)], self.labels(revision=side, version="v1.1.0"))
        r = run_tool("gate", "--config", self.cfg, "--image-archive", archive,
                     "--git-remote", self.repo.d, "--manifest", man)
        self.assertEqual(r.returncode, 1)
        self.assertIn(f"tag 'v1.1.0' ({side[:12]}) is not on 'main'", r.stdout)

    def test_tag_missing_on_the_remote_is_refused_even_without_manifest(self):
        archive = self.image([tree(FILES)], self.labels(version="2026.07.29B"))
        r = self.gate(archive, manifest=False)
        self.assertEqual(r.returncode, 1, r.stdout + r.stderr)
        self.assertIn("release tag '2026.07.29B' does not exist on the remote", r.stdout)
        self.assertIn("content: not compared, no release manifest", r.stdout)
        self.assertIn("summary: 1 provenance, 1 not compared", r.stdout)

    # --- unmeasurable: fail closed ---

    def test_unreadable_image_is_unmeasurable(self):
        bogus = self.tmp / "bogus.tar"
        bogus.write_bytes(b"not a tarball")
        r = self.gate(bogus)
        self.assertEqual(r.returncode, 2)
        self.assertIn("cannot read image archive", r.stderr)

    def test_no_release_tag_named_is_unmeasurable(self):
        r = self.gate(self.image([tree(FILES)], self.labels(version=None)))
        self.assertEqual(r.returncode, 2)
        self.assertIn("no release tag", r.stderr)

    def test_unavailable_manifest_without_other_findings_is_unmeasurable(self):
        write_config(self.cfg, gh_cmd=[sys.executable, "-c", "import sys; sys.exit(1)"])
        r = self.gate(self.image([tree(FILES)]), manifest=False)
        self.assertEqual(r.returncode, 2, r.stdout)
        self.assertIn("downloading 'image-provenance-manifest.json' of release 'v1.0.0'",
                      r.stderr)

    # --- diff against the running image ---

    def test_diff_against_the_running_image_is_printed_without_changing_the_verdict(self):
        running = self.image([tree({**FILES, "feature.php": b"f\n"})], name="running.tar")
        r = self.gate(self.image([tree(FILES)]), "--current-archive", running)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertIn("DIFF AGAINST THE RUNNING IMAGE", r.stdout)
        self.assertIn("only in the running image (disappears): feature.php", r.stdout)

    def test_running_image_is_read_from_the_cluster_deployment(self):
        running = self.image([tree({"index.php": b"prod\n"})], name="running.tar")
        engine = ("import shutil, sys; a = sys.argv; "
                  f"shutil.copy({str(running)!r}, a[a.index('-o') + 1])")
        oc = "import sys; print('reg.example.test/team/app:prod', sys.argv[1:])"
        write_config(self.cfg, exclude=["vendor/"],
                     engine_cmd=[sys.executable, "-c", engine],
                     oc_cmd=[sys.executable, "-c", oc],
                     openshift={"namespace": "ns", "deployment": "app", "container": "web"})
        r = self.gate(self.image([tree(FILES)]), "--current-from-cluster")
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertIn("reg.example.test/team/app:prod", r.stdout)
        self.assertIn("deployment/app", r.stdout)
        self.assertIn("changed: index.php", r.stdout)
        self.assertIn("only in the new image: css/a.css", r.stdout)


if __name__ == "__main__":
    unittest.main()
