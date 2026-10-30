#!/usr/bin/env python3
# ci-pattern asset tests — image-provenance archive reader (#1389)
"""Suite of ``ImageTree``: the deployable tree and labels of an image tarball.

Both formats an engine saves are read (docker-archive and oci-archive), the
layers apply in order with their whiteouts, and the excluded paths stay out.
An archive that cannot be read, or that has no file under the deployable
path, is unmeasurable (``GateError``, exit 2 in the gate)."""
import sys
import tarfile
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import image_provenance as ip  # noqa: E402
from _fixtures import _tar_bytes, docker_archive, layer, oci_archive, sha  # noqa: E402

ROOT = "var/www/html/"
LABELS = {"org.opencontainers.image.revision": "abc", "org.opencontainers.image.version": "v1"}


class ImageTreeTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self):
        self._tmp.cleanup()

    def read(self, layers, fmt=docker_archive, exclude=(), raw=False):
        blobs = layers if raw else [layer(x) for x in layers]
        archive = fmt(self.tmp / "img.tar", blobs, LABELS)
        return ip.ImageTree(str(archive), "/var/www/html", list(exclude))

    def test_docker_archive_files_labels_and_line_ending_digest(self):
        img = self.read([{ROOT + "a.php": b"a\r\n", "etc/x": b"x\n", ROOT + "d": None}])
        self.assertEqual(img.labels, LABELS)
        self.assertEqual(img.files, {"a.php": sha(b"a\r\n")})
        self.assertEqual(img.lf["a.php"], sha(b"a\n"))

    def test_oci_archive_applies_layers_in_order(self):
        img = self.read([{ROOT + "a.php": b"old\n", ROOT + "b.php": b"b\n"},
                         {"./" + ROOT + "a.php": b"new\n"}], fmt=oci_archive)
        self.assertEqual(img.files, {"a.php": sha(b"new\n"), "b.php": sha(b"b\n")})

    def test_whiteouts_hide_lower_layer_files_and_directories(self):
        img = self.read([{ROOT + "keep.php": b"k\n", ROOT + "old.php": b"o\n",
                          ROOT + "tmp/x": b"x\n", ROOT + "tmp/y": b"y\n"},
                         {ROOT + ".wh.old.php": b"", ROOT + "tmp/.wh..wh..opq": b"",
                          ROOT + "tmp/y": b"y2\n"}])
        self.assertEqual(img.files, {"keep.php": sha(b"k\n"), "tmp/y": sha(b"y2\n")})

    def test_symlinks_hardlinks_and_exclusions(self):
        link = tarfile.TarInfo(ROOT + "copy.php")
        link.type, link.linkname = tarfile.LNKTYPE, ROOT + "a.php"
        data = layer({ROOT + "a.php": b"a\n", ROOT + "latest": "->a.php",
                      ROOT + "vendor/v.php": b"v\n"})
        extra = _tar_bytes([(link, None)])
        img = self.read([data, extra], exclude=["vendor/"], raw=True)
        self.assertEqual(img.files, {"a.php": sha(b"a\n"), "copy.php": sha(b"a\n"),
                                     "latest": "symlink:a.php"})

    def test_unreadable_archive_is_unmeasurable(self):
        bogus = self.tmp / "bogus.tar"
        bogus.write_bytes(b"not a tarball")
        with self.assertRaisesRegex(ip.GateError, "cannot read image archive"):
            ip.ImageTree(str(bogus), "/var/www/html", [])

    def test_archive_that_is_not_an_image_is_unmeasurable(self):
        plain = self.tmp / "plain.tar"
        plain.write_bytes(layer({ROOT + "a.php": b"a\n"}))
        with self.assertRaisesRegex(ip.GateError, "neither manifest.json"):
            ip.ImageTree(str(plain), "/var/www/html", [])

    def test_image_without_the_deployable_tree_is_unmeasurable(self):
        with self.assertRaisesRegex(ip.GateError, "has no files under '/var/www/html'"):
            self.read([{"etc/x": b"x\n"}])


if __name__ == "__main__":
    unittest.main()
