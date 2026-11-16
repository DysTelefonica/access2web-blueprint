#!/usr/bin/env python3
"""Tests del aprovisionamiento del runner nocturno (#1627, slice 6). stdlib.

Se ejecutan en WSL (el script es bash): PATH con herramientas falsas y nombres de
herramienta que no existen en la máquina, para que «falta» sea determinista.
"""

from __future__ import annotations

import re
import subprocess
import tempfile
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
SCRIPT = HERE.parent / "provision-night-runner.sh"
PINS = HERE.parent / "night-runner.pins"
REPO = HERE.parents[5]
WORKFLOW = REPO / ".github" / "workflows" / "tests.yml"
# El catálogo es personal/ardelperal/ci-pattern; un consumer vendoriza la skill en otra ruta y no
# tiene tests.yml del catálogo (#1700). Solo en el catálogo su ausencia es un fallo.
CI_PATTERN = HERE.parents[2]
IN_CATALOGUE = CI_PATTERN.parent.name == "ardelperal" and CI_PATTERN.parent.parent.name == "personal"


def stub(bin_dir: Path, name: str, version_line: str) -> None:
    path = bin_dir / name
    path.write_text(f"#!/bin/sh\necho '{version_line}'\n", encoding="utf-8")
    path.chmod(0o755)


class ProvisionTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.bin = self.root / "bin"
        self.bin.mkdir()
        self.pins = self.root / "pins.txt"
        self.pins.write_text("# fixture\nnightstub-a=1.2.3\nnightstub-b=4.5.6\n", encoding="utf-8")
        stub(self.bin, "nightstub-a", "nightstub-a version 1.2.3")
        stub(self.bin, "nightstub-b", "nightstub-b version 4.5.6")
        self.runner = self.root / "night-runner"

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def run_script(self, *args: str, pins: Path | None = None) -> subprocess.CompletedProcess:
        env = {"PATH": f"{self.bin}:/usr/bin:/bin", "LC_ALL": "C"}
        return subprocess.run(
            ["bash", str(SCRIPT), "--runner-dir", str(self.runner), "--pins", str(pins or self.pins), *args],
            capture_output=True, text=True, env=env, check=False)

    def snapshot(self) -> dict[str, bytes]:
        if not self.runner.exists():
            return {}
        return {str(p.relative_to(self.runner)): p.read_bytes()
                for p in sorted(self.runner.rglob("*")) if p.is_file()}

    def test_a_wrong_version_exits_two_naming_tool_found_and_expected(self) -> None:
        stub(self.bin, "nightstub-a", "nightstub-a version 9.9.9")
        result = self.run_script()
        self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
        self.assertIn("nightstub-a", result.stderr)
        self.assertIn("9.9.9", result.stderr)
        self.assertIn("1.2.3", result.stderr)

    def test_a_missing_tool_exits_two_naming_it(self) -> None:
        self.pins.write_text("nightstub-absent=1.0.0\n", encoding="utf-8")
        result = self.run_script()
        self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
        self.assertIn("nightstub-absent", result.stderr)

    def test_a_failed_check_writes_nothing(self) -> None:
        stub(self.bin, "nightstub-b", "nightstub-b version 0.0.1")
        self.run_script()
        self.assertEqual(self.snapshot(), {})

    def test_dry_run_writes_nothing_and_says_what_it_would_do(self) -> None:
        result = self.run_script("--dry-run")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("ensayo", result.stdout)
        self.assertFalse(self.runner.exists())

    def test_provisioning_creates_the_directory_and_its_marker(self) -> None:
        result = self.run_script()
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(self.snapshot(), {".night-runner": b"profile=1\n"})

    def test_a_second_run_changes_nothing(self) -> None:
        self.assertEqual(self.run_script().returncode, 0)
        before = self.snapshot()
        result = self.run_script()
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("sin cambios", result.stdout)
        self.assertEqual(self.snapshot(), before)

    def test_a_marker_of_another_profile_exits_two_and_is_not_overwritten(self) -> None:
        self.runner.mkdir()
        (self.runner / ".night-runner").write_text("profile=7\n", encoding="utf-8")
        result = self.run_script()
        self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
        self.assertIn("profile=7", result.stderr)
        self.assertEqual((self.runner / ".night-runner").read_text(encoding="utf-8"), "profile=7\n")

    def test_an_unknown_argument_exits_two(self) -> None:
        self.assertEqual(self.run_script("--bogus").returncode, 2)

    def test_without_runner_dir_it_exits_two(self) -> None:
        env = {"PATH": f"{self.bin}:/usr/bin:/bin", "LC_ALL": "C"}
        result = subprocess.run(["bash", str(SCRIPT), "--pins", str(self.pins)],
                                capture_output=True, text=True, env=env, check=False)
        self.assertEqual(result.returncode, 2)


class ShippedPinsTests(unittest.TestCase):
    def test_the_night_pins_match_the_ci_python_and_node(self) -> None:
        if not IN_CATALOGUE and not WORKFLOW.is_file():
            self.skipTest("consumer without the catalogue's tests.yml: the pins are checked in the catalogue (#1700)")
        pins = dict(re.findall(r"^([A-Za-z0-9_.-]+)=(\S+)$", PINS.read_text(encoding="utf-8"), re.MULTILINE))
        workflow = WORKFLOW.read_text(encoding="utf-8")
        self.assertIn(f"python-version: '{pins['python3']}'", workflow)
        self.assertIn(f"node-version: '{pins['node']}'", workflow)


if __name__ == "__main__":
    unittest.main()
