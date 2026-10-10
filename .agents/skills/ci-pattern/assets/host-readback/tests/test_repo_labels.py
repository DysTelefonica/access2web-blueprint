#!/usr/bin/env python3
"""#237: check_host_drift detecta la ausencia de chain:partial en el host.

Usa el contrato del propio repo (.github/host-contract.json) y fixtures
de snapshots de etiquetas. El snapshot sin chain:partial produce un
hallazgo de drift; con ella, pasa. RED: el contrato no existe aún.
"""

import json
import sys
import subprocess
import tempfile
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[6]
DRIFT = REPO_ROOT / "personal" / "ardelperal" / "ci-pattern" / "assets" / "host-readback" / "check_host_drift.py"
CONTRACT = REPO_ROOT / ".github" / "host-contract.json"

SKILL = "personal/ardelperal/ci-pattern"

GOV_LABELS = [
    {"name": "status:approved", "class": "host-enforced"},
    {"name": "chain:partial", "class": "host-enforced"},
    {"name": "type:bug", "class": "host-enforced"},
    {"name": "type:feature", "class": "host-enforced"},
    {"name": "type:docs", "class": "host-enforced"},
    {"name": "type:chore", "class": "host-enforced"},
    {"name": "type:refactor", "class": "host-enforced"},
]


def _run_drift(labels_snapshot):
    with tempfile.TemporaryDirectory() as tmp:
        snap_dir = Path(tmp)
        (snap_dir / "labels.json").write_text(json.dumps(labels_snapshot), encoding="utf-8")
        for kind, default in (("repo", "{}"), ("branch-protection", "{}"), ("rulesets", "[]")):
            (snap_dir / f"{kind}.json").write_text(default, encoding="utf-8")
        return subprocess.run(
            [sys.executable, str(DRIFT),
             "--contract", str(CONTRACT),
             "--snapshot", f"repo={snap_dir / 'repo.json'}",
             "--snapshot", f"labels={snap_dir / 'labels.json'}",
             "--snapshot", f"branch-protection={snap_dir / 'branch-protection.json'}",
             "--snapshot", f"rulesets={snap_dir / 'rulesets.json'}"],
            capture_output=True, text=True, check=False,
        )


class RepoHostContractTests(unittest.TestCase):

    def test_contract_exists_and_declares_chain_partial(self):
        """El contrato existe y declara chain:partial como host-enforced."""
        self.assertTrue(CONTRACT.is_file(), f".github/host-contract.json no existe")
        contract = json.loads(CONTRACT.read_text(encoding="utf-8"))
        labels = {l["name"]: l["class"] for l in contract.get("labels", [])}
        self.assertEqual(labels.get("chain:partial"), "host-enforced")
        self.assertEqual(labels.get("status:approved"), "host-enforced")

    @staticmethod
    def _contract_labels():
        contract = json.loads(CONTRACT.read_text(encoding="utf-8"))
        return [l["name"] for l in contract.get("labels", []) if l.get("class") == "host-enforced"]

    def test_drift_without_chain_partial_produces_finding(self):
        """Snapshot sin chain:partial → drift (hallazgo HR-34)."""
        self.assertTrue(CONTRACT.is_file(), "ejecutar después de crear el contrato")
        labels = [n for n in self._contract_labels() if n != "chain:partial"]
        snapshot = [{"name": n} for n in labels]
        proc = _run_drift(snapshot)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("chain:partial", proc.stdout)

    def test_drift_with_chain_partial_passes(self):
        """Snapshot con chain:partial → sin drift para labels."""
        self.assertTrue(CONTRACT.is_file(), "ejecutar después de crear el contrato")
        labels = self._contract_labels()
        snapshot = [{"name": n} for n in labels]
        proc = _run_drift(snapshot)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
