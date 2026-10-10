#!/usr/bin/env python3
# ci-pattern asset tests — runner binding (DysTelefonica/team-skills#303)
"""Runner binding suite (black-box): every runner-enforced area declared in
the contract must be enforced by a workflow that runs on the declared
runner label and invokes the covering asset, and the three compensating
assets must all be wired on that runner. Missing wiring is exit 1 naming
the gap; unscannable input fails closed (exit 2). No gh, no network."""
import json
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent / "check_runner_binding.py"
import subprocess
import sys

LABEL = "cadete"
GUARD = "check_force_push_guard.py"
DETECTIVE = "check_push_compliance.py"
MERGE = "check_governed_merge.py"


def contract(**over: object) -> dict:
    base: dict = {"contract_version": 1, "protected_branches": ["main"],
                  "required_checks": [{"name": "unit", "class": "runner-enforced"}],
                  "protection": {"enforce_admins": {"declared": True, "class": "runner-enforced"}},
                  "runner": {"label": LABEL}}
    base.update(over)
    return base


def workflow(name: str, asset: str, label: str = LABEL, extra_asset: str | None = None) -> str:
    steps = f"      - run: python3 personal/ardelperal/ci-pattern/assets/runner-controls/{asset} --x 1\n"
    if extra_asset:
        steps += f"      - run: python3 personal/ardelperal/ci-pattern/assets/runner-controls/{extra_asset} --x 1\n"
    return (f"name: {name}\non: push\njobs:\n  job:\n"
            f"    runs-on: [self-hosted, {label}]\n    steps:\n{steps}")


def write_workflows(root: Path, files: dict[str, str]) -> Path:
    wdir = root / "workflows"
    wdir.mkdir(exist_ok=True)
    for name, text in files.items():
        (wdir / name).write_text(text, encoding="utf-8")
    return wdir


def run_gate(root: Path, contract_data: dict, files: dict[str, str]) -> subprocess.CompletedProcess:
    cpath = root / "contract.json"
    cpath.write_text(json.dumps(contract_data), encoding="utf-8")
    wdir = write_workflows(root, files)
    return subprocess.run([sys.executable, str(ASSET), "--contract", str(cpath),
                           "--workflows", str(wdir)], capture_output=True, text=True, timeout=60)


class RunnerBindingTests(unittest.TestCase):
    """#303 punto 5: la declaracion runner-enforced sin workflow es
    autodeclaracion; el gate exige el binding ejecutado."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.addCleanup(self._tmp.cleanup)

    def test_all_assets_wired_on_the_runner_exits_zero(self) -> None:
        proc = run_gate(self.root, contract(), {
            "guard.yml": workflow("guard", GUARD, extra_asset=DETECTIVE),
            "merge.yml": workflow("merge", MERGE)})
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("BINDING OK", proc.stdout)

    def test_missing_asset_for_declared_area_is_a_violation(self) -> None:
        # protection runner-enforced sin workflow del guard: autodeclaracion.
        proc = run_gate(self.root, contract(), {"merge.yml": workflow("merge", MERGE)})
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn(GUARD, proc.stdout + proc.stderr)
        self.assertIn(LABEL, proc.stdout + proc.stderr)

    def test_workflow_on_wrong_runner_does_not_count(self) -> None:
        files = {"guard.yml": workflow("guard", GUARD, label="otro-runner", extra_asset=DETECTIVE),
                 "merge.yml": workflow("merge", MERGE)}
        proc = run_gate(self.root, contract(), files)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn(GUARD, proc.stdout + proc.stderr)

    def test_missing_compensating_asset_is_a_violation(self) -> None:
        # areas cubiertas pero sin cablear el detective: nada marca el bypass.
        files = {"guard.yml": workflow("guard", GUARD), "merge.yml": workflow("merge", MERGE)}
        proc = run_gate(self.root, contract(), files)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn(DETECTIVE, proc.stdout + proc.stderr)

    def test_fail_closed_input(self) -> None:
        bad = contract()
        del bad["runner"]
        proc = run_gate(self.root, bad, {"guard.yml": workflow("guard", GUARD, extra_asset=DETECTIVE),
                                         "merge.yml": workflow("merge", MERGE)})
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        cpath = self.root / "contract.json"
        cpath.write_text(json.dumps(contract()), encoding="utf-8")
        proc = subprocess.run([sys.executable, str(ASSET), "--contract", str(cpath),
                               "--workflows", str(self.root / "nope")],
                              capture_output=True, text=True, timeout=60)
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("fail", (proc.stdout + proc.stderr).lower())


if __name__ == "__main__":
    unittest.main()
