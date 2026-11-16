#!/usr/bin/env python3
"""Tests de la instantánea nocturna: carga y validación (#1627, slice 4).

Cada caso del contrato sale con 2, nombra la clave o el campo y no imprime traceback.
"""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
SCRIPT = HERE.parent / "snapshot.py"


def snapshot() -> dict:
    return {
        "repository": "DysTelefonica/team-skills",
        "night": {"sha": "b" * 40, "started_at": "2026-11-13T02:00:00+00:00"},
        "last_good_night": {"sha": "a" * 40, "finished_at": "2026-11-12T04:00:00+00:00"},
        "prs": [{"number": 1001, "title": "fix(a): uno", "merged_at": "2026-11-12T10:00:00+00:00",
                 "merge_sha": "c" * 40, "issues": [900]}],
        "jobs": [{"name": "full", "expensive": True, "conclusion": "success",
                  "run_url": "https://ci/full"}],
        "deploy": {"deployed": False, "sha": "b" * 40, "at": None},
    }


INVALID_CASES = (
    ("job sin expensive", lambda s: s["jobs"][0].pop("expensive"), "expensive"),
    ("job sin conclusion", lambda s: s["jobs"][0].pop("conclusion"), "conclusion"),
    ("deploy sin deployed", lambda s: s.update(deploy={"at": None}), "deployed"),
    ("last_good_night sin finished_at", lambda s: s.update(last_good_night={"sha": "a" * 40}),
     "finished_at"),
    ("night sin started_at", lambda s: s.update(night={"sha": "b" * 40}), "started_at"),
    ("fecha no valida", lambda s: s["prs"][0].update(merged_at="ayer por la tarde"), "merged_at"),
    ("fecha sin zona horaria",
     lambda s: s.update(night={"sha": "b" * 40, "started_at": "2026-11-13T02:00:00"}), "started_at"),
    ("fichero ausente", None, "no-existe.json"),
)


class SnapshotContractTests(unittest.TestCase):
    def run_check(self, path: Path) -> subprocess.CompletedProcess:
        return subprocess.run([sys.executable, str(SCRIPT), "--snapshot", str(path)],
                              capture_output=True, text=True, check=False)

    def test_each_contract_case_exits_two_naming_the_field_without_traceback(self) -> None:
        for label, break_it, name in INVALID_CASES:
            with self.subTest(case=label), tempfile.TemporaryDirectory() as tmp:
                path = Path(tmp) / ("no-existe.json" if break_it is None else "snap.json")
                if break_it is not None:
                    snap = snapshot()
                    break_it(snap)
                    path.write_text(json.dumps(snap), encoding="utf-8")
                result = self.run_check(path)
                self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
                self.assertIn(name, result.stderr)
                self.assertNotIn("Traceback", result.stderr)

    def test_a_valid_snapshot_exits_zero(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "snap.json"
            path.write_text(json.dumps(snapshot()), encoding="utf-8")
            result = self.run_check(path)
        self.assertEqual(result.returncode, 0, result.stderr)


if __name__ == "__main__":
    unittest.main()
