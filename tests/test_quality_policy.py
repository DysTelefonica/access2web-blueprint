"""BP's quality-policy.json governs only the gates BP runs, with BP's ceilings."""

from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"


def _load_quality_policy():
    if str(SCRIPTS) not in sys.path:
        sys.path.insert(0, str(SCRIPTS))
    spec = importlib.util.spec_from_file_location("quality_policy", SCRIPTS / "quality_policy.py")
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    sys.modules["quality_policy"] = module
    spec.loader.exec_module(module)
    return module


def _write_subjects(root: Path) -> None:
    (root / "app" / "pkg" / "migrations").mkdir(parents=True)
    (root / "app" / "pkg" / "b.py").write_text("B = 2\n", encoding="utf-8")
    (root / "app" / "pkg" / "a.py").write_text("A = 1\n", encoding="utf-8")
    (root / "app" / "pkg" / "migrations" / "ignored.py").write_text("X = 0\n", encoding="utf-8")


def test_policy_declares_only_the_gates_bp_runs() -> None:
    policy = _load_quality_policy().load_policy(ROOT / "quality-policy.json")

    assert set(policy["gates"]) == {"complexity", "dry", "mutation_sites", "mutation"}
    assert policy["gates"]["complexity"]["ceilings"] == {
        "max_complexity": 10,
        "functions_over_ceiling": 0,
    }


def test_policy_manifest_is_ordered_and_excludes_migrations(tmp_path: Path) -> None:
    module = _load_quality_policy()
    _write_subjects(tmp_path)
    policy = module.load_policy(ROOT / "quality-policy.json")

    evidence = module.build_evidence(tmp_path, policy, "complexity")

    assert [item["path"] for item in evidence["subject_manifest"]] == [
        "app/pkg/a.py",
        "app/pkg/b.py",
    ]
    assert evidence["tool"]["command"] == ["python", "scripts/check_complexity.py"]


def test_an_undeclared_gate_is_not_governed(tmp_path: Path) -> None:
    module = _load_quality_policy()
    _write_subjects(tmp_path)
    policy = module.load_policy(ROOT / "quality-policy.json")

    with pytest.raises(module.PolicyError, match="'layers' is not governed"):
        module.build_evidence(tmp_path, policy, "layers")


def test_drifted_ceilings_are_rejected(tmp_path: Path) -> None:
    module = _load_quality_policy()
    _write_subjects(tmp_path)
    envelope = {"gate": "complexity", "status": "pass", "ceilings": {"max_complexity": 15}}

    with pytest.raises(module.PolicyError, match="ceilings drifted"):
        module.bind_envelope(envelope, tmp_path, "complexity", checked=[])


def test_policy_rejects_a_gate_without_non_ownership(tmp_path: Path) -> None:
    module = _load_quality_policy()
    payload = json.loads((ROOT / "quality-policy.json").read_text(encoding="utf-8"))
    del payload["gates"]["dry"]["non_ownership"]
    policy_path = tmp_path / "quality-policy.json"
    policy_path.write_text(json.dumps(payload), encoding="utf-8")

    with pytest.raises(module.PolicyError, match="non_ownership"):
        module.load_policy(policy_path)
