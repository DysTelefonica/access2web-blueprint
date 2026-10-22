"""BP's complexity ceiling (10) is policy data, and the gate evaluates it at an explicit date."""

from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
POLICY_DATE = "2026-10-21"


def _write_function(root: Path, complexity: int) -> None:
    package = root / "app" / "pkg"
    package.mkdir(parents=True)
    branches = "".join(f"    if value == {n}:\n        return {n}\n" for n in range(complexity - 1))
    source = f"def branchy(value: int) -> int:\n{branches}    return -1\n"
    (package / "branchy.py").write_text(source, encoding="utf-8")


def _run(root: Path, *extra: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(SCRIPTS / "check_complexity.py"), "--root", str(root), *extra],
        capture_output=True,
        text=True,
        check=False,
        encoding="utf-8",
    )


def _load():
    if str(SCRIPTS) not in sys.path:
        sys.path.insert(0, str(SCRIPTS))
    spec = importlib.util.spec_from_file_location(
        "check_complexity", SCRIPTS / "check_complexity.py"
    )
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    sys.modules["check_complexity"] = module
    spec.loader.exec_module(module)
    return module


def test_complexity_eleven_fails_against_the_ceiling_of_ten(tmp_path: Path) -> None:
    _write_function(tmp_path, 11)
    result = _run(tmp_path, "--policy-date", POLICY_DATE)
    assert result.returncode == 1, result.stdout + result.stderr
    assert "branchy has complexity 11, ceiling is 10" in result.stdout


def test_complexity_ten_passes(tmp_path: Path) -> None:
    _write_function(tmp_path, 10)
    result = _run(tmp_path, "--policy-date", POLICY_DATE)
    assert result.returncode == 0, result.stdout + result.stderr
    assert result.stdout == "OK    every function at or below complexity 10\n"


def test_the_policy_date_is_required(tmp_path: Path) -> None:
    _write_function(tmp_path, 2)
    result = _run(tmp_path)
    assert result.returncode == 2
    assert "--policy-date" in result.stderr


def test_envelope_is_bound_to_the_policy_and_its_date(tmp_path: Path) -> None:
    _write_function(tmp_path, 10)
    result = _run(tmp_path, "--policy-date", POLICY_DATE, "--json")
    envelope = json.loads(result.stdout)
    assert result.returncode == 0
    assert envelope["policy_date"] == POLICY_DATE
    assert envelope["ceilings"] == {"max_complexity": 10, "functions_over_ceiling": 0}
    assert envelope["subjects"]["checked"] == ["app/pkg/branchy.py"]
    assert envelope["policy_identity"].startswith("sha256:")


def test_an_expired_baseline_fails_and_a_stale_entry_is_noted(tmp_path: Path, capsys) -> None:
    module = _load()
    _write_function(tmp_path, 11)
    key = "app/pkg/branchy.py::branchy"
    module.BASELINE = {
        key: module.BaselineEntry(complexity=11, target=10, target_date="2026-10-20"),
        "app/gone.py::f": module.BaselineEntry(complexity=12, target=10, target_date="2027-01-01"),
    }
    assert module.main(["--root", str(tmp_path), "--policy-date", "2026-10-20"]) == 0
    assert module.main(["--root", str(tmp_path), "--policy-date", POLICY_DATE]) == 1
    out = capsys.readouterr().out
    assert "branchy BASELINE expired on 2026-10-20 at complexity 11, target was 10" in out
    assert "NOTE  app/gone.py::f: now under the ceiling; remove it from BASELINE" in out


def test_the_repository_baseline_is_empty() -> None:
    assert _load().BASELINE == {}
