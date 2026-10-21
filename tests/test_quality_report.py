"""quality_report forwards the policy date to the complexity gate, and only to it."""

from __future__ import annotations

import importlib.util
import subprocess
import sys
from datetime import UTC, datetime
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"


def _load():
    spec = importlib.util.spec_from_file_location("quality_report", SCRIPTS / "quality_report.py")
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    sys.modules["quality_report"] = module
    spec.loader.exec_module(module)
    return module


def _gate_commands(monkeypatch, tmp_path: Path, *extra: str) -> dict[str, list[str]]:
    module = _load()
    commands: dict[str, list[str]] = {}

    def fake_run(command, **_kwargs):
        if command[0] == "git":
            return subprocess.CompletedProcess(command, 0, "abc\n", "")
        commands[Path(command[1]).name] = command[2:]
        return subprocess.CompletedProcess(command, 0, '{"gate": "x", "status": "pass"}', "")

    monkeypatch.setattr(module.subprocess, "run", fake_run)
    out = tmp_path / "report.json"
    assert module.main(["--root", str(tmp_path), "--out", str(out), *extra]) == 0
    return commands


@pytest.mark.parametrize(
    "extra, expected",
    [
        (("--policy-date", "2026-10-21"), "2026-10-21"),
        ((), datetime.now(UTC).date().isoformat()),
    ],
)
def test_policy_date_reaches_only_the_complexity_gate(monkeypatch, tmp_path, extra, expected):
    commands = _gate_commands(monkeypatch, tmp_path, *extra)

    assert commands["check_complexity.py"][-2:] == ["--policy-date", expected]
    others = [args for script, args in commands.items() if script != "check_complexity.py"]
    assert others and all("--policy-date" not in args for args in others)
