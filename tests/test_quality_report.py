"""The aggregator forwards the policy date, and trusts no gate.

``quality_report.py`` is the canon's asset (``deterministic-quality-harness`` v2.2) with BP's gate
list. Two properties are load-bearing: the date reaches the gates whose canon copy requires it and
only those, and a gate that lies — malformed envelope, someone else's identity, exit code that
disagrees with its verdict, evidence that no longer matches the tree — is an ``error``, never a
quiet pass.
"""

from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"


FIXTURES = Path(__file__).resolve().parent / "fixtures"
#: DQH 2.2: the date is explicit and recorded in the report.
POLICY_DATE = "2026-10-21"


def _load():
    # The canon's gates import their siblings by module name, so `scripts/` has to be importable.
    if str(SCRIPTS) not in sys.path:
        sys.path.insert(0, str(SCRIPTS))
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
        script = Path(command[1]).name
        commands[script] = command[2:]
        # A well-formed envelope for the gate being asked, so the aggregator's own validation is
        # not what fails here: this helper exists to capture the command line, not the verdict.
        gate = script.removesuffix(".py").removeprefix("check_")
        payload = {
            "gate": gate,
            "status": "pass",
            "indicators": {},
            "ceilings": {},
            "findings": [],
        }
        return subprocess.CompletedProcess(command, 0, json.dumps(payload), "")

    monkeypatch.setattr(module.subprocess, "run", fake_run)
    out = tmp_path / "report.json"
    module.main(["--root", str(tmp_path), "--out", str(out), *extra])
    return commands


def test_policy_date_reaches_only_the_dated_gates(monkeypatch, tmp_path):
    commands = _gate_commands(monkeypatch, tmp_path, "--policy-date", "2026-10-21")

    dated = {"check_complexity.py", "check_mutation_sites.py", "check_dry.py"}
    for script in dated:
        assert commands[script][-2:] == ["--policy-date", "2026-10-21"]
    others = [args for script, args in commands.items() if script not in dated]
    assert others and all("--policy-date" not in args for args in others)


def test_without_a_policy_date_the_aggregator_refuses_to_run(monkeypatch, tmp_path):
    """No date, no report: the same commit must not produce two verdicts (BP #1324)."""
    module = _load()
    out = tmp_path / "report.json"
    with pytest.raises(SystemExit) as exit_info:
        module.main(["--root", str(tmp_path), "--out", str(out)])
    assert exit_info.value.code == 2
    assert not out.exists(), "nothing was measured, so nothing was written"


# --------------------------------------------------------------------------------------------
# The aggregator trusts no gate (DQH 2.2)
# --------------------------------------------------------------------------------------------

#: Evidence every bound gate is checked against, stubbed so the only error a test produces is the
#: one it injected; with real evidence, every fake envelope would also fail the cross-check.
_EVIDENCE = {
    "candidate": {"commit": "abc", "tree": "def"},
    "policy_identity": "sha256:policy",
    "scope_identity": "sha256:scope",
    "subject_manifest": [],
    "tool": {"gate": "x", "command": [], "sha256": "sha256:tool"},
    "owner": "quality-architecture",
    "non_ownership": [],
}


def _envelope(gate: str, bound: bool) -> dict:
    payload: dict = {
        "gate": gate,
        "status": "pass",
        "indicators": {},
        "ceilings": {},
        "findings": [],
    }
    if bound:
        payload.update(
            {key: _EVIDENCE[key] for key in _EVIDENCE if key != "subject_manifest"}
            | {"subject_manifest": [], "subjects": {}}
        )
    return payload


def _gate(report: dict, script: str) -> dict:
    name = script.removesuffix(".py").removeprefix("check_")
    return next(item for item in report["gates"] if item["gate"] == name)


def _report_with(monkeypatch, tmp_path: Path, script: str, mutate, code: int = 0) -> dict:
    """Run the aggregator where `script` answers `mutate(envelope)` (None keeps it).

    `code` is the exit status that gate reports; every other gate answers cleanly.
    """
    module = _load()
    monkeypatch.setattr(module, "build_evidence", lambda *a, **k: _EVIDENCE)

    def fake_run(command, **_kwargs):
        if command[0] == "git":
            return subprocess.CompletedProcess(command, 0, "abc\n", "")
        name = Path(command[1]).name
        gate = name.removesuffix(".py").removeprefix("check_")
        envelope = _envelope(gate, gate in module.KNOWN_GATES)
        payload, exit_code = envelope, 0
        if name == script:
            replacement = mutate(envelope)
            payload = envelope if replacement is None else replacement
            exit_code = code
        return subprocess.CompletedProcess(
            command, exit_code, payload if isinstance(payload, str) else json.dumps(payload), ""
        )

    monkeypatch.setattr(module.subprocess, "run", fake_run)
    out = tmp_path / "report.json"
    module.main(["--root", str(tmp_path), "--out", str(out), "--policy-date", POLICY_DATE])
    return json.loads(out.read_text(encoding="utf-8"))


@pytest.mark.parametrize(
    "script, mutate, code, expected",
    [
        (
            "check_legacy_hashes.py",
            lambda envelope: json.dumps({"gate": "legacy_hashes"}),
            0,
            "missing required field",
        ),
        (
            "check_legacy_hashes.py",
            lambda envelope: json.dumps(_envelope("complexity", False)),
            0,
            "does not match expected",
        ),
        ("check_legacy_hashes.py", lambda envelope: None, 1, "disagrees with envelope status"),
        (
            "check_complexity.py",
            lambda envelope: {
                **envelope,
                "subject_manifest": [{"path": "app/gone.py", "sha256": "x"}],
            },
            0,
            "no longer matches current subjects",
        ),
    ],
    ids=["malformed", "wrong-identity", "exit-code-disagrees", "stale-evidence"],
)
def test_a_gate_that_lies_is_an_error_not_a_pass(
    monkeypatch, tmp_path: Path, script, mutate, code, expected
) -> None:
    report = _report_with(monkeypatch, tmp_path, script, mutate, code)
    entry = _gate(report, script)
    assert report["status"] == "fail"
    assert entry["status"] == "error"
    assert expected in entry["detail"]


def test_the_report_names_the_policy_date_and_the_evidence(tmp_path: Path) -> None:
    """A real run: the report carries the date and the evidence its gates were checked against."""
    module = _load()
    out = tmp_path / "report.json"
    code = module.main(
        [
            "--root",
            str(FIXTURES / "complexity_clean"),
            "--out",
            str(out),
            "--policy-date",
            POLICY_DATE,
        ]
    )
    report = json.loads(out.read_text(encoding="utf-8"))
    assert code == 0, report["gates"]
    assert report["policy_date"] == POLICY_DATE
    assert report["policy_identity"].startswith("sha256:")
    assert report["scope_identity"].startswith("sha256:")
    assert report["subject_manifest"]
    assert f"Policy date `{POLICY_DATE}`" in module.render_markdown(report)
