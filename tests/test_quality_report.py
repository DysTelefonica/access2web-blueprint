"""The aggregator: it forwards the policy date to the gates that require it, and it trusts no gate.

``quality_report.py`` is the canon's asset (``deterministic-quality-harness`` v2.2) with BP's gate
list. Two properties are load-bearing and are pinned here:

  1. The date reaches the gates whose canon copy requires it, and only those.
  2. A gate that lies — a malformed envelope, an envelope for another gate, an exit code that
     disagrees with its own verdict, or evidence that no longer matches the tree — is an ``error``,
     never a quiet pass. An aggregator that cannot tell a broken gate from a clean one reports a
     green that nobody earned.
"""

from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
from datetime import UTC, datetime
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


@pytest.mark.parametrize(
    "extra, expected",
    [
        (("--policy-date", "2026-10-21"), "2026-10-21"),
        ((), datetime.now(UTC).date().isoformat()),
    ],
)
def test_policy_date_reaches_only_the_dated_gates(monkeypatch, tmp_path, extra, expected):
    commands = _gate_commands(monkeypatch, tmp_path, *extra)

    dated = {"check_complexity.py", "check_mutation_sites.py", "check_dry.py"}
    for script in dated:
        assert commands[script][-2:] == ["--policy-date", expected]
    others = [args for script, args in commands.items() if script not in dated]
    assert others and all("--policy-date" not in args for args in others)


# --------------------------------------------------------------------------------------------
# The aggregator trusts no gate (DQH 2.2)
# --------------------------------------------------------------------------------------------

#: Evidence every bound gate is checked against. Stubbed so that the only error a test produces is
#: the one it injected: with real evidence, every fake envelope would also fail the cross-check.
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
            {
                "candidate": _EVIDENCE["candidate"],
                "policy_identity": _EVIDENCE["policy_identity"],
                "scope_identity": _EVIDENCE["scope_identity"],
                "subject_manifest": _EVIDENCE["subject_manifest"],
                "tool": _EVIDENCE["tool"],
                "owner": _EVIDENCE["owner"],
                "non_ownership": _EVIDENCE["non_ownership"],
                "subjects": {},
            }
        )
    return payload


def _report_with(monkeypatch, tmp_path: Path, script: str, payload: str | None, code: int = 0):
    """Run the aggregator where `script` answers `payload`, and return its report."""
    module = _load()
    monkeypatch.setattr(module, "build_evidence", lambda *a, **k: _EVIDENCE)

    def fake_run(command, **_kwargs):
        if command[0] == "git":
            return subprocess.CompletedProcess(command, 0, "abc\n", "")
        name = Path(command[1]).name
        gate = name.removesuffix(".py").removeprefix("check_")
        bound = gate in module.KNOWN_GATES
        body = (
            payload
            if (name == script and payload is not None)
            else json.dumps(_envelope(gate, bound))
        )
        return subprocess.CompletedProcess(command, code, body, "")

    monkeypatch.setattr(module.subprocess, "run", fake_run)
    out = tmp_path / "report.json"
    module.main(["--root", str(tmp_path), "--out", str(out), "--policy-date", POLICY_DATE])
    return json.loads(out.read_text(encoding="utf-8"))


def _gate(report: dict, name: str) -> dict:
    return next(item for item in report["gates"] if item["gate"] == name)


def test_a_malformed_envelope_is_an_error_not_a_pass(monkeypatch, tmp_path: Path) -> None:
    report = _report_with(
        monkeypatch,
        tmp_path,
        "check_legacy_hashes.py",
        json.dumps({"gate": "legacy_hashes"}),
    )
    assert report["status"] == "fail"
    entry = _gate(report, "legacy_hashes")
    assert entry["status"] == "error"
    assert "missing required field" in entry["detail"]


def test_an_envelope_for_another_gate_is_rejected(monkeypatch, tmp_path: Path) -> None:
    """A gate answering with someone else's name has not measured what the report claims."""
    report = _report_with(
        monkeypatch,
        tmp_path,
        "check_legacy_hashes.py",
        json.dumps(_envelope("complexity", False)),
    )
    assert report["status"] == "fail"
    entry = _gate(report, "legacy_hashes")
    assert entry["status"] == "error"
    assert "does not match expected" in entry["detail"]


def test_an_exit_code_that_disagrees_with_the_envelope_is_an_error(
    monkeypatch, tmp_path: Path
) -> None:
    report = _report_with(
        monkeypatch,
        tmp_path,
        "check_legacy_hashes.py",
        None,
        code=1,
    )
    assert report["status"] == "fail"
    entry = _gate(report, "legacy_hashes")
    assert entry["status"] == "error"
    assert "disagrees with envelope status" in entry["detail"]


def test_evidence_that_no_longer_matches_the_tree_is_an_error(monkeypatch, tmp_path: Path) -> None:
    """A gate that measured yesterday's tree cannot vouch for today's report.

    The gate answers with a subject manifest from a different tree; the aggregator recomputes the
    manifest from the policy and refuses the evidence instead of copying it into the report.
    """
    module = _load()
    stale = dict(_envelope("complexity", True))
    stale["subject_manifest"] = [{"path": "app/gone.py", "sha256": "sha256:gone"}]
    current = {**_EVIDENCE, "subject_manifest": [{"path": "app/here.py", "sha256": "sha256:here"}]}
    monkeypatch.setattr(module, "build_evidence", lambda *a, **k: current)

    def fake_run(command, **_kwargs):
        if command[0] == "git":
            return subprocess.CompletedProcess(command, 0, "abc\n", "")
        name = Path(command[1]).name
        gate = name.removesuffix(".py").removeprefix("check_")
        body = (
            stale if name == "check_complexity.py" else _envelope(gate, gate in module.KNOWN_GATES)
        )
        return subprocess.CompletedProcess(command, 0, json.dumps(body), "")

    monkeypatch.setattr(module.subprocess, "run", fake_run)
    out = tmp_path / "report.json"
    module.main(["--root", str(tmp_path), "--out", str(out), "--policy-date", POLICY_DATE])
    report = json.loads(out.read_text(encoding="utf-8"))
    entry = _gate(report, "complexity")
    assert entry["status"] == "error"
    assert "no longer matches current subjects" in entry["detail"]


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
