# ci-pattern asset tests — ci-red: gh simulado compartido (#1508, HR-89)
"""``gh`` simulado (issues en un JSON) para ``report_ci_red.py``; nunca el real.

Windows solo resuelve extensiones de ``PATHEXT``: el doble lleva ``gh.cmd``.
"""
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
SCRIPT = ASSET / "report_ci_red.py"

STUB = r'''
import json, os, sys

state_path = os.environ["GH_STUB_STATE"]
with open(state_path, encoding="utf-8") as fh:
    state = json.load(fh)
argv = sys.argv[1:]
state["calls"].append(argv)


def opt(name):
    return argv[argv.index(name) + 1] if name in argv else None


def save():
    with open(state_path, "w", encoding="utf-8") as fh:
        json.dump(state, fh)


fail = state.get("fail")
if fail and argv[:2] == fail:
    save()
    sys.stderr.write("gh: HTTP 502\n")
    sys.exit(1)
if argv[:2] == ["issue", "list"]:
    label = opt("--label")
    rows = [{"number": i["number"], "title": i["title"]} for i in state["issues"]
            if i["state"] == "open" and (label is None or label in i["labels"])]
    print(json.dumps(rows))
elif argv[:2] == ["issue", "create"]:
    number = state["next"]
    state["next"] += 1
    with open(opt("--body-file"), encoding="utf-8") as fh:
        body = fh.read()
    state["issues"].append({"number": number, "title": opt("--title"), "state": "open",
                            "labels": [opt("--label")], "assignees": [opt("--assignee")],
                            "body": body, "comments": []})
    print(f"https://github.com/{opt('--repo')}/issues/{number}")
elif argv[:2] in (["issue", "comment"], ["issue", "close"]):
    issue = next(i for i in state["issues"] if i["number"] == int(argv[2]))
    if argv[1] == "comment":
        with open(opt("--body-file"), encoding="utf-8") as fh:
            issue["comments"].append(fh.read())
    else:
        issue["state"] = "closed"
else:
    save()
    sys.stderr.write("gh stub: unexpected call %r\n" % argv)
    sys.exit(3)
save()
'''

CONTRACT = {
    "contract_version": 1,
    "coordinator": "coordinador",
    "labels": [
        {"name": "chain:partial", "class": "host-enforced"},
        {"name": "ci-red", "class": "host-enforced"},
    ],
}

SCHEDULED_ON_MAIN = {
    "GITHUB_REPOSITORY": "example/consumer",
    "GITHUB_EVENT_NAME": "schedule",
    "GITHUB_REF_NAME": "main",
    "GITHUB_WORKFLOW": "main-audit",
    "GITHUB_SERVER_URL": "https://github.com",
    "GITHUB_RUN_ID": "4242",
    "GITHUB_SHA": "0123456789abcdef0123456789abcdef01234567",
}


class Host:
    """Directorio temporal con el ``gh`` simulado, su estado y el contrato."""

    def __init__(self, root: Path, issues=None, contract=CONTRACT, fail=None):
        self.root = root
        self.bin = root / "bin"
        self.bin.mkdir()
        stub = self.bin / "gh_stub.py"
        stub.write_text(STUB, encoding="utf-8")
        if os.name == "nt":
            (self.bin / "gh.cmd").write_text(
                f'@"{sys.executable}" "{stub}" %*\r\n', encoding="utf-8")
        else:
            launcher = self.bin / "gh"
            launcher.write_text(
                f'#!/bin/sh\nexec "{sys.executable}" "{stub}" "$@"\n', encoding="utf-8")
            launcher.chmod(0o755)
        self.state_path = root / "state.json"
        issues = issues or []
        self.state_path.write_text(json.dumps({
            "issues": issues,
            "next": max([i["number"] for i in issues], default=0) + 1,
            "calls": [],
            "fail": fail,
        }), encoding="utf-8")
        self.contract = root / "host-contract.json"
        self.contract.write_text(
            contract if isinstance(contract, str) else json.dumps(contract), encoding="utf-8")

    def env(self, **overrides):
        env = {k: v for k, v in os.environ.items() if not k.startswith("GITHUB_")}
        env.update(SCHEDULED_ON_MAIN)
        env.update(overrides)
        env["PATH"] = str(self.bin) + os.pathsep + env.get("PATH", "")
        env["GH_STUB_STATE"] = str(self.state_path)
        # Windows escribe la consola en cp1252; gh y los runners, en UTF-8.
        env["PYTHONIOENCODING"] = "utf-8"
        return env

    def run(self, outcome, *extra, **env_overrides):
        argv = [sys.executable, str(SCRIPT), "--outcome", outcome,
                "--default-branch", "main", "--host-contract", str(self.contract), *extra]
        return subprocess.run(argv, capture_output=True, text=True, encoding="utf-8",
                              env=self.env(**env_overrides), check=False)

    @property
    def state(self):
        return json.loads(self.state_path.read_text(encoding="utf-8"))


def open_issue(number, title, labels=("ci-red",)):
    return {"number": number, "title": title, "state": "open", "labels": list(labels),
            "assignees": ["coordinador"], "body": "", "comments": []}


class HostTestCase(unittest.TestCase):
    def host(self, **kwargs) -> Host:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        return Host(Path(tmp.name), **kwargs)
