#!/usr/bin/env python3
"""Suite de `ci-pattern pr contract` (#912).

El contrato de PR se corre entero en local contra un fichero de cuerpo, sin
token ni red. Cada caso monta un repo git temporal con una rama de feature,
ejecuta el comando y compara con el gate real (`check_pr_contract.py`) corrido
como proceso aparte sobre el mismo evento y los mismos datos del rango.
"""

from __future__ import annotations

import importlib.machinery
import importlib.util
import io
import json
import os
import subprocess
import sys
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path

ASSETS = Path(__file__).resolve().parents[2]
LOCAL = Path(__file__).resolve().parents[1] / "pr_contract_local.py"
GATE = ASSETS / "pr-contract" / "check_pr_contract.py"
BRANCH = "feat/912-pr-contract-local"
TEST_REL = "tests/test_fix.py"

BODY = """Closes #912

## Chain Context

- chain: 912-pr-contract-local
- position: 1/1
- base: main
- depends-on: none
- follow-up: none
- starts-at: {branch}
- ends-with: {branch}
- review-budget: {budget}/400

📍 este PR

## Asistencia de IA

- Asistencia material: Pi con un agente IA; la revisión es del autor.

## Tests que prueban el cierre

- `{test}::test_caso_corregido`
"""

POLICY = {
    "chain_tip_requires_closes": True,
    "closes_pattern": "Closes|Fixes|Resolves",
    "review_budget_lines": 400,
    "size_exception_label": "size:exception",
    "chain_partial_label": "chain:partial",
    "checks": {"docs_evidence": False},
}


def _load(name, path):
    spec = importlib.util.spec_from_loader(
        name, importlib.machinery.SourceFileLoader(name, str(path)))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


LOCAL_MOD = _load("ci_pattern_pr_contract_local", LOCAL)


def _git(repo, *args):
    return subprocess.run(["git", "-C", str(repo), *args], capture_output=True,
                          text=True, check=True).stdout


def _make_repo(root):
    """main con un test; la rama de feature añade un cambio y el test que lo cubre."""
    _git(root, "init", "-q", "-b", "main")
    _git(root, "config", "user.email", "t@example.invalid")
    _git(root, "config", "user.name", "t")
    (root / "app.py").write_text("X = 1\n", encoding="utf-8")
    (root / "tests").mkdir()
    (root / TEST_REL).write_text("def test_caso_corregido():\n    pass\n", encoding="utf-8")
    _git(root, "add", "-A")
    _git(root, "commit", "-q", "-m", "chore: base")
    _git(root, "checkout", "-q", "-b", BRANCH)
    (root / "app.py").write_text("X = 2\n", encoding="utf-8")
    _git(root, "commit", "-qam", "fix: cambia X")


def _body_file(root, text):
    path = root / "body.md"
    path.write_text(text, encoding="utf-8")
    return path


def _policy_file(root):
    path = root / "policy.json"
    path.write_text(json.dumps(POLICY), encoding="utf-8")
    return path


def _run_local(argv):
    # #938: el modo vivo es opt-in (`PR_CONTRACT_USE_LIVE_API=1`). Por
    # defecto el gate lee del payload, igual que en `origin/main`. Esta
    # suite invoca la CLI local en proceso; para que el comportamiento sea
    # el de un usuario sin opt-in, suprimimos la env var en el proceso
    # del test (la corrida humana de la CLI no está afectada porque la
    # suite es lo único que modifica el env).
    out, err = io.StringIO(), io.StringIO()
    saved = os.environ.get("PR_CONTRACT_USE_LIVE_API")
    os.environ.pop("PR_CONTRACT_USE_LIVE_API", None)
    try:
        with redirect_stdout(out), redirect_stderr(err):
            rc = LOCAL_MOD.main(argv)
    finally:
        if saved is not None:
            os.environ["PR_CONTRACT_USE_LIVE_API"] = saved
    return rc, out.getvalue(), err.getvalue()


def _gate_direct(root, body_text):
    """El gate real sobre el mismo evento y los mismos datos, como proceso aparte."""
    meta = {"number": 912, "action": "opened", "full_name": "DysTelefonica/team-skills",
            "default_branch": "main"}
    payload = LOCAL_MOD.build_event(body_text, BRANCH, "main", ["type:feature"], meta)
    files = {
        "event": root / "event.json",
        "policy": _policy_file(root),
        "changed": root / "changed.txt",
        "numstat": root / "numstat.txt",
        "commits": root / "commits.txt",
        "tracked": root / "tracked.txt",
    }
    files["event"].write_text(json.dumps(payload), encoding="utf-8")
    data = LOCAL_MOD.range_data(root, "main", "HEAD")
    files["changed"].write_text(data["changed"], encoding="utf-8")
    files["numstat"].write_text(data["numstat"], encoding="utf-8")
    files["commits"].write_text(data["commits"], encoding="utf-8")
    files["tracked"].write_text(data["tracked"], encoding="utf-8")
    proc = subprocess.run(
        [sys.executable, str(GATE), "--event-file", str(files["event"]),
         "--policy-file", str(files["policy"]), "--repo", str(root),
         "--changed-files", str(files["changed"]),
         "--numstat-file", str(files["numstat"]),
         "--commits-file", str(files["commits"]),
         "--tracked-files", str(files["tracked"])],
        capture_output=True, text=True, check=False,
        env={k: v for k, v in os.environ.items() if k != "PR_CONTRACT_USE_LIVE_API"})
    return proc.returncode, proc.stdout, proc.stderr


class PrContractLocalTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        _make_repo(self.root)
        self.policy = _policy_file(self.root)

    def tearDown(self):
        self._tmp.cleanup()

    def _argv(self, body_text, *extra):
        return ["--body", str(_body_file(self.root, body_text)), "--repo", str(self.root),
                "--policy-file", str(self.policy), "--branch", BRANCH,
                "--base", "main", "--label", "type:feature", "--number", "912",
                "--full-name", "DysTelefonica/team-skills", *extra]

    def test_clean_body_exits_0_with_the_gate_verdict(self):
        body = BODY.format(branch=BRANCH, test=TEST_REL, budget=1)
        rc, out, err = _run_local(self._argv(body))
        self.assertEqual(rc, 0, err)
        self.assertIn("PR-CONTRACT OK", out)

    def test_same_body_gives_the_same_verdict_and_text_as_the_gate(self):
        # Sección HR-53 sin test ni «No aplica»: el gate avisa y el comando debe
        # decir lo mismo (mismo texto, mismo exit code).
        body = BODY.format(branch=BRANCH, test=TEST_REL, budget=1).replace(
            "- `tests/test_fix.py::test_caso_corregido`", "- la prosa no nombra ningún test")
        local_rc, local_out, local_err = _run_local(self._argv(body))
        gate_rc, gate_out, gate_err = _gate_direct(self.root, body)
        self.assertEqual(local_rc, gate_rc)
        self.assertEqual(local_out, gate_out)
        self.assertEqual(local_err, gate_err)
        self.assertEqual(local_rc, 0)
        self.assertIn("AVISO", local_out)

    def test_missing_body_file_exits_2(self):
        rc, _, err = _run_local(["--body", str(self.root / "no-existe.md"),
                                 "--repo", str(self.root), "--policy-file", str(self.policy)])
        self.assertEqual(rc, 2)
        self.assertIn("body", err.lower())

    def test_unresolvable_base_exits_2_and_names_the_ref(self):
        rc, _, err = _run_local(self._argv(BODY.format(branch=BRANCH, test=TEST_REL,
                                                       budget=1), "--base", "rama-inexistente"))
        self.assertEqual(rc, 2)
        self.assertIn("rama-inexistente", err)

    def test_print_payload_dumps_the_event_the_gate_receives(self):
        body = BODY.format(branch=BRANCH, test=TEST_REL, budget=1)
        rc, out, _ = _run_local(self._argv(body, "--print-payload"))
        self.assertEqual(rc, 0)
        event = json.loads(out)
        self.assertEqual(event["pull_request"]["body"], body)
        self.assertEqual(event["pull_request"]["base"]["ref"], "main")
        self.assertEqual(event["pull_request"]["head"]["ref"], BRANCH)
        self.assertEqual(event["repository"]["full_name"], "DysTelefonica/team-skills")

    def test_repo_full_name_is_read_from_the_remote_url_without_network(self):
        cases = {
            "git@github.com:DysTelefonica/team-skills.git": "DysTelefonica/team-skills",
            "https://github.com/DysTelefonica/team-skills": "DysTelefonica/team-skills",
        }
        for url, expected in cases.items():
            self.assertEqual(LOCAL_MOD.repo_full_name_from_url(url), expected)
        source = LOCAL.read_text(encoding="utf-8")
        self.assertNotIn("urllib", source)
        self.assertNotIn('"gh"', source)


class DispatchTests(unittest.TestCase):
    CLI = ASSETS / "bin" / "ci-pattern"

    def test_pr_contract_is_dispatched_by_the_ci_pattern_cli(self):
        proc = subprocess.run(
            [sys.executable, str(self.CLI), "pr", "contract", "--body", "/no/existe.md"],
            capture_output=True, text=True, check=False)
        self.assertEqual(proc.returncode, 2, proc.stderr)
        self.assertIn("body", proc.stderr.lower())


if __name__ == "__main__":
    unittest.main()
