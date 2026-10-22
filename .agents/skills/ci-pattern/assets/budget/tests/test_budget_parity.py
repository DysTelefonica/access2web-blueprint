#!/usr/bin/env python3
"""Suite de paridad del presupuesto de revisión (#931).

``pr-size`` y ``pr-contract`` (A1) deben contar el mismo diff con las mismas
exclusiones. La fuente única es ``assets/budget/vendored_budget.py``. Un
renombrado desde una ruta authored hacia una vendorizada sigue contando,
también cuando el origen está en otro commit del mismo PR.

stdlib exclusivamente, ``unittest``, repos ``git`` reales en ``tempfile``.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ASSETS = Path(__file__).resolve().parents[2]
PR_SIZE = ASSETS / "pr-size" / "check_pr_size.py"
PR_CONTRACT = ASSETS / "pr-contract" / "check_pr_contract.py"
POLICY = ASSETS / "pr-contract" / "pr-contract.policy.example.json"
BUDGET = ASSETS / "budget" / "vendored_budget.py"
AI_OK = "## Asistencia de IA\n\nninguna\n"
TEST_PATH = "personal/ardelperal/ci-pattern/assets/budget/tests/test_vendored_budget.py"
CHAIN_BLOCK = (
    "\n## Chain Context\n\n- chain: 931\n- position: 1/1\n- base: main\n"
    "- depends-on: #929\n- follow-up: \u2014\n"
    "- starts-at: gate sin disciplina de cadenas\n"
    "- ends-with: Chain Context verificado\n"
    "- review-budget: 280/400\n\n\U0001f4cd este PR (1/1)\n"
)
CONTRACT_BODY = (
    "Closes #931\n\n## Tests que prueban el cierre\n\n"
    f"- `{TEST_PATH}`\n"
    + CHAIN_BLOCK + "\n" + AI_OK
)


def _clean_env() -> dict[str, str]:
    return {k: v for k, v in os.environ.items() if not k.startswith("GIT_")}


def _git(repo: Path, *args: str) -> str:
    out = subprocess.run(
        ["git", "-c", "user.name=t", "-c", "user.email=t@example.invalid", *args],
        cwd=repo, capture_output=True, text=True, check=True, env=_clean_env(),
    )
    return out.stdout


def _write(repo: Path, rel: str, lines: int) -> None:
    path = repo / rel
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("".join(f"line {i}\n" for i in range(lines)), encoding="utf-8")


def _load_budget():
    import importlib.util

    spec = importlib.util.spec_from_file_location("vendored_budget", BUDGET)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class ParityFixture(unittest.TestCase):
    """Repo con main, un commit de autoría y un commit de propagación."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.repo = Path(self._tmp.name)
        _git(self.repo, "init", "-q", "-b", "main")
        _write(self.repo, "app/a.py", 100)
        _git(self.repo, "add", "-A")
        _git(self.repo, "commit", "-q", "-m", "chore: base")
        _git(self.repo, "checkout", "-q", "-b", "feat/x")
        # Autoría a mano: 120 líneas añadidas en app/a.py y un fichero nuevo de 300.
        with (self.repo / "app/a.py").open("a", encoding="utf-8") as handle:
            handle.write("".join(f"extra {i}\n" for i in range(120)))
        _write(self.repo, "src/x.py", 300)
        _git(self.repo, "add", "-A")
        _git(self.repo, "commit", "-q", "-m", "feat: autoria")
        # Propagación: el fichero de src/x.py pasa a .agents/skills (cruce entre
        # commits) y se añade 1000 líneas de contenido vendorizado.
        (self.repo / ".agents/skills").mkdir(parents=True, exist_ok=True)
        _git(self.repo, "mv", "src/x.py", ".agents/skills/x.py")
        _write(self.repo, ".agents/skills/prop.md", 1000)
        _git(self.repo, "add", "-A")
        _git(self.repo, "commit", "-q", "-m", "chore(fleet): propagar")

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def numstat(self) -> str:
        return _git(self.repo, "diff", "--numstat", "--find-renames=50%",
                    "--ignore-cr-at-eol", "main...HEAD")

    def origins_text(self) -> str:
        budget = _load_budget()
        return budget.origins_for("main", lambda *args: _git(self.repo, *args))


class TestSingleSource(unittest.TestCase):
    def test_pr_size_takes_the_roots_from_the_shared_module(self) -> None:
        budget = _load_budget()
        import importlib.util

        spec = importlib.util.spec_from_file_location("check_pr_size", PR_SIZE)
        assert spec is not None and spec.loader is not None
        gate = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(gate)
        self.assertEqual(gate.VENDORED_ROOTS, budget.VENDORED_ROOTS)
        self.assertEqual(gate.counts_to_budget.__module__, "vendored_budget")
        self.assertEqual(gate.original_path.__module__, "vendored_budget")

    def test_pr_contract_does_not_declare_its_own_roots(self) -> None:
        source = PR_CONTRACT.read_text(encoding="utf-8")
        self.assertNotIn(".agents/skills/", source)


class TestParity(ParityFixture):
    """RED: pr-size y el cálculo compartido dan la misma cifra para el mismo PR."""

    def test_shared_figure_matches_pr_size(self) -> None:
        budget = _load_budget()
        figure = budget.budget_lines(self.numstat(), budget.parse_origins(self.origins_text()))
        proc = subprocess.run(
            [sys.executable, str(PR_SIZE), "--base-ref", "main", "--json"],
            cwd=self.repo, capture_output=True, text=True, env=_clean_env(),
        )
        envelope = json.loads(proc.stdout)
        self.assertEqual(figure, envelope["indicators"]["changed_lines"])
        # 120 de autoría en app/a.py + 300 de src/x.py que cruzó a vendorizado.
        self.assertEqual(figure, 420)

    def test_pr_contract_accepts_the_origins_and_passes_a_propagation_only_pr(self) -> None:
        """RED: sin --origins-file no hay forma de excluir lo vendorizado."""
        propagation = self.repo / "prop-numstat.txt"
        propagation.write_text("1000\t0\t.agents/skills/prop.md\n", encoding="utf-8")
        origins = self.repo / "origins.txt"
        origins.write_text(".agents/skills/prop.md\t-\n", encoding="utf-8")
        proc = self._contract(numstat=propagation, origins=origins)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_pr_contract_still_counts_the_authored_cross_commit_rename(self) -> None:
        """RED: el origen authored en otro commit cuenta aunque el destino sea vendorizado."""
        numstat = self.repo / "x-numstat.txt"
        numstat.write_text("300\t0\t.agents/skills/x.py\n", encoding="utf-8")
        origins = self.repo / "origins-x.txt"
        origins.write_text(".agents/skills/x.py\tsrc/x.py\n", encoding="utf-8")
        policy = json.loads(POLICY.read_text(encoding="utf-8"))
        policy["review_budget_lines"] = 200
        policy_path = self.repo / "policy.json"
        policy_path.write_text(json.dumps(policy), encoding="utf-8")
        proc = self._contract(numstat=numstat, origins=origins, policy=policy_path)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("A1", proc.stdout + proc.stderr)

    def test_without_origins_the_budget_is_unchanged(self) -> None:
        """Pin: sin --origins-file pr-contract cuenta todo, como hoy."""
        numstat = self.repo / "all-numstat.txt"
        numstat.write_text("1000\t0\t.agents/skills/prop.md\n", encoding="utf-8")
        proc = self._contract(numstat=numstat, origins=None)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)

    def _contract(self, numstat: Path, origins: Path | None,
                  policy: Path = POLICY) -> subprocess.CompletedProcess:
        commits = self.repo / "commits.txt"
        commits.write_text("p\nfeat: x\n---COMMIT-END---\n", encoding="utf-8")
        changed = self.repo / "changed.txt"
        changed.write_text("", encoding="utf-8")
        # Simula `git ls-files`: el test que cierra el PR está versionado en el repo real.
        tracked = self.repo / "tracked.txt"
        tracked.write_text(TEST_PATH + "\n", encoding="utf-8")
        event = self.repo / "event.json"
        event.write_text(json.dumps({
            "repository": {"full_name": "o/r", "default_branch": "main"},
            "pull_request": {"number": 1, "title": "feat: x", "body": CONTRACT_BODY,
                             "base": {"ref": "main"}, "head": {"ref": "feat/x"},
                             "labels": [{"name": "type:feature"}]},
        }), encoding="utf-8")
        args = [sys.executable, str(PR_CONTRACT), "--event-file", str(event),
                "--policy-file", str(policy), "--repo", str(self.repo),
                "--changed-files", str(changed), "--numstat-file", str(numstat),
                "--commits-file", str(commits), "--tracked-files", str(tracked)]
        if origins is not None:
            args += ["--origins-file", str(origins)]
        return subprocess.run(args, capture_output=True, text=True, env=_clean_env())


if __name__ == "__main__":
    unittest.main()
