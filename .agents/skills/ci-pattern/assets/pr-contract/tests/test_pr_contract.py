#!/usr/bin/env python3
"""Suite del gate de contrato de PR y commits (#193). stdlib exclusivamente."""

import importlib.machinery
import importlib.util
import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

GATE = Path(__file__).resolve().parents[1] / "check_pr_contract.py"
POLICY = Path(__file__).resolve().parents[1] / "pr-contract.policy.example.json"

POLICY_DATA = json.loads(POLICY.read_text(encoding="utf-8")) if POLICY.is_file() else {
    "required_labels": ["type:bug", "type:feature", "type:docs", "type:chore", "type:refactor"],
    "chain_tip_requires_closes": True,
    "closes_pattern": "Closes|Fixes|Resolves",
    "review_budget_lines": 400,
    "size_exception_field": "size-exception-reason:",
    "conventional_commit_pattern": "^(feat|fix|docs|chore|ci|test|perf|refactor|revert)(\\([^)]+\\))?!?: ",
    "ai_attribution_patterns": ["Co-Authored-By:", "Generated-by:", "Generated with"],
}

def make_repo(tmp, commits=None, base_branch="main"):
    """Crea un repo git con una rama base y opcionalmente una feature."""
    repo = Path(tmp) / "repo"
    repo.mkdir(parents=True)
    env = dict(os.environ)
    run = lambda *a: subprocess.run(a, cwd=repo, capture_output=True, text=True, check=True,
                                    env={**env, "GIT_AUTHOR_NAME": "t", "GIT_AUTHOR_EMAIL": "t@t",
                                         "GIT_COMMITTER_NAME": "t", "GIT_COMMITTER_EMAIL": "t@t"})
    run("git", "init", "-q", "-b", base_branch)
    (repo / "f.txt").write_text("base\n")
    (repo / "tests").mkdir(exist_ok=True)
    (repo / "tests" / "test_gate.py").write_text(
        "def test_gate() -> None:\n    pass\n", encoding="utf-8")

    run("git", "add", "-A")
    run("git", "commit", "-qm", "chore: base")
    base_sha = run("git", "rev-parse", "HEAD").stdout.strip()
    head_sha = base_sha
    if commits:
        run("git", "checkout", "-q", "-b", "feature/x")
        for msg in commits:
            (repo / "f.txt").open("a").write(msg + "\n")
            run("git", "add", "-A")
            run("git", "commit", "-qm", msg)
        head_sha = run("git", "rev-parse", "HEAD").stdout.strip()
    return repo, base_sha, head_sha

def make_event(tmp, body=None, labels=None, base="main", base_sha=None, head_sha=None, action="opened"):
    payload = {
        "action": action,
        "pull_request": {
            "number": 1,
            "base": {"ref": base, "sha": base_sha or "0" * 40},
            "head": {"ref": "feature/x", "sha": head_sha or "0" * 40},
            "labels": [{"name": l} for l in (labels or [])],
            "body": body,
            "title": "test PR",
        },
        "repository": {"default_branch": "main"},
    }
    p = Path(tmp) / "event.json"
    p.write_text(json.dumps(payload), encoding="utf-8")
    return str(p)

def make_policy(tmp, **overrides):
    d = {**POLICY_DATA, **overrides}
    p = Path(tmp) / "policy.json"
    p.write_text(json.dumps(d), encoding="utf-8")
    return str(p)

def run_gate(event_file, policy_file, repo=None):
    cmd = [sys.executable, str(GATE), "--event-file", event_file, "--policy-file", policy_file]
    if repo:
        cmd += ["--repo", str(repo)]
    env = dict(os.environ)
    env.pop("GITHUB_EVENT_PATH", None)
    return subprocess.run(cmd, capture_output=True, text=True, check=False, env=env)

class PrContractGateTests(unittest.TestCase):
    """#193: gate de contrato de PR y commits. Cada regla: positivo, negativo, degenerado."""

    def _run(self, tmp, *, body=None, labels=None, base="main", commits=("feat: x",),
             action="opened", policy_overrides=None):
        repo, base_sha, head_sha = make_repo(tmp, commits=commits)
        ev = make_event(tmp, body=body, labels=labels, base=base,
                        base_sha=base_sha, head_sha=head_sha, action=action)
        pol = make_policy(tmp, **(policy_overrides or {}))
        return run_gate(ev, pol, repo)

    def test_positive_clean_pr(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = self._run(tmp, body="Closes #42\n\n## Tests que prueban el cierre\n\n"
                                "- `tests/test_gate.py` (positivo del gate)\n" + chain_block("main"),
                            labels=["type:feature"])
            self.assertEqual(p.returncode, 0, p.stdout + p.stderr)

    def test_hr6_labels_missing_and_degenerate(self):
        with tempfile.TemporaryDirectory() as tmp1, tempfile.TemporaryDirectory() as tmp2:
            p = self._run(tmp1, labels=[])
            self.assertEqual(p.returncode, 1)
            self.assertIn("etiqueta", p.stdout)
            p = self._run(tmp2, labels=None)
            self.assertEqual(p.returncode, 1)

    def test_hr7_tip_without_closes_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = self._run(tmp, labels=["type:feature"], body="sin closes\n",
                          policy_overrides={"chain_tip_requires_closes": True})
            self.assertEqual(p.returncode, 1)
            self.assertIn("HR-7", p.stdout)

    def test_hr7_closes_on_non_tip_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = self._run(tmp, labels=["type:feature"], base="feature/base",
                          body="Closes #42\n")
            self.assertEqual(p.returncode, 1)
            self.assertIn("HR-7", p.stdout)

    def test_hr8_budget_exceeded_without_exception_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = self._run(tmp, labels=["type:feature"], body="sin excepción\n",
                          commits=["feat: x"] + [f"feat: bulk {i}" for i in range(5)],
                          policy_overrides={"review_budget_lines": 3})
            self.assertEqual(p.returncode, 1)
            self.assertIn("presupuesto", p.stdout)

    def test_hr8_budget_with_exception_passes(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = self._run(tmp, labels=["type:feature"],
                          body="Closes #42\nsize-exception-reason: refactor\n\n"
                               "## Tests que prueban el cierre\n\n- `tests/test_gate.py`\n" + chain_block("main"),
                          commits=["feat: x"] + [f"feat: bulk {i}" for i in range(5)],
                          policy_overrides={"review_budget_lines": 3})
            self.assertEqual(p.returncode, 0, p.stdout)

    def test_hr14_non_conventional_and_ai_attribution_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, base_sha, head_sha = make_repo(tmp, commits=["feat: ok", "arreglo sin conventional"])
            subprocess.run(["git", "-C", repo, "commit", "--amend", "-qm",
                            "arreglo sin conventional\n\nCo-Authored-By: Claude (ai)"],
                           check=True, capture_output=True, text=True,
                           env={**os.environ, "GIT_AUTHOR_NAME": "t", "GIT_AUTHOR_EMAIL": "t@t",
                                "GIT_COMMITTER_NAME": "t", "GIT_COMMITTER_EMAIL": "t@t"})
            head_sha = subprocess.run(["git", "-C", repo, "rev-parse", "HEAD"],
                                      capture_output=True, text=True).stdout.strip()
            ev = make_event(tmp, labels=["type:feature"], base_sha=base_sha, head_sha=head_sha)
            pol = make_policy(tmp)
            proc = run_gate(ev, pol, repo)
            self.assertEqual(proc.returncode, 1)
            self.assertIn("convencional", proc.stdout)
            self.assertIn("HR-14", proc.stdout)

    def test_hr14_merge_commit_is_not_judged_as_conventional(self):
        """Un commit de merge no tiene asunto conventional: la regla mide los
        commits de trabajo, y la disciplina de cadenas obliga a sincronizar
        con main mergeando (#333). Antes de este arreglo, cualquier PR que
        hubiera sincronizado con main quedaba marcado por HR-14."""
        with tempfile.TemporaryDirectory() as tmp:
            repo, base_sha, _head = make_repo(tmp, commits=["feat: x"])
            env = {**os.environ, "GIT_AUTHOR_NAME": "t", "GIT_AUTHOR_EMAIL": "t@t",
                   "GIT_COMMITTER_NAME": "t", "GIT_COMMITTER_EMAIL": "t@t"}

            def git(*args):
                return subprocess.run(["git", "-C", str(repo), *args], capture_output=True,
                                      text=True, check=True, env=env)

            git("checkout", "-q", "-b", "side", base_sha)
            (repo / "side.txt").write_text("x\n", encoding="utf-8")
            git("add", "-A")
            git("commit", "-qm", "docs: rama lateral")
            git("checkout", "-q", "feature/x")
            git("merge", "-q", "--no-ff", "-m", "Merge branch 'side' into feature/x", "side")
            head_sha = git("rev-parse", "HEAD").stdout.strip()
            ev = make_event(tmp, body="Closes #42\n\n## Tests que prueban el cierre\n\n"
                            "- `tests/test_gate.py`\n" + chain_block("main"),
                            labels=["type:feature"], base_sha=base_sha, head_sha=head_sha)
            proc = run_gate(ev, make_policy(tmp), repo)
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
            self.assertNotIn("HR-14", proc.stdout)

    def test_chain_partial_scenarios(self):
        """chain:partial: Refs obligatorio, Closes prohibido, ausencia de ambos falla."""
        for body, rc, needle in [
            ("Refs #161\n", 0, None),
            ("Closes #161\n", 1, "Closes"),
            ("descripción sin refs\n", 1, "Refs"),
        ]:
            with self.subTest(body=body):
                with tempfile.TemporaryDirectory() as tmp:
                    p = self._run(tmp, labels=["type:feature", "chain:partial"],
                                  body=body + chain_block("main"),
                                  policy_overrides={"chain_partial_label": "chain:partial"})
                    self.assertEqual(p.returncode, rc, p.stdout + p.stderr)
                    if needle:
                        self.assertIn(needle, p.stdout)

    def test_hr31_edited_action_still_validates(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = self._run(tmp, labels=[], action="edited")
            self.assertEqual(p.returncode, 1)
            self.assertIn("etiqueta", p.stdout)

    def test_empty_event_and_missing_file_exit_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            ev = Path(tmp) / "event.json"
            ev.write_text("{}", encoding="utf-8")
            pol = make_policy(tmp)
            p = run_gate(str(ev), pol)
            self.assertEqual(p.returncode, 2)
            p = run_gate(str(Path(tmp) / "no-existe.json"), pol)
            self.assertEqual(p.returncode, 2)


CLOSURE_HEADING = "Tests que prueban el cierre"

_spec = importlib.util.spec_from_loader(
    "pr_contract_gate",
    importlib.machinery.SourceFileLoader("pr_contract_gate", str(GATE)))
pr_contract = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(pr_contract)
section_body = pr_contract.section_body


def chain_block(base, position="1/1", budget="280/400"):
    """Bloque Chain Context conforme (#333), para los fixtures que deben salir 0."""
    return (f"\n## Chain Context\n\n- chain: 333\n- position: {position}\n"
            f"- base: {base}\n- depends-on: #335\n- follow-up: —\n"
            "- starts-at: gate sin disciplina de cadenas\n"
            "- ends-with: Chain Context verificado\n"
            f"- review-budget: {budget}\n\n📍 este PR ({position})\n")


def make_closure_repo(tmp):
    """Repo con un test real VERSIONADO: el cierre tiene que poder nombrarlo
    y la ruta tiene que estar en git (un fichero sin versionar no prueba)."""
    repo, _base, _head = make_repo(tmp)
    (repo / "tests").mkdir(exist_ok=True)
    (repo / "tests" / "test_x.py").write_text(
        "def test_algo():\n    assert True\n", encoding="utf-8")
    subprocess.run(["git", "-C", str(repo), "add", "-A"],
                   check=True, capture_output=True, text=True)
    # identidad explícita: en CI no hay git global configurado y el commit
    # fallaría (el meta-gate ve el fichero de test como roto).
    subprocess.run(["git", "-C", str(repo), "-c", "user.name=fixture",
                    "-c", "user.email=fixture@invalid",
                    "commit", "-qm", "test: fixture"],
                   check=True, capture_output=True, text=True)
    sha = subprocess.run(["git", "-C", str(repo), "rev-parse", "HEAD"],
                         check=True, capture_output=True, text=True).stdout.strip()
    return repo, sha, sha


class ClosureTestsGateTests(unittest.TestCase):
    """#334 (HR-53): un PR que cierra una issue nombra el test que la prueba,
    en formato ESTRUCTURADO (`ruta` o `ruta::caso`), versionado en git y con
    el caso definido en su fichero. La prosa no cierra nada."""

    def run_case(self, body, labels=("type:chore",), base="main", extra_file=None):
        tmp = tempfile.mkdtemp()
        repo, base_sha, head_sha = make_closure_repo(tmp)
        if extra_file:
            (repo / extra_file).write_text(
                "def test_sin_versionar():\n    pass\n", encoding="utf-8")
        event = make_event(tmp, body=body + chain_block(base), labels=list(labels),
                           base=base, base_sha=base_sha, head_sha=head_sha)
        return run_gate(event, make_policy(tmp), repo=repo)

    def test_section_body_returns_the_section_text(self):
        body = ("## Otro\n\nalgo\n\n## Tests que prueban el cierre\n\n"
                "`tests/test_x.py` -> `tests/test_x.py::test_algo`\n\n"
                "## Siguiente\n\nnada\n")
        found = section_body(body, CLOSURE_HEADING)
        self.assertIn("test_algo", found)
        self.assertNotIn("nada", found)

    def test_section_body_absent_is_empty(self):
        self.assertEqual(section_body("## Solo esto\n\ntexto\n", CLOSURE_HEADING), "")

    def test_tip_without_the_closure_section_fails(self):
        proc = self.run_case("Closes #12\n")
        self.assertEqual(proc.returncode, 1)
        self.assertIn("HR-53", proc.stdout + proc.stderr)

    def test_closure_section_without_an_entry_fails(self):
        proc = self.run_case(f"Closes #12\n\n## {CLOSURE_HEADING}\n\ndebería estar arreglado\n")
        self.assertEqual(proc.returncode, 1)
        self.assertIn("formato", proc.stdout)

    def test_prose_that_looks_like_a_test_does_not_pass(self):
        # «test assert» produce el token `assert`, presente en cualquier
        # fichero de test: con el formato libre, esto pasaba en verde.
        proc = self.run_case(f"Closes #12\n\n## {CLOSURE_HEADING}\n\n"
                             "test assert cubre el caso\n")
        self.assertEqual(proc.returncode, 1)
        self.assertIn("HR-53", proc.stdout)

    def test_closure_naming_an_existing_file_passes(self):
        proc = self.run_case(f"Closes #12\n\n## {CLOSURE_HEADING}\n\n- `tests/test_x.py`\n")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_closure_naming_a_defined_case_passes(self):
        proc = self.run_case(f"Closes #12\n\n## {CLOSURE_HEADING}\n\n"
                             "- `tests/test_x.py::test_algo`\n")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_closure_case_not_defined_in_the_file_fails(self):
        proc = self.run_case(f"Closes #12\n\n## {CLOSURE_HEADING}\n\n"
                             "- `tests/test_x.py::test_que_no_esta`\n")
        self.assertEqual(proc.returncode, 1)
        self.assertIn("no está definido", proc.stdout)

    def test_closure_path_not_tracked_in_git_fails(self):
        proc = self.run_case(f"Closes #12\n\n## {CLOSURE_HEADING}\n\n"
                             "- `tests/test_sin_versionar.py`\n",
                             extra_file="tests/test_sin_versionar.py")
        self.assertEqual(proc.returncode, 1)
        self.assertIn("no está versionada", proc.stdout)

    def test_closure_naming_a_path_that_does_not_exist_fails(self):
        proc = self.run_case(f"Closes #12\n\n## {CLOSURE_HEADING}\n\n"
                             "- `tests/test_que_no_existe.py`\n")
        self.assertEqual(proc.returncode, 1)
        self.assertIn("test_que_no_existe.py", proc.stdout)

    def test_pr_that_does_not_close_is_not_asked_for_a_test(self):
        proc = self.run_case("Refs #12\n", labels=("type:chore", "chain:partial"),
                             base="feature/x")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)



def chain_ok(base="feature/x"):
    return ("## Chain Context\n\n- chain: 333\n- position: 1/1\n"
            f"- base: {base}\n- depends-on: #335\n- follow-up: —\n"
            "- starts-at: gate sin disciplina de cadenas\n"
            "- ends-with: Chain Context verificado\n- review-budget: 280/400\n\n"
            "📍 este PR (1/1)\n")


class ChainContextGateTests(unittest.TestCase):
    """#333: la sección Chain Context es obligatoria con `chain:partial` y en la
    punta de una cadena; sus campos son datos que el gate verifica."""

    def run_case(self, body, labels=("type:chore", "chain:partial"), base="feature/x"):
        tmp = tempfile.mkdtemp()
        repo, base_sha, head_sha = make_repo(tmp)
        event = make_event(tmp, body=body, labels=list(labels), base=base,
                           base_sha=base_sha, head_sha=head_sha)
        return run_gate(event, make_policy(tmp), repo=repo)

    def test_conforming_chain_context_passes(self):
        proc = self.run_case("Refs #333\n\n" + chain_ok())
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_missing_chain_context_fails_naming_the_section(self):
        proc = self.run_case("Refs #333\n")
        self.assertEqual(proc.returncode, 1)
        self.assertIn("Chain Context", proc.stdout)

    def test_missing_field_fails_naming_it(self):
        proc = self.run_case("Refs #333\n\n" + chain_ok().replace("- follow-up: —\n", ""))
        self.assertEqual(proc.returncode, 1)
        self.assertIn("follow-up", proc.stdout)

    def test_standalone_pr_declares_chain_context_1_1(self):
        # Auditoría de #352: el gate no distingue la punta de una cadena de un PR
        # suelto, así que la regla lo declara explícito: todo PR contra la rama por
        # defecto lleva Chain Context, y un PR suelto declara 1/1 con
        # `depends-on`/`follow-up` en `none`.
        proc = self.run_case(
            "Closes #333\n\n## " + CLOSURE_HEADING + "\n\n- `tests/test_gate.py`\n"
            "## Chain Context\n\n- chain: ninguna\n- position: 1/1\n- base: main\n"
            "- depends-on: none\n- follow-up: none\n"
            "- starts-at: main sin PR previo\n- ends-with: PR suelto conforme\n"
            "- review-budget: 12/400\n\n📍 este PR\n",
            labels=("type:chore",), base="main")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_standalone_pr_without_chain_context_fails(self):
        proc = self.run_case("Closes #333\n\n## " + CLOSURE_HEADING
                             + "\n\n- `tests/test_gate.py`\n",
                             labels=("type:chore",), base="main")
        self.assertEqual(proc.returncode, 1)
        self.assertIn("Chain Context", proc.stdout)

    def test_base_mismatch_fails_naming_the_field(self):
        proc = self.run_case("Refs #333\n\n" + chain_ok("otra-rama"))
        self.assertEqual(proc.returncode, 1)
        self.assertIn("base", proc.stdout)

    def test_no_marker_fails(self):
        proc = self.run_case("Refs #333\n\n"
                             + chain_ok().replace("📍 este PR (1/1)", "el PR actual (1/1)"))
        self.assertEqual(proc.returncode, 1)
        self.assertIn("📍", proc.stdout)

    def test_two_markers_fail(self):
        proc = self.run_case("Refs #333\n\n" + chain_ok() + "\n📍 y otro 📍\n")
        self.assertEqual(proc.returncode, 1)
        self.assertIn("📍", proc.stdout)

    def test_malformed_position_and_budget_fail(self):
        proc = self.run_case("Refs #333\n\n" + chain_ok().replace("position: 1/1", "position: uno")
                             .replace("review-budget: 280/400", "review-budget: mucho"))
        self.assertEqual(proc.returncode, 1)
        self.assertIn("position", proc.stdout)
        self.assertIn("review-budget", proc.stdout)


class IssueLinkGateTests(unittest.TestCase):
    """#333: el enlace de la issue no admite mezclas ni referencias a otro repo."""

    def run_case(self, body, labels=("type:chore",)):
        tmp = tempfile.mkdtemp()
        repo, base_sha, head_sha = make_repo(tmp)
        event = make_event(tmp, body=body, labels=list(labels), base="main",
                           base_sha=base_sha, head_sha=head_sha)
        return run_gate(event, make_policy(tmp), repo=repo)

    def test_closes_and_refs_for_the_same_issue_fails(self):
        proc = self.run_case("Closes #333\n\n" + chain_ok("main") + "\nRefs #333\n")
        self.assertEqual(proc.returncode, 1)
        self.assertIn("#333", proc.stdout)

    def test_foreign_repository_link_fails(self):
        proc = self.run_case("Closes #333\n\nRefs otro-org/otro-repo#12\n\n" + chain_ok("main"))
        self.assertEqual(proc.returncode, 1)
        self.assertIn("otro-org/otro-repo#12", proc.stdout)

    def test_prose_mention_of_a_foreign_issue_is_not_a_link(self):
        # Auditoría de #352: el gate marcaba CUALQUIER mención `owner/repo#N` del
        # cuerpo, y citar una issue de otro repositorio como contexto —el piloto
        # de Cadete lo hace a diario— no es enlazarla.
        proc = self.run_case("Refs #333\n\nContexto: se cita DysTelefonica/cadete#318 "
                             "como antecedente del piloto.\n\n" + chain_ok("main"),
                             labels=("type:chore", "chain:partial"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_own_closing_link_passes(self):
        proc = self.run_case("Closes #333\n\n## " + CLOSURE_HEADING
                             + "\n\n- `tests/test_gate.py`\n" + chain_ok("main"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

if __name__ == "__main__":
    unittest.main(verbosity=2)
