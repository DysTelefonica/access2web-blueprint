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
# #876: cuerpo de PR con la declaración de asistencia de IA, para los casos que deben pasar.
AI_OK = "## Asistencia de IA\n\nninguna\n"

POLICY_DATA = json.loads(POLICY.read_text(encoding="utf-8")) if POLICY.is_file() else {
    "required_labels": ["type:bug", "type:feature", "type:docs", "type:chore", "type:refactor"],
    "chain_tip_requires_closes": True,
    "closes_pattern": "Closes|Fixes|Resolves",
    "review_budget_lines": 400,
    "size_exception_label": "size:exception",
    "conventional_commit_pattern": "^(feat|fix|docs|chore|ci|test|perf|refactor|revert)(\\([^)]+\\))?!?: ",
    "ai_attribution_patterns": ["Co-Authored-By:", "Generated-by:", "Generated with"],
}

def make_repo(tmp, commits=None, base_branch="main"):
    """Crea un repo git con una rama base y opcionalmente una feature."""
    repo = Path(tmp) / "repo"
    repo.mkdir(parents=True)
    env = dict(os.environ)
    def run(*a):
        return subprocess.run(a, cwd=repo, capture_output=True, text=True, check=True,
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
            "labels": [{"name": item} for item in (labels or [])],
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

def git_text(repo, *args):
    """Salida de git en el repo de pruebas ('' si falla): hace de workflow del consumer."""
    done = subprocess.run(["git", "-C", str(repo), *args], capture_output=True, text=True, check=False)
    return done.stdout if done.returncode == 0 else ""


def write_range_data(event_file, repo, dest):
    """Genera los cuatro ficheros de datos del rango, como haría el workflow."""
    pr = json.loads(Path(event_file).read_text(encoding="utf-8")).get("pull_request", {})
    base = pr.get("base", {}).get("sha", "")
    head = pr.get("head", {}).get("sha", "")
    rng = f"{base}..{head}" if base and head and base != head else "HEAD..HEAD"
    # Diff de la PR: merge-base (...), no el árbol de base.sha, que se mueve al reapuntar la PR (#1013).
    diff_rng = rng.replace("..", "...")
    files = {
        "changed.txt": git_text(repo, "diff", "--name-only", diff_rng),
        "numstat.txt": git_text(repo, "diff", "--numstat", diff_rng),
        "commits.txt": git_text(repo, "log", "--format=%P%n%B%n---COMMIT-END---", rng),
        "tracked.txt": git_text(repo, "ls-files"),
    }
    for name, text in files.items():
        (Path(dest) / name).write_text(text, encoding="utf-8")
    return ["--changed-files", str(Path(dest) / "changed.txt"),
            "--numstat-file", str(Path(dest) / "numstat.txt"),
            "--commits-file", str(Path(dest) / "commits.txt"),
            "--tracked-files", str(Path(dest) / "tracked.txt")]


def run_gate(event_file, policy_file, repo=None, data_flags=None):
    cmd = [sys.executable, str(GATE), "--event-file", event_file, "--policy-file", policy_file]
    if repo:
        cmd += ["--repo", str(repo)]
    env = dict(os.environ)
    env.pop("GITHUB_EVENT_PATH", None)
    # #938: el modo vivo es opt-in (`PR_CONTRACT_USE_LIVE_API=1`); por defecto
    # el gate lee del payload, igual que en `origin/main`. Esta suite NO setea
    # el opt-in: prueba el modo heredado de lectura del payload, que es el
    # contrato que el modo local, el propagador y la suite de parity esperan.
    env.pop("PR_CONTRACT_USE_LIVE_API", None)
    with tempfile.TemporaryDirectory() as dest:
        if data_flags is None and repo:
            data_flags = write_range_data(event_file, repo, dest)
        return subprocess.run(cmd + (data_flags or []), capture_output=True, text=True,
                              check=False, env=env)

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
                                "- `tests/test_gate.py` (positivo del gate)\n" + chain_block("main") + "\n" + AI_OK,
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

    def test_a1_budget_exceeded_without_label_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = self._run(tmp, labels=["type:feature"], body="sin excepción\n",
                          commits=["feat: x"] + [f"feat: bulk {i}" for i in range(5)],
                          policy_overrides={"review_budget_lines": 3})
            self.assertEqual(p.returncode, 1)
            self.assertIn("presupuesto", p.stdout)
            self.assertIn("size:exception", p.stdout)
            self.assertNotIn("size-exception-reason", p.stdout)

    def test_a1_budget_with_label_passes(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = self._run(tmp, labels=["type:feature", "size:exception"],
                          body="Closes #42\n\n## Excepción de tamaño\n\n"
                               "1. parte en PR previa\n2. sí, por capa\n3. no\n4. ninguno\n\n"
                               "## Tests que prueban el cierre\n\n- `tests/test_gate.py`\n"
                               + chain_block("main") + "\n" + AI_OK,
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
                            "- `tests/test_gate.py`\n" + chain_block("main") + "\n" + AI_OK,
                            labels=["type:feature"], base_sha=base_sha, head_sha=head_sha)
            proc = run_gate(ev, make_policy(tmp), repo)
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
            self.assertNotIn("HR-14", proc.stdout)

    # #944: el gate aplicaba cada patrón de atribución de IA con
    # `re.search(pat, msg, re.IGNORECASE)`, sin límite de palabra ni ancla, así
    # que una palabra que CONTIENE el patrón disparaba el hallazgo:
    # «regenerated with» es la redacción natural de un commit que describe una
    # regeneración, y el guion hace lo mismo con «No-Co-Authored-By:» o
    # «hand-generated-by:». Estas cadenas fijan los DOS lados de la regla.
    HR14_NOT_ATTRIBUTION = [
        ("regenerado", "docs: doc regenerado\n\nthe operating doc is regenerated "
                       "with `ci-pattern adoption generate-doc`"),
        ("negado", "docs: matiz del trailer\n\n"
                   "No-Co-Authored-By: no aplica a este commit"),
        ("prefijo con guion", "docs: matiz del trailer\n\n"
                              "hand-generated-by: el generador de este repo es propio"),
        ("prefijo con guion, otro nombre", "docs: matiz del trailer\n\n"
                                           "un-generated-by: variante con prefijo"),
    ]
    HR14_ATTRIBUTION = [
        ("co-authored-by", "chore: trailer\n\nCo-Authored-By: Claude (ai)"),
        ("generated with", "chore: trailer\n\nGenerated with [Claude Code]"),
        ("generated-by", "chore: trailer\n\nGenerated-by: claude-code"),
        ("co-authored-by tras un emoji", "chore: trailer\n\n🤖 Co-Authored-By: Claude (ai)"),
        ("generated with tras un emoji", "chore: trailer\n\n🤖 Generated with [Claude Code]"),
    ]

    def _hr14(self, tmp, message, **policy_overrides):
        """PR por lo demás limpio (etiqueta, Closes, Chain Context y asistencia de
        IA) con UN commit: el único hallazgo posible es el de HR-14."""
        body = ("Closes #42\n\n## Tests que prueban el cierre\n\n"
                "- `tests/test_gate.py` (límite de palabra de HR-14)\n"
                + chain_block("main") + "\n" + AI_OK)
        return self._run(tmp, body=body, labels=["type:feature"], commits=[message],
                         policy_overrides=policy_overrides or None)

    def test_hr14_word_containing_the_pattern_is_not_attribution(self):
        for caso, message in self.HR14_NOT_ATTRIBUTION:
            with self.subTest(caso=caso), tempfile.TemporaryDirectory() as tmp:
                p = self._hr14(tmp, message)
                self.assertEqual(p.returncode, 0, p.stdout + p.stderr)
                self.assertNotIn("atribución de IA", p.stdout)

    def test_hr14_real_trailer_is_still_attribution(self):
        """El arreglo no debe relajar la detección: el trailer real, solo o tras un
        emoji que lo precede, sigue siendo hallazgo."""
        for caso, message in self.HR14_ATTRIBUTION:
            with self.subTest(caso=caso), tempfile.TemporaryDirectory() as tmp:
                p = self._hr14(tmp, message)
                self.assertEqual(p.returncode, 1, p.stdout + p.stderr)
                self.assertIn("atribución de IA", p.stdout)

    def test_hr14_consumer_declared_patterns_are_normalized(self):
        """La normalización va en el PATRÓN, no en la lista de cada política: un
        consumer que declara `ai_attribution_patterns` sin ancla no reintroduce
        el falso positivo, y su patrón sigue cazando el trailer real."""
        consumer = ["Co-Authored-By:", "Generated-by:", "Generated with", "X-Authored-By:"]
        with tempfile.TemporaryDirectory() as tmp:
            p = self._hr14(tmp, "docs: matiz\n\n"
                               "pre-X-Authored-By: este commit no lo firma un agente\n",
                           ai_attribution_patterns=consumer)
            self.assertEqual(p.returncode, 0, p.stdout + p.stderr)
            self.assertNotIn("atribución de IA", p.stdout)
        with tempfile.TemporaryDirectory() as tmp:
            p = self._hr14(tmp, "docs: matiz\n\nX-Authored-By: agente\n",
                           ai_attribution_patterns=consumer)
            self.assertEqual(p.returncode, 1, p.stdout + p.stderr)
            self.assertIn("atribución de IA", p.stdout)

    def test_chain_partial_scenarios(self):
        """chain:partial: Refs obligatorio, Closes prohibido, ausencia de ambos falla."""
        for body, rc, needle in [
            ("Refs #161\n\n" + AI_OK, 0, None),
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



def _commit_all(repo, message):
    """Commit de fixture con identidad explícita (CI no tiene git global)."""
    subprocess.run(["git", "-C", str(repo), "add", "-A"], check=True, capture_output=True, text=True)
    subprocess.run(["git", "-C", str(repo), "-c", "user.name=fixture", "-c", "user.email=fixture@invalid",
                    "commit", "-qm", message], check=True, capture_output=True, text=True)

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
        proc = self.run_case(f"Closes #12\n\n## {CLOSURE_HEADING}\n\n- `tests/test_x.py`\n" + "\n" + AI_OK)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_closure_naming_a_defined_case_passes(self):
        proc = self.run_case(f"Closes #12\n\n## {CLOSURE_HEADING}\n\n"
                             "- `tests/test_x.py::test_algo`\n\n" + AI_OK)
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
        proc = self.run_case("Refs #12\n\n" + AI_OK, labels=("type:chore", "chain:partial"),
                             base="feature/x")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)


class ClosureNotApplicableGateTests(unittest.TestCase):
    """#911: un PR que cierra una issue SIN test declara «No aplica» con un
    encabezado estable y nombra la evidencia del cierre: una ruta versionada
    en git. La evidencia no tiene que ser un fichero de test. Sin exención por
    etiqueta: `type:docs` no exonera nada.
    #918: «No aplica» solo vale si TODOS los ficheros cambiados son no
    ejecutables (documentación por defecto); un .py o un workflow exige test."""

    MARKER = "No aplica"

    def run_case(self, body, labels=("type:chore",), base="main", evidence=None,
                 commit_evidence=True, changes=("README.md",), policy=None,
                 no_changed=False):
        tmp = tempfile.mkdtemp()
        repo, base_sha, _head = make_closure_repo(tmp)
        for rel in changes:
            path = Path(repo) / rel
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text("# cambio\n", encoding="utf-8")
        if changes:
            _commit_all(repo, "docs: cambios del PR")
        if evidence:
            path = Path(repo) / evidence
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text("# evidencia del cierre\n", encoding="utf-8")
            if commit_evidence:
                _commit_all(repo, "docs: evidencia")
        head_sha = subprocess.run(["git", "-C", str(repo), "rev-parse", "HEAD"],
                                  check=True, capture_output=True, text=True).stdout.strip()
        event = make_event(tmp, body=body + chain_block(base), labels=list(labels),
                           base=base, base_sha=base_sha, head_sha=head_sha)
        flags = None
        if no_changed:
            # sin --changed-files, docs_evidence (si está activo) saldría con 2 antes
            # de llegar a la comprobación de «No aplica»: se apaga para probar esa rama.
            policy = {"checks": {"docs_evidence": False}, **(policy or {})}
            flags = write_range_data(event, repo, tmp)[2:]
        policy_file = make_policy(tmp, **(policy or {}))
        return run_gate(event, policy_file, repo=repo, data_flags=flags)

    def test_not_applicable_with_a_versioned_evidence_path_passes(self):
        # La evidencia es un documento: no es un fichero de test y el gate la
        # acepta igual (no es una exención: es una forma con su propia prueba).
        proc = self.run_case(
            f"Closes #12\n\n## {CLOSURE_HEADING}\n\n### {self.MARKER}\n\n"
            "- `docs/decisions/cierre.md`\n\n" + AI_OK,
            evidence="docs/decisions/cierre.md")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_not_applicable_without_evidence_fails(self):
        proc = self.run_case(
            f"Closes #12\n\n## {CLOSURE_HEADING}\n\n### {self.MARKER}\n\n"
            "el cierre lo cubre la documentación del repo\n")
        self.assertEqual(proc.returncode, 1)
        self.assertIn("evidencia del cierre", proc.stdout)

    def test_not_applicable_with_unversioned_evidence_fails(self):
        proc = self.run_case(
            f"Closes #12\n\n## {CLOSURE_HEADING}\n\n### {self.MARKER}\n\n"
            "- `docs/decisions/cierre.md`\n",
            evidence="docs/decisions/cierre.md", commit_evidence=False)
        self.assertEqual(proc.returncode, 1)
        self.assertIn("no está versionada", proc.stdout)

    def test_not_applicable_with_an_undefined_case_fails(self):
        proc = self.run_case(
            f"Closes #12\n\n## {CLOSURE_HEADING}\n\n### {self.MARKER}\n\n"
            "- `docs/decisions/cierre.md::test_que_no_esta`\n",
            evidence="docs/decisions/cierre.md")
        self.assertEqual(proc.returncode, 1)
        self.assertIn("no está definido", proc.stdout)

    def test_the_marker_outside_the_section_fails_naming_the_section(self):
        proc = self.run_case(
            f"Closes #12\n\n## {self.MARKER}\n\n- `docs/decisions/cierre.md`\n\n"
            f"## {CLOSURE_HEADING}\n\nsin nada\n",
            evidence="docs/decisions/cierre.md")
        self.assertEqual(proc.returncode, 1)
        self.assertIn("debe ir dentro", proc.stdout)

    def test_the_type_docs_label_does_not_exempt_the_section(self):
        # Sin exención por etiqueta: la etiqueta la pone el autor del PR y no
        # es una evidencia. Mismo cuerpo, dos etiquetas, mismo veredicto.
        body = (f"Closes #12\n\n## {CLOSURE_HEADING}\n\n### {self.MARKER}\n\n"
                "sin evidencia\n")
        docs = self.run_case(body, labels=("type:docs",))
        other = self.run_case(body, labels=("type:chore",))
        self.assertEqual(docs.returncode, 1)
        self.assertEqual(other.returncode, 1)

    def test_a_section_without_the_marker_still_demands_a_test(self):
        proc = self.run_case(f"Closes #12\n\n## {CLOSURE_HEADING}\n\n"
                             "debería estar arreglado\n")
        self.assertEqual(proc.returncode, 1)
        self.assertIn("formato", proc.stdout)

    def not_applicable_body(self):
        return (f"Closes #12\n\n## {CLOSURE_HEADING}\n\n### {self.MARKER}\n\n"
                "- `docs/decisions/cierre.md`\n\n" + AI_OK)

    def test_docs_only_change_with_not_applicable_passes(self):
        proc = self.run_case(self.not_applicable_body(), changes=("README.md",),
                             evidence="docs/decisions/cierre.md")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_a_python_file_with_not_applicable_fails_naming_it(self):
        proc = self.run_case(self.not_applicable_body(), changes=("app/codigo.py",),
                             evidence="docs/decisions/cierre.md")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("app/codigo.py", proc.stdout)
        self.assertIn("test", proc.stdout)

    def test_a_workflow_with_not_applicable_fails_naming_it(self):
        proc = self.run_case(self.not_applicable_body(),
                             changes=(".github/workflows/ci.yml",),
                             evidence="docs/decisions/cierre.md")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn(".github/workflows/ci.yml", proc.stdout)

    def test_a_mixed_change_with_not_applicable_names_only_the_executable_file(self):
        proc = self.run_case(self.not_applicable_body(),
                             changes=("README.md", "app/codigo.py"),
                             evidence="docs/decisions/cierre.md")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("app/codigo.py", proc.stdout)
        self.assertNotIn("README.md", proc.stdout)

    def test_not_applicable_without_changed_files_fails_closed(self):
        proc = self.run_case(self.not_applicable_body(), changes=("README.md",),
                             evidence="docs/decisions/cierre.md", no_changed=True)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("--changed-files", proc.stdout)

    def test_the_policy_can_declare_its_own_non_executable_paths(self):
        proc = self.run_case(self.not_applicable_body(), changes=("scripts/x.py",),
                             evidence="docs/decisions/cierre.md",
                             policy={"closure_not_applicable_paths": ["scripts/**", "docs/**"]})
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
        proc = self.run_case("Refs #333\n\n" + chain_ok() + "\n" + AI_OK)
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
            "- review-budget: 12/400\n\n📍 este PR\n\n" + AI_OK,
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


class ClosingKeywordInCommitsTests(unittest.TestCase):
    """#1056 (HR-7): GitHub cierra la issue por el CUERPO del PR y también por
    los MENSAJES de commit, y HR-7 solo miraba el cuerpo.

    Un PR con `chain:partial` cuyo commit dice `(Closes #N)` cierra su issue al
    mergearse aunque el cuerpo diga `Refs #N`: así se cerraron la issue 1035 (PR
    1050, commit `27d6817d`, cerrado a los cinco segundos del merge y reabierto
    por el mantenedor) y la 1033 (PR 1037). El dato es el `--commits-file` que
    el workflow ya genera para HR-14 con merge-base (#1013): el gate no llama a
    la API ni ejecuta git para esto.
    """

    KEYWORDS = ("close", "closes", "closed", "fix", "fixes", "fixed",
                "resolve", "resolves", "resolved")

    def run_case(self, commits, *, body=None,
                 labels=("type:bug", "chain:partial"), base="feature/x"):
        tmp = tempfile.mkdtemp()
        repo, base_sha, head_sha = make_repo(tmp, commits=commits)
        event = make_event(
            tmp, body=body if body is not None else "Refs #1035\n\n" + chain_ok(base) + "\n" + AI_OK,
            labels=list(labels), base=base, base_sha=base_sha, head_sha=head_sha)
        return run_gate(event, make_policy(tmp), repo=repo)

    def test_chain_partial_con_closes_en_el_commit_falla(self):
        """Caso real (PR 1050): el cuerpo cumple HR-7 y el commit cierra la issue."""
        proc = self.run_case(
            ["fix(propagate): el banner del ensayo calcula la longitud (Closes #1035)"])
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("HR-7", proc.stdout)
        self.assertIn("Closes #1035", proc.stdout)
        self.assertIn("Refs #1035", proc.stdout)

    def test_tip_sin_chain_partial_puede_llevar_closes_en_su_commit(self):
        proc = self.run_case(
            ["feat(x): el cierre de la cadena (Closes #1035)"],
            body="Closes #1035\n\n## " + CLOSURE_HEADING + "\n\n- `tests/test_gate.py`\n"
                 + chain_block("main") + "\n" + AI_OK,
            labels=("type:bug",), base="main")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_las_nueve_palabras_clave_sin_distinguir_mayusculas(self):
        for keyword in self.KEYWORDS:
            for forma in (keyword, keyword.capitalize(), keyword.upper()):
                with self.subTest(palabra=forma):
                    proc = self.run_case([f"feat(x): algo ({forma} #1035)"])
                    self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
                    self.assertIn("HR-7", proc.stdout)

    def test_palabra_clave_sin_numero_no_cierra_nada(self):
        proc = self.run_case(["fix(x): el close del banner, sin referencia (#7 cerrado a mano)"])
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_palabra_clave_en_el_cuerpo_del_commit_tambien_falla(self):
        proc = self.run_case(["feat(x): algo\n\nCloses #1035"])
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("HR-7", proc.stdout)
class ClosingTargetsInBodyTests(unittest.TestCase):
    """#1059 (HR-7): el gate decía «tiene o no tiene una palabra de cierre», no
    «cuáles». Un PR punta que declara `Closes #1056` y que además nombra
    `Closes #1035` al citar un ejemplo pasaba igual, y GitHub cerraba las dos:
    pasó con la PR 1058 de este repo, que habría cerrado la issue 1035 —la
    que tiene trabajo pendiente— al mergearse.

    Aqui el veredicto NO cambia (cerrar varias issues en un PR punta es
    legitimo): lo que se agrega es la visibilidad de la lista, por `stderr`.
    Con `chain:partial` el hallazgo de HR-7 no se toca: ahi la palabra de cierre
    ya es el fallo.
    """

    def run_case(self, body, labels=("type:bug",), base="main"):
        tmp = tempfile.mkdtemp()
        repo, base_sha, head_sha = make_repo(tmp)
        event = make_event(tmp, body=body, labels=list(labels), base=base,
                           base_sha=base_sha, head_sha=head_sha)
        return run_gate(event, make_policy(tmp), repo=repo)

    def test_el_cuerpo_que_cierra_dos_issues_reporta_las_dos(self):
        proc = self.run_case(
            "Closes #1056\n\nEl commit decia (Closes #1035) y por eso lo arreglo.\n\n"
            "## " + CLOSURE_HEADING + "\n\n- `tests/test_gate.py`\n"
            + chain_block("main") + "\n" + AI_OK)
        # El veredicto no cambia: cerrar varias issues en el punta es legítimo.
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("#1035", proc.stderr)
        self.assertIn("#1056", proc.stderr)

    def test_una_sola_issue_se_comporta_igual_que_hoy(self):
        proc = self.run_case(
            "Closes #1056\n\n## " + CLOSURE_HEADING + "\n\n- `tests/test_gate.py`\n"
            + chain_block("main") + "\n" + AI_OK)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertNotIn("#1056", proc.stderr)

    def test_la_misma_issue_dos_veces_no_es_una_lista(self):
        proc = self.run_case(
            "Closes #1056\n\nCorrijo lo que decia Fixes #1056.\n\n"
            "## " + CLOSURE_HEADING + "\n\n- `tests/test_gate.py`\n"
            + chain_block("main") + "\n" + AI_OK)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertNotIn("#1056", proc.stderr)


class IssueLinkGateTests(unittest.TestCase):
    """#333: el enlace de la issue no admite referencias a otro repo (la mezcla Closes/Refs es aviso, #787)."""

    def run_case(self, body, labels=("type:chore",)):
        tmp = tempfile.mkdtemp()
        repo, base_sha, head_sha = make_repo(tmp)
        event = make_event(tmp, body=body, labels=list(labels), base="main",
                           base_sha=base_sha, head_sha=head_sha)
        return run_gate(event, make_policy(tmp), repo=repo)

    def test_closes_and_refs_for_the_same_issue_passes(self):
        # #787: Closes y Refs de la misma issue no es un fallo, ni aviso en issue-spec.
        proc = self.run_case("Closes #333\n\n## " + CLOSURE_HEADING
                             + "\n\n- `tests/test_gate.py`\n" + chain_ok("main") + "\nRefs #333\n\n" + AI_OK)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_foreign_repository_link_fails(self):
        proc = self.run_case("Closes #333\n\nRefs otro-org/otro-repo#12\n\n" + chain_ok("main"))
        self.assertEqual(proc.returncode, 1)
        self.assertIn("otro-org/otro-repo#12", proc.stdout)

    def test_prose_mention_of_a_foreign_issue_is_not_a_link(self):
        # Auditoría de #352: el gate marcaba CUALQUIER mención `owner/repo#N` del
        # cuerpo, y citar una issue de otro repositorio como contexto —el piloto
        # de Cadete lo hace a diario— no es enlazarla.
        proc = self.run_case("Refs #333\n\nContexto: se cita DysTelefonica/cadete#318 "
                             "como antecedente del piloto.\n\n" + chain_ok("main") + "\n" + AI_OK,
                             labels=("type:chore", "chain:partial"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_own_closing_link_passes(self):
        proc = self.run_case("Closes #333\n\n## " + CLOSURE_HEADING
                             + "\n\n- `tests/test_gate.py`\n" + chain_ok("main") + "\n" + AI_OK)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

ALL_CHECKS = ("labels", "closes", "chain_context", "issue_links", "closure_tests",
              "docs_evidence", "budget", "commits", "ai_disclosure")
ONLY_DOCS = {name: False for name in ALL_CHECKS if name != "docs_evidence"}
DOCS_EVIDENCE = {"pages": [{"page": "docs/p.md", "covers": ["f.txt"]}], "checklists": {}}


class ChecksSwitchTests(unittest.TestCase):
    """#735: la clave `checks` de la politica enciende o apaga cada comprobacion."""

    def run_case(self, checks, *, body="sin nada\n", labels=(), docs=DOCS_EVIDENCE, **extra):
        with tempfile.TemporaryDirectory() as tmp:
            repo, base_sha, head_sha = make_repo(tmp, commits=["feat: x"])
            event = make_event(tmp, body=body, labels=list(labels), base_sha=base_sha, head_sha=head_sha)
            overrides = dict(extra)
            overrides["checks"] = checks  # None -> `null`: clave ausente a efectos del gate
            if docs is not None:
                overrides["docs_evidence"] = docs
            return run_gate(event, make_policy(tmp, **overrides), repo)

    def test_without_checks_every_check_stays_active(self):
        proc = self.run_case(None)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        for rule in ("HR-6", "HR-7", "HR-54", "HR-64"):
            self.assertIn(rule, proc.stdout)

    def test_only_docs_evidence_ignores_the_rest_of_the_contract(self):
        proc = self.run_case(ONLY_DOCS, body="## Documentación\n\nNo cambia el contrato de f.txt.\n")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_only_docs_evidence_still_fails_hr64_and_nothing_else(self):
        proc = self.run_case(ONLY_DOCS)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("HR-64", proc.stdout)
        for rule in ("HR-6:", "HR-7:", "HR-53", "HR-54"):
            self.assertNotIn(rule, proc.stdout)

    def test_disabled_checks_are_listed_on_stderr(self):
        proc = self.run_case(ONLY_DOCS, body="## Documentación\n\nmotivo\n")
        self.assertIn("apagadas por política", proc.stderr)
        for name in ONLY_DOCS:
            self.assertIn(name, proc.stderr)

    def test_docs_evidence_switched_off_is_not_enforced_even_if_declared(self):
        checks = {name: False for name in ALL_CHECKS}
        checks["labels"] = True
        proc = self.run_case(checks, labels=("type:feature",))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_unknown_check_name_exits_2(self):
        proc = self.run_case({**ONLY_DOCS, "docs_evidnce": True})
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("docs_evidnce", proc.stderr)

    def test_non_boolean_value_exits_2(self):
        proc = self.run_case({**ONLY_DOCS, "docs_evidence": "yes"})
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("docs_evidence", proc.stderr)

    def test_checks_not_an_object_exits_2(self):
        proc = self.run_case(["docs_evidence"])
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)

    def test_every_check_switched_off_exits_2(self):
        proc = self.run_case({name: False for name in ALL_CHECKS})
        self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
        self.assertIn("ninguna", proc.stderr)

class RangeDataFilesTests(unittest.TestCase):
    """#753: el gate no ejecuta git; el workflow le pasa los datos del rango por fichero."""

    FLAGS = {"closure_tests": "--tracked-files", "docs_evidence": "--changed-files",
             "budget": "--numstat-file", "commits": "--commits-file"}

    def _only(self, check, tmp, flags):
        checks = {name: name == check for name in ALL_CHECKS}
        policy = make_policy(tmp, checks=checks, docs_evidence=DOCS_EVIDENCE)
        event = make_event(tmp, body="x\n", labels=["type:chore"])
        return run_gate(event, policy, data_flags=flags)

    def test_the_gate_never_imports_subprocess(self):
        source = GATE.read_text(encoding="utf-8")
        self.assertNotIn("subprocess", source)

    def test_each_active_check_without_its_file_exits_2_naming_the_flag(self):
        for check, flag in self.FLAGS.items():
            with self.subTest(check=check), tempfile.TemporaryDirectory() as tmp:
                proc = self._only(check, tmp, [])
                self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
                self.assertIn(flag, proc.stderr)
                self.assertIn(check, proc.stderr)

    def test_unreadable_data_file_exits_2(self):
        with tempfile.TemporaryDirectory() as tmp:
            proc = self._only("budget", tmp, ["--numstat-file", str(Path(tmp) / "no-existe.txt")])
            self.assertEqual(proc.returncode, 2, proc.stdout + proc.stderr)
            self.assertIn("--numstat-file", proc.stderr)

    def test_a_check_that_is_off_does_not_need_its_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            checks = {name: name == "labels" for name in ALL_CHECKS}
            event = make_event(tmp, body="x\n", labels=["type:chore"])
            proc = run_gate(event, make_policy(tmp, checks=checks), data_flags=[])
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_budget_reads_the_numstat_file_and_ignores_binaries(self):
        with tempfile.TemporaryDirectory() as tmp:
            numstat = Path(tmp) / "numstat.txt"
            numstat.write_text("300\t150\ta.py\n-\t-\timg.png\n", encoding="utf-8")
            proc = self._only("budget", tmp, ["--numstat-file", str(numstat)])
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertIn("450 > 400", proc.stdout)
            numstat.write_text("300\t100\ta.py\n-\t-\timg.png\n", encoding="utf-8")
            proc = self._only("budget", tmp, ["--numstat-file", str(numstat)])
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)


class MovedBaseRangeTests(unittest.TestCase):
    """HR-64 tras reapuntar una PR apilada: el rango de los ficheros es el de la PR, no el de main."""

    COVERED = {"checks": {name: name == "docs_evidence" for name in ALL_CHECKS},
               "docs_evidence": {"pages": [{"page": "docs/codebase/scripts.md",
                                           "covers": ["scripts/check_*.py"]}],
                                 "checklists": {}}}

    def _git(self, repo, *args):
        subprocess.run(["git", "-C", str(repo), "-c", "user.name=t", "-c", "user.email=t@e.invalid",
                        "-c", "commit.gpgsign=false", *args], check=True, capture_output=True)

    def _commit(self, repo, msg, files):
        for rel, text in files.items():
            (repo / rel).parent.mkdir(parents=True, exist_ok=True)
            (repo / rel).write_text(text, encoding="utf-8")
        self._git(repo, "add", "-A")
        self._git(repo, "commit", "-q", "-m", msg)
        return git_text(repo, "rev-parse", "HEAD").strip()

    def _stacked_pr_after_main_moved(self, tmp):
        """PR de una rama sobre B0 que solo toca tests/. Después main mergea otra PR
        (M1, toca scripts/) y la PR se reapunta a main: base.sha = M1, head = F1."""
        repo = Path(tmp) / "repo"
        repo.mkdir()
        self._git(repo, "init", "-q", "-b", "main")
        self._commit(repo, "base", {"docs/codebase/scripts.md": "scripts\n"})
        self._git(repo, "checkout", "-q", "-b", "feature/x")
        head = self._commit(repo, "feat: la PR", {"tests/test_own.py": "x = 1\n"})
        self._git(repo, "checkout", "-q", "main")
        base = self._commit(repo, "merge de otra PR", {"scripts/check_crap.py": "y = 1\n"})
        return repo, base, head

    def test_pr_que_solo_toca_tests_no_hereda_lo_que_main_mergeo(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo, base, head = self._stacked_pr_after_main_moved(tmp)
            event = make_event(tmp, body="x\n", labels=["type:chore"], base_sha=base, head_sha=head)
            proc = run_gate(event, make_policy(tmp, **self.COVERED), repo=repo)
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_el_rango_documentado_del_gate_es_el_de_tres_puntos(self):
        done = subprocess.run([sys.executable, str(GATE), "--help"], capture_output=True,
                              text=True, check=True)
        self.assertIn("git diff --name-only BASE...HEAD", " ".join(done.stdout.split()))


def ai_section(text):
    return f"## Asistencia de IA\n\n{text}\n"


class AiDisclosureGateTests(unittest.TestCase):
    """#876: el cuerpo del PR declara la asistencia de IA en la sección «Asistencia
    de IA». Fail-closed: ausente, vacía o solo con la plantilla no pasa; «ninguna» o
    una descripción válida sí. La declaración vive en el PR, nunca en los commits."""

    def run_case(self, body, labels=("type:feature",), base="feature/x"):
        tmp = tempfile.mkdtemp()
        repo, base_sha, head_sha = make_repo(tmp)
        event = make_event(tmp, body=body, labels=list(labels), base=base,
                           base_sha=base_sha, head_sha=head_sha)
        return run_gate(event, make_policy(tmp), repo=repo)

    def test_missing_section_fails_naming_it(self):
        proc = self.run_case("Refs #876\n")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("Asistencia de IA", proc.stdout)

    def test_missing_body_fails_naming_it(self):
        proc = self.run_case(None)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("Asistencia de IA", proc.stdout)

    def test_empty_section_fails(self):
        proc = self.run_case("Refs #876\n\n## Asistencia de IA\n\n")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("Asistencia de IA", proc.stdout)

    def test_section_with_only_template_comment_fails(self):
        body = ("Refs #876\n\n## Asistencia de IA\n\n"
                "<!-- Escriba «ninguna» o la herramienta y la parte ayudada. -->\n")
        proc = self.run_case(body)
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("Asistencia de IA", proc.stdout)

    def test_declaration_as_prose_without_heading_fails(self):
        proc = self.run_case("Refs #876\n\nAsistencia de IA: ninguna.\n")
        self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
        self.assertIn("Asistencia de IA", proc.stdout)

    def test_ninguna_passes(self):
        proc = self.run_case("Refs #876\n\n" + ai_section("ninguna"))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_valid_description_passes(self):
        text = "Copilot redactó el primer borrador del test; lo revisé y lo reescribí."
        proc = self.run_case("Refs #876\n\n" + ai_section(text))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_propagator_body_passes(self):
        proc = self.run_case("Refs #876\n\n" + ai_section("ninguna — PR generado por el propagador."))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)


# #938: el gate lee `labels` y `body` desde la API en vivo, no del payload del
# evento. Sin `GITHUB_TOKEN` o con la API caída: exit 2 con motivo (HR-3); el
# cuerpo del payload no se imprime en `stderr`. El precedent es #830/#831 para
# `issue-spec`; este gate aplica el mismo eje a `pr-contract`.
LIVE_REPO = "team/skills"
LIVE_PR = 7


class _MockApiServer:
    """HTTP server local (stdlib) sirviendo los dos endpoints del API que
    consume el gate (`/repos/{r}/issues/{n}/labels` y `/repos/{r}/pulls/{n}`)."""

    def __init__(self, pull_body, labels):
        import threading
        from http.server import BaseHTTPRequestHandler, HTTPServer
        from urllib.parse import urlsplit
        state = {"pull": pull_body, "labels": labels}
        routes = {
            f"/repos/{LIVE_REPO}/issues/{LIVE_PR}/labels": state["labels"],
            f"/repos/{LIVE_REPO}/pulls/{LIVE_PR}": state["pull"],
        }

        class _Handler(BaseHTTPRequestHandler):
            def do_GET(self):  # noqa: N802
                raw = json.dumps(routes.get(urlsplit(self.path).path, None)).encode("utf-8")
                if raw == b"null":
                    self.send_response(404)
                    self.end_headers()
                    return
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(raw)))
                self.end_headers()
                self.wfile.write(raw)

            def log_message(self, *_a, **_kw): return

        self._server = HTTPServer(("127.0.0.1", 0), _Handler)
        threading.Thread(target=self._server.serve_forever, daemon=True).start()
        self.base_url = f"http://127.0.0.1:{self._server.server_address[1]}"

    def close(self):
        self._server.shutdown()


def _good_api_body():
    """`pull_response` válido: dict con `body` cumplidor."""
    return {"body": ("Closes #42\n\n## Tests que prueban el cierre\n\n"
                     "- `tests/test_gate.py`\n" + chain_block("main") + "\n" + AI_OK)}


class LiveApiTests(unittest.TestCase):
    """#938 — etiquetas y cuerpo por API; fail-closed sin token o con API caída."""

    LIVE_POLICY = {**POLICY_DATA,
                   "checks": {n: (n == "labels") for n in
                              ("labels", "closes", "chain_context", "issue_links",
                               "closure_tests", "docs_evidence", "budget",
                               "commits", "ai_disclosure")}}

    def _event(self, tmp, payload_labels=None, payload_body=""):
        payload = {
            "action": "opened",
            "pull_request": {
                "number": LIVE_PR,
                "base": {"ref": "main", "sha": "0" * 40},
                "head": {"ref": "feature/x", "sha": "1" * 40},
                "labels": [{"name": n} for n in (payload_labels or [])],
                "body": payload_body, "title": "test",
            },
            "repository": {"full_name": LIVE_REPO, "default_branch": "main"},
        }
        event = Path(tmp) / "event.json"
        event.write_text(json.dumps(payload), encoding="utf-8")
        return event

    def _run(self, env, payload_labels=None, payload_body=""):
        # #938: el modo vivo es opt-in; esta suite lo activa explícitamente.
        env.setdefault("PR_CONTRACT_USE_LIVE_API", "1")
        with tempfile.TemporaryDirectory() as tmp:
            event = self._event(tmp, payload_labels=payload_labels,
                                payload_body=payload_body)
            policy = Path(tmp) / "policy.json"
            policy.write_text(json.dumps(self.LIVE_POLICY), encoding="utf-8")
            return subprocess.run(
                [sys.executable, str(GATE), "--event-file", str(event),
                 "--policy-file", str(policy)],
                capture_output=True, text=True, check=False, env=env, timeout=60)

    def test_payload_labels_missing_but_api_labels_present_passes(self):
        """Snapshot sin etiquetas, API con `type:chore`+`chain:partial` → 0."""
        env = {k: v for k, v in os.environ.items() if k != "GITHUB_TOKEN"}
        env["GITHUB_TOKEN"] = "test-token"
        server = _MockApiServer(pull_body=_good_api_body(),
                                labels=[{"name": "type:chore"}, {"name": "chain:partial"}])
        try:
            env["GITHUB_API_URL"] = server.base_url
            proc = self._run(env, payload_labels=[], payload_body="cuerpo obsoleto")
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        finally:
            server.close()

    def test_payload_labels_present_but_api_labels_absent_fails(self):
        """Snapshot con `type:chore`, API sin etiquetas → 1 con `HR-6`."""
        env = {k: v for k, v in os.environ.items() if k != "GITHUB_TOKEN"}
        env["GITHUB_TOKEN"] = "test-token"
        body = _good_api_body()["body"]
        server = _MockApiServer(pull_body={"body": body}, labels=[])
        try:
            env["GITHUB_API_URL"] = server.base_url
            proc = self._run(env, payload_labels=["type:chore"], payload_body=body)
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertIn("HR-6", proc.stdout)
        finally:
            server.close()

    def test_pr_event_without_token_fails_closed_with_explicit_reason(self):
        """Sin `GITHUB_TOKEN` ni `GH_TOKEN` → 2 con motivo del cliente
        compartido; el cuerpo del payload no se imprime."""
        env = {k: v for k, v in os.environ.items()
               if k not in ("GITHUB_TOKEN", "GH_TOKEN")}
        proc = self._run(env, payload_labels=["type:chore"],
                         payload_body="Closes #42 - cuerpo unico")
        self.assertEqual(proc.returncode, 2, proc.stderr)
        self.assertIn("GITHUB_TOKEN", proc.stderr)
        self.assertNotIn("cuerpo unico", proc.stderr)

    def test_pr_event_api_outage_fails_closed_with_explicit_reason(self):
        """API caída (puerto cerrado) → 2 con motivo de transporte."""
        env = {k: v for k, v in os.environ.items()
               if k not in ("GITHUB_TOKEN", "GH_TOKEN")}
        env["GITHUB_TOKEN"] = "test-token"
        env["GITHUB_API_URL"] = "http://127.0.0.1:9"  # puerto cerrado
        proc = self._run(env, payload_labels=["type:chore"],
                         payload_body="Closes #42 - cuerpo unico")
        self.assertEqual(proc.returncode, 2, proc.stderr)
        self.assertIn("transport error", proc.stderr)
        self.assertNotIn("cuerpo unico", proc.stderr)

    def test_default_mode_reads_payload_without_token_or_live_api(self):
        """#938 (review #940 ronda 3): por defecto el gate lee labels y body
        del payload, sin llamar a la API. Este modo es el contrato del modo
        local (`bin/pr_contract_local.py`), del propagador PowerShell, de la
        suite de parity y de la suite de propagación: todos invocan el gate
        como subproceso sin `GITHUB_TOKEN` y deben seguir funcionando igual
        que en `origin/main`. El modo vivo por API es opt-in y se activa
        solo con `PR_CONTRACT_USE_LIVE_API=1`; esta suite NO lo setea, así
        que el gate evalúa el payload directamente."""
        with tempfile.TemporaryDirectory() as tmp:
            # El payload del evento tiene labels y body explícitos; el
            # gate los debe respetar tal cual.
            event = self._event(tmp,
                                payload_labels=["type:bug"],
                                payload_body="Closes #42\n\n## Tests\n\n- `t.py`\n")
            policy = Path(tmp) / "policy.json"
            policy.write_text(json.dumps(self.LIVE_POLICY), encoding="utf-8")
            env = {k: v for k, v in os.environ.items()
                   if k not in ("GITHUB_TOKEN", "PR_CONTRACT_USE_LIVE_API",
                                "GITHUB_EVENT_PATH")}
            self.assertNotIn("PR_CONTRACT_USE_LIVE_API", env)
            self.assertNotIn("GITHUB_TOKEN", env)
            proc = subprocess.run(
                [sys.executable, str(GATE), "--event-file", str(event),
                 "--policy-file", str(policy)],
                capture_output=True, text=True, check=False, env=env, timeout=60)
            # Sin opt-in ni token, el gate es verde porque las labels y el
            # body del payload cumplen el contrato (HR-6 y la política).
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)


    # --- #996: el modo vivo es el defecto cuando se juzga el evento real de Actions ---

    def _run_actions(self, env, *, payload_labels=None, payload_body="", event_path_is_actions=True):
        """Como en CI: sin `PR_CONTRACT_USE_LIVE_API` y con `GITHUB_EVENT_PATH` apuntando al
        evento que se juzga (o a otro fichero, si `event_path_is_actions` es falso)."""
        env = dict(env)
        env.pop("PR_CONTRACT_USE_LIVE_API", None)
        with tempfile.TemporaryDirectory() as tmp:
            event = self._event(tmp, payload_labels=payload_labels, payload_body=payload_body)
            env["GITHUB_EVENT_PATH"] = str(event if event_path_is_actions else Path(tmp) / "otro.json")
            policy = Path(tmp) / "policy.json"
            policy.write_text(json.dumps(self.LIVE_POLICY), encoding="utf-8")
            return subprocess.run(
                [sys.executable, str(GATE), "--event-file", str(event),
                 "--policy-file", str(policy)],
                capture_output=True, text=True, check=False, env=env, timeout=60)

    def _token_env(self, server):
        env = {k: v for k, v in os.environ.items() if k not in ("GITHUB_TOKEN", "GH_TOKEN")}
        env["GITHUB_TOKEN"] = "test-token"
        env["GITHUB_API_URL"] = server.base_url
        return env

    def test_996_actions_event_with_token_reads_live_labels_by_default(self):
        """Payload congelado sin etiquetas, API con `type:chore`: el PR etiquetado después de
        abrirse pasa HR-6 sin cerrar y reabrir ni activar nada."""
        server = _MockApiServer(pull_body=_good_api_body(), labels=[{"name": "type:chore"}])
        try:
            proc = self._run_actions(self._token_env(server), payload_labels=[],
                                     payload_body="cuerpo obsoleto")
            self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        finally:
            server.close()

    def test_996_live_mode_does_not_relax_the_contract(self):
        """Sin etiquetas en la API sigue siendo HR-6 aunque el payload las trajera."""
        body = _good_api_body()["body"]
        server = _MockApiServer(pull_body={"body": body}, labels=[])
        try:
            proc = self._run_actions(self._token_env(server), payload_labels=["type:chore"],
                                     payload_body=body)
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertIn("HR-6", proc.stdout)
        finally:
            server.close()

    def test_996_explicit_opt_out_keeps_reading_the_payload(self):
        """`PR_CONTRACT_USE_LIVE_API=0` vuelve al payload: sin etiquetas en él, HR-6."""
        server = _MockApiServer(pull_body=_good_api_body(), labels=[{"name": "type:chore"}])
        try:
            env = self._token_env(server)
            with tempfile.TemporaryDirectory() as tmp:
                event = self._event(tmp, payload_labels=[], payload_body="cuerpo obsoleto")
                env.update(GITHUB_EVENT_PATH=str(event), PR_CONTRACT_USE_LIVE_API="0")
                policy = Path(tmp) / "policy.json"
                policy.write_text(json.dumps(self.LIVE_POLICY), encoding="utf-8")
                proc = subprocess.run(
                    [sys.executable, str(GATE), "--event-file", str(event),
                     "--policy-file", str(policy)],
                    capture_output=True, text=True, check=False, env=env, timeout=60)
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertIn("HR-6", proc.stdout)
        finally:
            server.close()

    def test_996_default_live_mode_fails_closed_when_the_api_is_down(self):
        env = {k: v for k, v in os.environ.items() if k not in ("GITHUB_TOKEN", "GH_TOKEN")}
        env["GITHUB_TOKEN"] = "test-token"
        env["GITHUB_API_URL"] = "http://127.0.0.1:9"  # puerto cerrado
        proc = self._run_actions(env, payload_labels=["type:chore"],
                                 payload_body="Closes #42 - cuerpo unico")
        self.assertEqual(proc.returncode, 2, proc.stderr)
        self.assertIn("transport error", proc.stderr)
        self.assertNotIn("cuerpo unico", proc.stderr)

    def test_996_actions_event_without_token_keeps_reading_the_payload(self):
        """Sin token no hay API que consultar: el gate sigue en modo payload, como en local."""
        env = {k: v for k, v in os.environ.items() if k not in ("GITHUB_TOKEN", "GH_TOKEN")}
        proc = self._run_actions(env, payload_labels=["type:bug"],
                                 payload_body="Closes #42\n\n## Tests\n\n- `t.py`\n")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)

    def test_996_an_event_file_that_is_not_the_actions_event_stays_in_payload_mode(self):
        """Los scripts locales, el propagador y la parity pasan su propio evento: no cambian."""
        server = _MockApiServer(pull_body=_good_api_body(), labels=[{"name": "type:chore"}])
        try:
            proc = self._run_actions(self._token_env(server), payload_labels=[],
                                     payload_body="cuerpo obsoleto",
                                     event_path_is_actions=False)
            self.assertEqual(proc.returncode, 1, proc.stdout + proc.stderr)
            self.assertIn("HR-6", proc.stdout)
        finally:
            server.close()


if __name__ == "__main__":
    unittest.main(verbosity=2)
