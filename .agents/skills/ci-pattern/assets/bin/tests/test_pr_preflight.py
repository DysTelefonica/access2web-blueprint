#!/usr/bin/env python3
"""Suite de paridad de `ci-pattern pr preflight` (#865).

Verifica que el preflight emite el mismo veredicto y texto que cada gate
individual reutilizado, sin reglas nuevas ni copia de lógica. Cada caso
carga el módulo en proceso (mismo patrón que las suites de gate) y
compara contra el texto literal que el gate emite.
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
import threading
import unittest
from contextlib import redirect_stderr, redirect_stdout
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

PREFLIGHT = Path(__file__).resolve().parents[1] / "pr_preflight.py"
ISSUE_SPEC_GATE = Path(__file__).resolve().parents[2] / "issue-spec" / "check_issue_spec.py"
BRANCH_NAME_GATE = Path(__file__).resolve().parents[2] / "branch-name" / "check_branch_name.py"
PR_CONTRACT_GATE = Path(__file__).resolve().parents[2] / "pr-contract" / "check_pr_contract.py"
DEFAULT_BRANCH_PATTERN = (
    r"^(feat|fix|chore|docs|style|refactor|perf|test|build|ci|revert)/[a-z0-9._-]+$")


def _load(name, path):
    spec = importlib.util.spec_from_loader(
        name, importlib.machinery.SourceFileLoader(name, str(path)))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


def _run_cli(*argv):
    env = {k: v for k, v in os.environ.items()
           if k not in ("GITHUB_TOKEN", "GITHUB_EVENT_PATH")}
    env["GITHUB_API_URL"] = "http://127.0.0.1:9"
    proc = subprocess.run(  # noqa: S603 — CLI del propio repo, args controlados.
        [sys.executable, str(PREFLIGHT), *argv],
        capture_output=True, text=True, check=False, env=env)
    return proc.returncode, proc.stdout, proc.stderr


def _invoke(preflight, kwargs):
    kwargs.setdefault("branch", None)
    out, err = io.StringIO(), io.StringIO()
    with redirect_stdout(out), redirect_stderr(err):
        code = preflight.run_preflight(**kwargs)
    return code, out.getvalue(), err.getvalue()


def _body_file(tmp, body):
    path = Path(tmp) / "body.md"
    path.write_text(body, encoding="utf-8")
    return str(path)


class BranchNameParityTests(unittest.TestCase):
    """Mismos casos que ``test_check_branch_name.py``."""

    @classmethod
    def setUpClass(cls):
        cls.preflight = _load("pr_preflight_under_test_branch", PREFLIGHT)
        # Carga previa bajo el nombre canónico para que ``_load_gate`` del
        # preflight (``name="check_branch_name"``) recupere el MISMO módulo y
        # la misma clase ``BadPatternError``/``EmptyRefError``.
        _load("check_branch_name", BRANCH_NAME_GATE)
        cls.tmp = tempfile.mkdtemp()

    def _kwargs(self, branch, **overrides):
        bf = _body_file(self.tmp, "Closes #12\n")
        kw = dict(body_file=bf, branch=branch)
        kw.update(overrides)
        return kw

    def test_conforming_branch_passes(self):
        for branch in ("feat/pr-preflight", "fix/a", "docs/v1.2_notes"):
            with self.subTest(branch=branch):
                code, _, err = _invoke(self.preflight, self._kwargs(branch))
                self.assertEqual(code, 0)
                self.assertIn("[branch-name] OK", err)

    def test_non_conforming_branch_warns_without_blocking(self):
        for branch in ("Feat/x", "feature/x", "feat/", "main2"):
            with self.subTest(branch=branch):
                code, _, err = _invoke(self.preflight, self._kwargs(branch))
                self.assertEqual(code, 0, err)
                self.assertIn("[branch-name] AVISO", err)
                self.assertNotIn("FAIL", err)
                self.assertIn(branch, err)
                self.assertIn("<type>/<", err)

    def test_empty_branch_fails_with_exit_two(self):
        code, _, err = _invoke(self.preflight, self._kwargs(""))
        self.assertEqual(code, 2)
        self.assertIn("the head ref is empty", err)

    def test_unreadable_pattern_is_exit_two(self):
        code, _, err = _invoke(
            self.preflight, self._kwargs("feat/x", branch_pattern="("))
        self.assertEqual(code, 2)
        self.assertIn("pattern", err.lower())


class IssueSpecLocalParityTests(unittest.TestCase):
    """Mismos casos que ``ReferenceFormTests`` del gate."""

    @classmethod
    def setUpClass(cls):
        cls.preflight = _load("pr_preflight_under_test_issuelocal", PREFLIGHT)
        cls.tmp = tempfile.mkdtemp()

    def test_closes_issue_passes_locally(self):
        bf = _body_file(self.tmp, "Closes #12\n")
        code, _, err = _invoke(self.preflight, dict(body_file=bf))
        self.assertEqual(code, 0)
        self.assertIn("[issue-spec]", err)

    def test_no_reference_fails_and_points_to_depends_on(self):
        bf = _body_file(self.tmp, "Sin enlace a ninguna issue.\n")
        code, _, err = _invoke(self.preflight, dict(body_file=bf))
        self.assertEqual(code, 1)
        self.assertIn("[issue-spec] issue-spec: el cuerpo no enlaza ninguna issue", err)

    def test_cross_repository_reference_fails_when_repo_known(self):
        bf = _body_file(self.tmp, "Closes other/repo#12\n")
        code, _, err = _invoke(self.preflight, dict(body_file=bf, repository="team/skills"))
        self.assertEqual(code, 1)
        self.assertIn("la referencia apunta a otro repositorio", err)

    def test_refs_inside_inline_code_is_not_a_reference(self):
        bf = _body_file(self.tmp, "Ejemplo: `Refs #12`, nada más.\n")
        code, _, err = _invoke(self.preflight, dict(body_file=bf))
        self.assertEqual(code, 1)
        self.assertIn("[issue-spec] issue-spec: el cuerpo no enlaza", err)


class IssueSpecRemoteParityTests(unittest.TestCase):
    """Mismos casos que ``RemoteRuleTests`` del gate."""

    @classmethod
    def setUpClass(cls):
        # Carga previa bajo el nombre canónico para que ``_load_gate`` del
        # preflight (``name="check_issue_spec"``) recupere el MISMO módulo y
        # la misma clase ``GitHubApiError``.
        cls.issue_spec = _load("check_issue_spec", ISSUE_SPEC_GATE)
        cls.preflight = _load("pr_preflight_under_test_issueremote", PREFLIGHT)
        cls.tmp = tempfile.mkdtemp()

    def _fetcher(self, *labels):
        def fetch(number):
            return {"number": number, "labels": [{"name": n} for n in labels],
                    "body": ""}
        return fetch

    def test_approved_typed_issue_passes(self):
        bf = _body_file(self.tmp, "Closes #12\n")
        code, _, err = _invoke(self.preflight,
                               dict(body_file=bf, repository="owner/repo",
                                    fetch_issue=self._fetcher(
                                        "status:approved", "type:feature")))
        self.assertEqual(code, 0, err)
        self.assertIn("[issue-spec]", err)

    def test_missing_approval_fails_with_gate_text(self):
        bf = _body_file(self.tmp, "Closes #12\n")
        code, _, err = _invoke(self.preflight,
                               dict(body_file=bf, repository="owner/repo",
                                    fetch_issue=self._fetcher("type:feature")))
        self.assertEqual(code, 1)
        self.assertIn("[issue-spec] issue-spec: #12: requiere status:approved", err)

    def test_missing_type_label_fails_with_gate_text(self):
        bf = _body_file(self.tmp, "Closes #12\n")
        code, _, err = _invoke(self.preflight,
                               dict(body_file=bf, repository="owner/repo",
                                    fetch_issue=self._fetcher("status:approved")))
        self.assertEqual(code, 1)
        self.assertIn("[issue-spec] issue-spec: #12: requiere una etiqueta type:", err)

    def test_api_error_fails_closed(self):
        def fetch(_number):
            raise self.issue_spec.GitHubApiError("upstream timeout")  # noqa: TRY003
        bf = _body_file(self.tmp, "Closes #12\n")
        code, _, err = _invoke(self.preflight,
                               dict(body_file=bf, repository="owner/repo",
                                    fetch_issue=fetch))
        self.assertEqual(code, 2)
        self.assertIn("cannot read", err)

    def _by_number(self, labels_by_number):
        def fetch(number):
            return {"number": number, "body": "",
                    "labels": [{"name": n} for n in labels_by_number[number]]}
        return fetch

    def test_closes_approved_with_non_approved_refs_passes(self):
        # #913 (a): la Refs no aprobada es informativa.
        bf = _body_file(self.tmp, "Closes #12\nRefs #13\n")
        code, _, err = _invoke(self.preflight,
                               dict(body_file=bf, repository="owner/repo",
                                    fetch_issue=self._by_number({
                                        12: ["status:approved", "type:bug"],
                                        13: ["type:bug"]})))
        self.assertEqual(code, 0, err)

    def test_only_refs_to_non_approved_issue_fails(self):
        # #913 (b): sin ninguna issue aprobada el PR sigue en rojo.
        bf = _body_file(self.tmp, "Refs #13\n")
        code, _, err = _invoke(self.preflight,
                               dict(body_file=bf, repository="owner/repo",
                                    fetch_issue=self._by_number({13: ["type:bug"]})))
        self.assertEqual(code, 1)
        self.assertIn("[issue-spec]", err)
        self.assertIn("aprobada", err)

    def test_closes_non_approved_with_approved_refs_fails(self):
        # #913 (c): el issue que se cierra debe estar aprobado.
        bf = _body_file(self.tmp, "Closes #12\nRefs #13\n")
        code, _, err = _invoke(self.preflight,
                               dict(body_file=bf, repository="owner/repo",
                                    fetch_issue=self._by_number({
                                        12: ["type:bug"],
                                        13: ["status:approved", "type:bug"]})))
        self.assertEqual(code, 1)
        self.assertIn("[issue-spec] issue-spec: #12: requiere status:approved", err)


class DocsEvidenceParityTests(unittest.TestCase):
    """Paridad con ``check_docs_evidence`` de ``check_pr_contract.py``."""

    @classmethod
    def setUpClass(cls):
        cls.preflight = _load("pr_preflight_under_test_docs", PREFLIGHT)

    def _kwargs(self, *, policy_text, body, files):
        tmp = Path(tempfile.mkdtemp())
        policy_path = tmp / "policy.json"
        policy_path.write_text(policy_text, encoding="utf-8")
        files_path = tmp / "files.txt"
        files_path.write_text("\n".join(files), encoding="utf-8")
        body_path = tmp / "body.md"
        body_path.write_text(body, encoding="utf-8")
        return dict(body_file=str(body_path), policy=str(policy_path),
                    changed_files_file=str(files_path))

    def test_missing_docs_evidence_emits_gate_warning(self):
        policy = json.dumps({"required_labels": ["type:bug"]})
        code, _, err = _invoke(self.preflight, self._kwargs(
            policy_text=policy, body="Closes #12\n", files=["app/x.py"]))
        self.assertEqual(code, 0)
        self.assertIn("[docs-evidence] HR-64: docs_evidence no declarado", err)

    def test_touched_doc_page_without_section_fails(self):
        policy = json.dumps({"docs_evidence": {
            "pages": [{"page": "docs/x.md", "covers": ["app/x.py"]}],
            "checklists": {}}})
        code, _, err = _invoke(self.preflight, self._kwargs(
            policy_text=policy, body="Closes #12\n", files=["app/x.py"]))
        self.assertEqual(code, 1)
        self.assertIn("[docs-evidence] HR-64: el PR toca app/x.py", err)
        self.assertIn("Documentación", err)


class OutputFormatTests(unittest.TestCase):
    """Formato de salida y JSON."""

    @classmethod
    def setUpClass(cls):
        cls.preflight = _load("pr_preflight_under_test_format", PREFLIGHT)
        cls.tmp = tempfile.mkdtemp()

    def test_clean_run_prints_each_gate_ok(self):
        bf = _body_file(self.tmp, "Closes #12\n")
        code, _, err = _invoke(self.preflight, dict(body_file=bf, branch="feat/pr-preflight"))
        self.assertEqual(code, 0)
        self.assertIn("[branch-name] OK", err)
        self.assertIn("[issue-spec]", err)

    def test_json_output(self):
        bf = _body_file(self.tmp, "Sin enlace.\n")
        code, out, _ = _invoke(self.preflight,
                               dict(body_file=bf, branch="feat/x", as_json=True))
        self.assertEqual(code, 1)
        payload = json.loads(out)
        self.assertFalse(payload["clean"])
        gates = {f["gate"] for f in payload["findings"]}
        self.assertIn("branch-name", gates)
        self.assertIn("issue-spec", gates)

    def test_missing_body_file_is_usage_error(self):
        code, _, err = _run_cli("--pr-body", "/no/such/file.md")
        self.assertEqual(code, 2, err)
        self.assertIn("cuerpo", err.lower())


class GateReuseContractTests(unittest.TestCase):
    """El preflight NO copia lógica: reusa las funciones de los gates."""

    def test_branch_name_reuses_evaluate(self):
        # El preflight reusa ``evaluate`` byte a byte; si copiara lógica, el
        # texto divergiría. Este test pinea el contrato del gate.
        gate = _load("check_branch_name", BRANCH_NAME_GATE)
        ok, _ = gate.evaluate("feat/x", gate.compile_pattern(DEFAULT_BRANCH_PATTERN),
                              frozenset(), "")
        self.assertTrue(ok)

    def test_issue_spec_reuses_validate_body(self):
        gate = _load("check_issue_spec", ISSUE_SPEC_GATE)
        body = "Sin enlace.\n"
        # Idempotencia: el veredicto es estable.
        self.assertEqual(gate.validate_body(body), gate.validate_body(body))


def _issue_api_handler(issues):
    """Fabrica una clase Handler con su propio dict de issues (#941).

    Sustituye al antiguo atributo de clase que el modulo tenia antes (un
    ``issues = {}`` a nivel de clase era mutable y se compartia entre clases
    de test del mismo modulo). Cada Stub lleva su propio ``issues`` en el
    cierre, aislado de cualquier otro Stub
    paralelo.
    """

    class Stub(BaseHTTPRequestHandler):
        def do_GET(self):  # noqa: N802 — nombre fijado por http.server.
            number = self.path.rsplit("/", 1)[-1]
            issue = issues.get(number)
            status, payload = (200, issue) if issue else (404, {"message": "Not Found"})
            data = json.dumps(payload).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        def log_message(self, *_args):
            return

    return Stub


class CliTokenPathTests(unittest.TestCase):
    """``main()`` por la CLI pública, con y sin ``--token`` (#924)."""

    @classmethod
    def setUpClass(cls):
        cls.preflight = _load("pr_preflight_under_test_cli_token", PREFLIGHT)
        cls.tmp = tempfile.mkdtemp()
        cls.issues = {"12": {
            "number": 12, "body": "",
            "labels": [{"name": "status:approved"}, {"name": "type:bug"}]}}
        cls.handler_cls = _issue_api_handler(cls.issues)
        cls.server = ThreadingHTTPServer(("127.0.0.1", 0), cls.handler_cls)
        cls.api_url = f"http://127.0.0.1:{cls.server.server_address[1]}"
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()

    def _main(self, *extra):
        bf = _body_file(self.tmp, "Closes #12\n")
        argv = ["--pr-body", bf, "--repository", "owner/repo",
                "--branch", "feat/924-preflight-token", *extra]
        # ``main`` fija UTF-8 con ``reconfigure``: el sustituto necesita ser un TextIOWrapper.
        out = io.TextIOWrapper(io.BytesIO(), encoding="utf-8", write_through=True)
        err_buf = io.BytesIO()
        err = io.TextIOWrapper(err_buf, encoding="utf-8", write_through=True)
        with redirect_stdout(out), redirect_stderr(err):
            code = self.preflight.main(argv)
        return code, err_buf.getvalue().decode("utf-8")

    def test_token_run_checks_linked_issue_through_the_api(self):
        code, err = self._main("--token", "fake", "--api-url", self.api_url)
        self.assertEqual(code, 0, err)
        self.assertIn("[issue-spec]", err)

    def test_run_without_token_keeps_local_only_verdict(self):
        code, err = self._main()
        self.assertEqual(code, 0, err)
        self.assertIn("[issue-spec]", err)


if __name__ == "__main__":
    unittest.main()

