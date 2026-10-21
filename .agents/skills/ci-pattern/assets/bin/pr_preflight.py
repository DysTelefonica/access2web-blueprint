#!/usr/bin/env python3
# ci-pattern subcomando `pr preflight` (#865)
"""Preflight local con la misma verificación que la CI del patrón (#865).

Corre antes de ``git push`` o ``gh pr create`` y emite los mismos veredictos
que la CI aplicaría sobre el PR abierto (rama, cuerpo, ``status:approved``
de la issue enlazada, HR-64 cuando hay política, deriva del doc operativo).

Reusa los gatees en proceso vía ``importlib`` (HR-46): cero reglas nuevas,
cero copia de lógica; el texto del hallazgo es el que cada gate emite,
prefidado por ``[<gate>]``. Códigos de salida: 0 verde · 1 hallazgos ·
2 uso o entrada inválida.
"""

from __future__ import annotations

import argparse
import importlib.machinery
import importlib.util
import json
import os
import subprocess
import sys
from pathlib import Path

BIN = Path(__file__).resolve().parent
SKILL_ROOT = BIN.parent.parent
ISSUE_SPEC_GATE = SKILL_ROOT / "assets" / "issue-spec" / "check_issue_spec.py"
BRANCH_NAME_GATE = SKILL_ROOT / "assets" / "branch-name" / "check_branch_name.py"
PR_CONTRACT_GATE = SKILL_ROOT / "assets" / "pr-contract" / "check_pr_contract.py"
CLI_BIN = BIN / "ci-pattern"

DEFAULT_BRANCH_PATTERN = (
    r"^(feat|fix|chore|docs|style|refactor|perf|test|build|ci|revert)/[a-z0-9._-]+$")
ADOPTION_TIMEOUT_S = 600  # segundos; ver PLR2004
EXIT_OK, EXIT_FINDINGS, EXIT_USAGE = 0, 1, 2
ADOPTION_EXIT_FAIL_CLOSED = 2  # rc del CLI ci-pattern ante contrato ausente o inválido


class _PreflightInputError(Exception):
    """Error de uso del preflight (mismo canal que ``FileNotFoundError``
    cuando se prefiere una excepción propia; TRY003)."""


def _pin_output_encoding() -> None:
    """Pin stdout/stderr a UTF-8 (vendorable-guard)."""
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")


def _load_gate(name, path):
    """Importa un gate en proceso; respeta ``sys.modules`` para que las clases
    del gate (p. ej. ``GitHubApiError``) sean las mismas que el caller ya tiene
    cargadas (#865).
    """
    if name in sys.modules:
        return sys.modules[name]
    spec = importlib.util.spec_from_loader(
        name, importlib.machinery.SourceFileLoader(name, str(path)))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


def _gate_branch(branch, pattern, actor, gate):
    try:
        rx = gate.compile_pattern(pattern)
    except gate.BadPatternError as exc:
        return [("branch-name", "bad-pattern", str(exc))]
    try:
        ok, note = gate.evaluate(branch, rx, frozenset(), actor)
    except gate.EmptyRefError:
        return [("branch-name", "empty", "the head ref is empty: there is nothing to measure")]
    text = (f"OK    {note or branch + ' matches the pattern'}" if ok
            else f"FAIL  {note}")
    return [("branch-name", "ok" if ok else "fail", text)]


def _gate_issue_spec_local(body, repository, gate):
    findings = []
    for v in gate.validate_body(body, repository, ()):
        findings.append(("issue-spec", "fail", f"issue-spec: {v}"))
    if not findings:
        findings.append(("issue-spec", "ok",
                         "issue-spec local: forma de las referencias correcta."))
    return findings


def _gate_issue_spec_remote(body, repository, fetcher, gate):
    refs = gate.references(body)
    same = [r for r in refs if not r.repo or r.repo.lower() == repository.lower()]
    if not same:
        return []
    # Espejo de ``validate_remote`` (#913): la aprobación se exige a los ``Closes``;
    # las ``Refs`` solo llevan el tipo. Sin ningún ``Closes`` debe haber una aprobada.
    closing = {r.number for r in same if r.kind == "close"}
    findings = []
    approved_linked = False
    for number in sorted({r.number for r in same}):
        try:
            issue = fetcher(number)
        except gate.GitHubApiError as exc:
            return [("issue-spec", "api-error",
                     f"issue-spec: #{number}: cannot read the issue ({exc})")]
        if "pull_request" in issue:
            findings.append(("issue-spec", "ref-is-pr",
                             f"issue-spec: #{number}: la referencia resuelve a "
                             "una pull request, no a una issue; muévala a "
                             "depends-on de Chain Context"))
            continue
        labels = {
            lbl["name"] if isinstance(lbl, dict) else str(lbl)
            for lbl in issue.get("labels", [])
        }
        has_approval = gate.APPROVAL_LABEL in labels
        has_type = any(name.startswith(gate.TYPE_PREFIX) for name in labels)
        if number in closing and not has_approval:
            findings.append(("issue-spec", "missing-approval",
                             f"issue-spec: #{number}: requiere {gate.APPROVAL_LABEL}"))
        if not has_type:
            findings.append(("issue-spec", "missing-type",
                             f"issue-spec: #{number}: requiere una etiqueta "
                             f"{gate.TYPE_PREFIX}*"))
        approved_linked = approved_linked or (has_approval and has_type)
    if not closing and not approved_linked:
        findings.append(("issue-spec", "no-approved-linked",
                         f"issue-spec: {gate.NO_APPROVED_LINKED_MESSAGE}"))
    return findings


def _gate_docs_evidence(body, repository, policy_path, files, gate):
    try:
        policy = gate.load_policy(policy_path)
    except FileNotFoundError as exc:
        return [("docs-evidence", "missing-policy", f"HR-64: {exc}")]
    pr = {"body": body, "labels": []}
    findings = []
    try:
        gate.check_docs_evidence(pr, repository or ".", files, policy, findings)
    except gate.DocsEvidenceConfigError as exc:
        return [("docs-evidence", "config-error", f"HR-64: {exc}")]
    if not findings:
        return [("docs-evidence", "not-declared",
                 "HR-64: docs_evidence no declarado en la politica.")]
    return [("docs-evidence", "fail", line) for line in findings]


def _gate_adoption_phase_five(repo):
    try:
        proc = subprocess.run(  # noqa: S603 — lista de args, shell desactivado.
            [sys.executable, str(CLI_BIN), "adoption", "check", "--phase", "5", repo],
            capture_output=True, text=True, check=False, timeout=ADOPTION_TIMEOUT_S,
            shell=False)
    except (OSError, subprocess.TimeoutExpired) as exc:
        return [("adoption-fase-5", "config-error", f"adoption-fase-5: {exc}")]
    if proc.returncode == 0:
        return [("adoption-fase-5", "ok", "adoption-fase-5: verde")]
    if proc.returncode == ADOPTION_EXIT_FAIL_CLOSED:
        text = (proc.stderr or proc.stdout).strip()
        return [("adoption-fase-5", "config-error", f"adoption-fase-5: {text}")]
    return [("adoption-fase-5", "fail", line.strip())
            for line in (proc.stderr or proc.stdout).splitlines() if line.strip()]


def _read_text(path: Path, label: str) -> str:
    if not path.is_file():
        raise _PreflightInputError(f"{label}: {path}")  # noqa: TRY003 — class helper
    return path.read_text(encoding="utf-8")


def _git_head(repo: str) -> str:
    try:
        out = subprocess.run(  # noqa: S603 — git en PATH, args de control.
            ["git", "-C", repo, "rev-parse", "--abbrev-ref", "HEAD"],  # noqa: S607
            capture_output=True, text=True, check=False, timeout=10,
            shell=False)
    except (OSError, subprocess.TimeoutExpired):
        return ""
    return out.stdout.strip() if out.returncode == 0 else ""


def _collect_branch(findings, branch, pattern, actor):
    """Mueve el gate de rama dentro de ``run_preflight`` para evitar
    un dict de branches→función demasiado complejo (PLR0913)."""
    if branch is None:
        return
    findings.extend(_gate_branch(branch, pattern, actor,
                                 _load_gate("check_branch_name", BRANCH_NAME_GATE)))


def _collect_issue_spec_local(findings, body, repository, issue_spec_gate):
    """Aplica las reglas de forma del cuerpo (rules 1, 5 y 7) sobre el body local."""
    findings.extend(_gate_issue_spec_local(body, repository, issue_spec_gate))


def _collect_issue_spec_remote(findings, body, repository, fetcher, issue_spec_gate):
    if not (fetcher and repository):
        return
    findings.extend(_gate_issue_spec_remote(body, repository, fetcher, issue_spec_gate))


def _collect_docs_evidence(findings, body, repository, policy, files_file):
    if not policy:
        return
    if not files_file:
        raise _PreflightInputError("--policy exige --changed-files")  # noqa: TRY003
    files = [line for line in _read_text(Path(files_file), "changed-files").splitlines()
             if line]
    findings.extend(_gate_docs_evidence(
        body, repository, policy, files,
        _load_gate("check_pr_contract", PR_CONTRACT_GATE)))


def _collect_adoption(findings, with_adoption, repo):
    if not with_adoption:
        return
    if not repo:
        raise _PreflightInputError("--with-adoption exige --repo")  # noqa: TRY003
    findings.extend(_gate_adoption_phase_five(repo))


def _emit(findings, as_json):
    """Formatea la salida y devuelve el código de salida."""
    codes = {code for _, code, _ in findings}
    has_usage = bool(codes & {"api-error", "bad-pattern", "empty",
                              "config-error", "missing-policy"})
    has_findings = bool(codes & {"fail", "missing-approval", "missing-type",
                                 "ref-is-pr", "no-approved-linked"})
    if as_json:
        print(json.dumps({
            "clean": not has_findings,
            "findings": [{"gate": g, "code": c, "detail": d}
                         for g, c, d in findings],
        }, indent=2, ensure_ascii=False))
    else:
        for gate, _code, detail in findings:
            # Prefijo ``[<gate>]`` siempre, también en el caso OK (HR-46:
            # el autor distingue al vuelo qué gate juzgó, sin mirar el código).
            print(f"[{gate}] {detail}", file=sys.stderr)
    if has_usage:
        return EXIT_USAGE
    if has_findings:
        return EXIT_FINDINGS
    return EXIT_OK


def run_preflight(  # noqa: PLR0913 — orquestador de gates; kwargs 1:1 con la CLI.
    *, body_file, repository="", branch=None,
    branch_pattern=DEFAULT_BRANCH_PATTERN, actor="", policy="",
    changed_files_file=None, with_adoption=False, repo="",
    fetch_issue=None, as_json=False,
):
    """Ejecuta el preflight; devuelve el código de salida (0/1/2)."""
    try:
        body = _read_text(Path(body_file), "cuerpo")
    except _PreflightInputError as exc:
        print(f"pr preflight: {exc}", file=sys.stderr)
        return EXIT_USAGE
    issue_spec = _load_gate("check_issue_spec", ISSUE_SPEC_GATE)
    findings = []
    try:
        _collect_branch(findings, branch, branch_pattern, actor)
        _collect_issue_spec_local(findings, body, repository, issue_spec)
        _collect_issue_spec_remote(findings, body, repository, fetch_issue, issue_spec)
        _collect_docs_evidence(findings, body, repository, policy, changed_files_file)
        _collect_adoption(findings, with_adoption, repo)
    except _PreflightInputError as exc:
        print(f"pr preflight: {exc}", file=sys.stderr)
        return EXIT_USAGE
    return _emit(findings, as_json)


def _build_parser():
    p = argparse.ArgumentParser(prog="ci-pattern pr preflight", description=__doc__)
    p.add_argument("--pr-body", required=True)
    p.add_argument("--issue", type=int, default=None)
    p.add_argument("--repository", default="")
    p.add_argument("--branch", default=None)
    p.add_argument("--branch-pattern", default=DEFAULT_BRANCH_PATTERN)
    p.add_argument("--actor", default="")
    p.add_argument("--token", default="")
    p.add_argument("--policy", default="")
    p.add_argument("--changed-files", dest="changed_files_file", default=None)
    p.add_argument("--with-adoption", action="store_true")
    p.add_argument("--repo", default="")
    p.add_argument("--api-url", dest="api_url", default="")
    p.add_argument("--json", dest="as_json", action="store_true")
    return p


def _resolve_defaults(args):
    if args.branch is None and args.repo:
        detected = _git_head(args.repo)
        args.branch = detected or None
    if not args.token:
        args.token = os.environ.get("GITHUB_TOKEN", "")
    if not args.api_url:
        args.api_url = os.environ.get("GITHUB_API_URL", "https://api.github.com")
    return args


def _make_fetcher(args, issue_spec):
    if not (args.token and args.repository):
        return None
    client = issue_spec.GitHubClient(token=args.token, api_url=args.api_url)

    def fetch_issue(number):
        return client.issue(args.repository, number)

    return fetch_issue


def main(argv=None):
    _pin_output_encoding()
    args = _resolve_defaults(_build_parser().parse_args(argv))
    issue_spec = _load_gate("check_issue_spec_main", ISSUE_SPEC_GATE)
    fetch_issue = _make_fetcher(args, issue_spec)
    if args.issue is not None and fetch_issue and args.repository:
        body = _read_text(Path(args.pr_body), "cuerpo")
        existing = {r.number for r in issue_spec.references(body)
                    if not r.repo or r.repo.lower() == args.repository.lower()}
        if args.issue not in existing:
            virtual = fetch_issue(args.issue)

            def _patched(number):
                return virtual if number == args.issue else fetch_issue(number)
            fetch_issue = _patched
    return run_preflight(
        body_file=args.pr_body, repository=args.repository,
        branch=args.branch, branch_pattern=args.branch_pattern, actor=args.actor,
        policy=args.policy, changed_files_file=args.changed_files_file,
        with_adoption=args.with_adoption, repo=args.repo,
        fetch_issue=fetch_issue, as_json=args.as_json)


if __name__ == "__main__":
    raise SystemExit(main())
