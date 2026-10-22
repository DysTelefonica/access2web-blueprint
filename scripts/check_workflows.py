#!/usr/bin/env python3
"""CI gate on .github/workflows/*.yml — issues #135, #141, #152, #552, BP#1040.

BP#1040 moved the structural checks to the canon `workflow-policy` gate
(`.agents/skills/ci-pattern/assets/workflow-policy/`, run by the `workflow policy` step of
ci.yml over every workflow): duplicate YAML keys, `timeout-minutes`, fixed host ports, SHA pins,
the `docker info` preflight and the existence of a `concurrency` group. What stays here is what
the canon does not check, pinned by `tests/test_check_workflows.py`:

 1. A `cancel-in-progress: true` job outside `_CANCEL_SAFE_JOBS` -> WARN.
 2. Workflow-level `permissions.contents` is `read` (or a list that contains
    `read`). Missing `permissions:` block -> WARN — the GITHUB_TOKEN inherits
    a restrictive default, but this is not pinned to `contents: read`.
 3. Every workflow that declares `env.PYTHON_VERSION` agrees on the value.
    The interpreter is part of the gate (Hard Rule #15 of the
    `oracle-vps-github-runners` skill); two declarations with two values
    means one half of CI ships on a Python the test half never saw.
 4. No two jobs in the same workflow share a `concurrency.group` literal.
    `gate-quality` on two jobs means the second job queues behind the
    first — FIFO at a granularity nobody asked for.
 5. A workflow triggered by `pull_request` (or callable via `workflow_call`,
    issue #702 — a reusable workflow's file cannot see whether its caller is
    PR-reachable, so it is treated the same way) may route PR-reachable jobs
    only to literal GitHub-hosted runner labels. Self-hosted and dynamic
    labels are rejected unless a static job condition excludes pull
    requests. A job that calls a reusable workflow (`uses: ./...`) has no
    `runs-on` of its own and is exempt; the called workflow's jobs are
    checked when that file is scanned.

Hard Rule 1 of the deterministic-quality-harness skill: no `|| true`, no
`continue-on-error`. Exit 0 = clean (warnings allowed), exit 1 = any
violation. A WARN is not a violation; it is a printed line that the next
operator gets to act on.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import yaml

# --------------------------------------------------------------------------------------------
# CONFIGURATION
# --------------------------------------------------------------------------------------------
WORKFLOWS_DIR = Path(__file__).resolve().parent.parent / ".github" / "workflows"

#: Repo-approved allowlist of jobs that may safely use `cancel-in-progress: true`.
#:
#: Empty by design (audit 2026-08-12, umbrella #117). Every gate that consumed
#: a runner slot already paid the cost. Adding a job here is a
#: deliberate decision that the in-flight work is cheap to discard — not a
#: stylistic preference. Tests patch this set via `monkeypatch.setattr` to
#: exercise the silent branch without modifying the source.
_CANCEL_SAFE_JOBS: set[str] = set()


# --------------------------------------------------------------------------------------------
# CHECKS (need a parsed dict)
# --------------------------------------------------------------------------------------------
# Convention: each check yields `Finding = (severity, location, message)`
# where severity is "error" (violation -> exit 1) or "warn" (printed, exit 0).
Finding = tuple[str, str, str]


def _check_concurrency(doc: dict[str, Any]) -> Iterator[Finding]:
    """Yield a warning for jobs that cancel in-flight runs without a repo decision.

    Audit 2026-08-12 (umbrella #117): `cancel-in-progress: true` on a job that is not in
    `_CANCEL_SAFE_JOBS` is a WARN (cancelling discards work that already paid for a runner slot);
    `false` or absent is silent, and so is an allowlisted job. A job with no concurrency group at
    all is the canon workflow-policy gate's finding (BP#1040), not this one's.

    Tests exercise the allowlist via `monkeypatch.setattr(check_workflows,
    "_CANCEL_SAFE_JOBS", {job_name})`; the production constant is empty.
    """
    jobs = doc.get("jobs") or {}
    if not isinstance(jobs, dict):
        return
    top_level = doc.get("concurrency")
    for job_name, job_def in jobs.items():
        if not isinstance(job_def, dict):
            continue
        concurrency = job_def.get("concurrency") or top_level
        # A missing group is the canon workflow-policy gate's finding (BP#1040); only the
        # cancel-in-progress warning is local.
        if not isinstance(concurrency, dict) or not concurrency.get("group"):
            continue
        if concurrency.get("cancel-in-progress") is True and job_name not in _CANCEL_SAFE_JOBS:
            yield (
                "warn",
                f"jobs.{job_name}",
                f"job `{job_name}` uses `cancel-in-progress: true`; "
                f"cancelling discards work that already consumed a runner "
                f"slot. Add `{job_name}` to _CANCEL_SAFE_JOBS in "
                f"scripts/check_workflows.py to silence this warning.",
            )


def _check_permissions(doc: dict[str, Any]) -> Iterator[Finding]:
    """Yield a finding for the workflow-level `permissions` block.

    Audit 2026-08-12 (oracle-vps-github-runners v2.0, Hard Rule #6): the
    GITHUB_TOKEN must be pinned to `contents: read` (or a list that contains
    `read`). The default is restrictive but not pinned, so a missing block is
    a WARN — the workflow happens to inherit a restrictive default, but the
    next operator to add a job that requests `contents: write` will not get a
    gate catching the privilege escalation.

      - no `permissions:` block           -> WARN  (default is restrictive but
                                                    not pinned to `read`)
      - `permissions.contents: read`      -> silent (the minimum this gate
                                                    enforces)
      - `permissions.contents: [read, …]` -> silent (superset accepted; the
                                                    runner-skill rule is
                                                    "no less than read")
      - `permissions.contents: write`     -> ERROR
      - shorthand `permissions: read-all`
        or `permissions: write-all`       -> ERROR (grants more than the
                                                    gate's minimum)
      - non-mapping permissions           -> ERROR (unparseable for this gate)
    """
    permissions = doc.get("permissions")
    if permissions is None:
        yield (
            "warn",
            "permissions",
            "workflow declares no `permissions:` block; the GITHUB_TOKEN "
            "inherits a restrictive default, but this is not pinned to "
            "`contents: read` — add a top-level `permissions: { contents: read }`",
        )
        return
    if not isinstance(permissions, dict):
        yield (
            "error",
            "permissions",
            f"workflow uses shorthand `permissions: {permissions}`; "
            f"replace with `permissions: {{ contents: read }}` so the "
            f"GITHUB_TOKEN is pinned to the minimum-readable scope",
        )
        return
    contents = permissions.get("contents")
    if contents == "read":
        return
    if isinstance(contents, list) and "read" in contents:
        return
    yield (
        "error",
        "permissions.contents",
        f"workflow `permissions.contents` is `{contents!r}`; must be "
        f"`read` or a list that contains `read`",
    )


def _check_concurrency_group_uniqueness(doc: dict[str, Any]) -> Iterator[Finding]:
    """Yield a violation when two jobs in the same workflow share a group.

    Audit 2026-08-12 (oracle-vps-github-runners v2.0, Hard Rule #9): two
    jobs in the same workflow sharing a `concurrency.group` means a new
    push queues behind the in-flight run of the OTHER job — so the second
    job does not actually start until the first finishes. That is FIFO at
    a granularity nobody asked for, and it silently serialises the gate.

    The check is per-workflow: the same group literal in two different
    workflows is fine, because the runners and the FIFO queues are
    different. (security.yml::gitleaks and ci.yml::quality are not on the
    same machine, and even if they were, the cross-workflow collision
    would not serialise either one.)
    """
    jobs = doc.get("jobs") or {}
    if not isinstance(jobs, dict):
        return
    top_level = doc.get("concurrency")
    groups: dict[str, list[str]] = {}
    for job_name, job_def in jobs.items():
        if not isinstance(job_def, dict):
            continue
        concurrency = job_def.get("concurrency") or top_level
        if not isinstance(concurrency, dict):
            continue
        group = concurrency.get("group")
        if not isinstance(group, str) or not group:
            continue
        groups.setdefault(group, []).append(job_name)
    for group, names in groups.items():
        if len(names) > 1:
            yield (
                "error",
                "concurrency.group",
                f"jobs {names} share `concurrency.group: {group!r}`; "
                f"each job needs a distinct group so they do not queue "
                f"behind each other on the same workflow run",
            )


# --------------------------------------------------------------------------------------------
# CROSS-WORKFLOW CHECK (check 8)
# --------------------------------------------------------------------------------------------


def _check_python_version_consistency(
    parsed: list[tuple[Path, dict[str, Any]]],
) -> Iterator[Finding]:
    """Yield a violation when workflows disagree on `env.PYTHON_VERSION`.

    Audit 2026-08-12 (oracle-vps-github-runners v2.0, Hard Rule #15): the
    interpreter is part of the gate. A patch bump is a different verdict.
    Two workflows declaring two different versions means the test suite
    is exercised on one and shipped on the other — the "passes locally"
    half of a bad CI is exactly that.

    Only workflows that DECLARE `env.PYTHON_VERSION` participate. A
    workflow that does not declare it does not contribute to the check,
    since it does not run Python (security.yml and security-deep.yml
    ship Docker images; they do not need setup-python at the workflow
    level). The check fires only when two declarations disagree.
    """
    declared: list[tuple[Path, str]] = []
    for path, doc in parsed:
        if not isinstance(doc, dict):
            continue
        env = doc.get("env")
        if not isinstance(env, dict):
            continue
        pv = env.get("PYTHON_VERSION")
        if isinstance(pv, str):
            declared.append((path, pv))
    by_value: dict[str, list[Path]] = {}
    for path, value in declared:
        by_value.setdefault(value, []).append(path)
    if len(by_value) > 1:
        # One finding that names every (workflow, value) pair so a
        # reviewer can fix the disagreement without rerunning the gate.
        pairs = ", ".join(
            f"`{p.name}` declares `{v!r}`" for v, paths in sorted(by_value.items()) for p in paths
        )
        yield (
            "error",
            "env.PYTHON_VERSION",
            f"workflows disagree on `env.PYTHON_VERSION`: {pairs}",
        )


# --------------------------------------------------------------------------------------------
# PUBLIC-PR TRUST BOUNDARY (check 11)
# --------------------------------------------------------------------------------------------


_HOSTED_RUNNER = re.compile(r"^(?:ubuntu|windows|macos)-[A-Za-z0-9._-]+$")
_NON_PR_EVENT = re.compile(r"^github\.event_name\s*==\s*(['\"])(?!pull_request\1)[A-Za-z0-9_]+\1$")


def _is_pr_reachable_workflow(doc: dict[str, Any]) -> bool:
    """Return whether the workflow can execute as part of a public PR run.

    PyYAML 1.1 parses the plain key `on` as boolean ``True``. Supporting both
    representations keeps the gate correct without replacing the repository's
    YAML parser.

    Issue #702: a workflow triggered only by `workflow_call` (a reusable
    workflow, e.g. `security.yml` / `codeql.yml` after being embedded into
    `ci.yml`) declares no `pull_request` trigger of its own, yet its jobs
    still run inside a public PR whenever its caller is PR-reachable — and a
    reusable workflow's file cannot see who calls it. Fail closed: treat
    `workflow_call` the same as `pull_request` for this check, since the
    opposite assumption (an uncalled, therefore safe, reusable workflow) is
    exactly the gap a self-hosted runner slipping into a PR path would need.
    """
    triggers = doc.get("on", doc.get(True))
    if isinstance(triggers, str):
        return triggers in ("pull_request", "workflow_call")
    if isinstance(triggers, list):
        return "pull_request" in triggers or "workflow_call" in triggers
    return isinstance(triggers, dict) and (
        "pull_request" in triggers or "workflow_call" in triggers
    )


def _condition_excludes_pull_requests(condition: object) -> bool:
    """Accept only simple event-name conditions that cannot match a PR."""
    if not isinstance(condition, str):
        return False
    expression = condition.strip()
    if expression.startswith("${{") and expression.endswith("}}"):
        expression = expression[3:-2].strip()
    if re.fullmatch(r"github\.event_name\s*!=\s*(['\"])pull_request\1", expression):
        return True
    terms = [term.strip() for term in expression.split("||")]
    return bool(terms) and all(_NON_PR_EVENT.fullmatch(term) for term in terms)


def _check_public_pr_runner_isolation(doc: dict[str, Any]) -> Iterator[Finding]:
    """Reject PR-reachable jobs that do not use a literal hosted runner.

    Issue #702: a job that calls a reusable workflow (`uses: ./...`) has no
    `runs-on` at all — GitHub's schema for that job shape does not accept
    it — so it is exempt here. Isolation for the reusable workflow's actual
    steps is enforced when that file is scanned on its own, via
    `_is_pr_reachable_workflow` now also treating `workflow_call` as
    PR-reachable.
    """
    if not _is_pr_reachable_workflow(doc):
        return
    jobs = doc.get("jobs") or {}
    if not isinstance(jobs, dict):
        return
    for job_name, job_def in jobs.items():
        if not isinstance(job_def, dict) or _condition_excludes_pull_requests(job_def.get("if")):
            continue
        if "uses" in job_def:
            continue
        runner = job_def.get("runs-on")
        if isinstance(runner, str) and _HOSTED_RUNNER.fullmatch(runner):
            continue
        yield (
            "error",
            f"jobs.{job_name}.runs-on",
            f"PR-reachable job `{job_name}` uses `{runner!r}`; public pull requests "
            f"must run on a literal GitHub-hosted label such as `ubuntu-24.04`",
        )


#: Etiqueta del runner Windows del propietario para extraer ERD de solo lectura
#: (#1074). Es la única exención a la regla de #800 y exige un workflow cuyo
#: único disparador sea `workflow_dispatch`.
_DISPATCH_ONLY_RUNNER_LABEL = "dysflow-erd"


def _is_dispatch_only_workflow(doc: dict[str, Any]) -> bool:
    triggers = doc.get("on", doc.get(True))
    if isinstance(triggers, str):
        return triggers == "workflow_dispatch"
    if isinstance(triggers, list):
        return triggers == ["workflow_dispatch"]
    return isinstance(triggers, dict) and list(triggers) == ["workflow_dispatch"]


def _check_no_self_hosted_runners(doc: dict[str, Any]) -> Iterator[Finding]:
    """Reject ANY job pinned to a self-hosted runner (issue #800).

    This is the repository rule, not a PR-isolation nuance: the repository is
    public and does not consume the operator's own VPS runners at all
    (`docs/10-runners.md`). `_check_public_pr_runner_isolation` deliberately
    exempts jobs excluded from pull requests, so a scheduled job could hide
    there — `mutation` did: pinned to a label with zero registered runners, red
    every week while the `required` aggregator published green (#785).
    """
    jobs = doc.get("jobs") or {}
    if not isinstance(jobs, dict):
        return
    for job_name, job_def in jobs.items():
        if not isinstance(job_def, dict) or "uses" in job_def:
            continue
        runner = job_def.get("runs-on")
        labels = runner if isinstance(runner, list) else [runner]
        if not any(str(label) == "self-hosted" for label in labels):
            continue
        if _DISPATCH_ONLY_RUNNER_LABEL in map(str, labels) and _is_dispatch_only_workflow(doc):
            continue
        yield (
            "error",
            f"jobs.{job_name}.runs-on",
            f"job `{job_name}` uses `{runner!r}`; a public repository does not consume "
            f"the operator's self-hosted runners (docs/10-runners.md): use a literal "
            f"GitHub-hosted label such as `ubuntu-24.04` or `ubuntu-24.04-arm`",
        )


# --------------------------------------------------------------------------------------------
# RUN-ALL-CHECKS
# --------------------------------------------------------------------------------------------


def _format_finding(path: Path, severity: str, check_name: str, location: str, message: str) -> str:
    """Render one finding to the on-disk `path:line: CHECK-NAME: location: message` shape.

    WARN findings add a `WARN: ` prefix so the existing error-line tests
    continue to find their substrings (the check name and message are
    identical to what an ERROR would have printed) and the new warn-line
    tests can grep for `WARN`. The prefix sits between the line number
    and the check name so the format stays line-anchored.
    """
    prefix = "WARN: " if severity == "warn" else ""
    return f"{path}:1: {prefix}{check_name}: {location}: {message}"


def _check_one(path: Path, doc: dict[str, Any]) -> Iterator[tuple[str, str]]:
    """Yield `(severity, formatted_line)` for every per-workflow finding on `path`.

    `doc` is the already-parsed workflow. The caller is responsible for
    reading the file and running the duplicate-key and parse-error checks
    (which short-circuit before any per-workflow check can run on a broken
    file — see `_check_one_with_parse`).
    """
    for severity, location, message in _check_concurrency(doc):
        yield (
            severity,
            _format_finding(path, severity, "concurrency", location, message),
        )
    for severity, location, message in _check_permissions(doc):
        yield (
            severity,
            _format_finding(path, severity, "permissions", location, message),
        )
    for severity, location, message in _check_concurrency_group_uniqueness(doc):
        yield (
            severity,
            _format_finding(path, severity, "concurrency-group", location, message),
        )
    for severity, location, message in _check_public_pr_runner_isolation(doc):
        yield (
            severity,
            _format_finding(path, severity, "public-pr-runner", location, message),
        )
    for severity, location, message in _check_no_self_hosted_runners(doc):
        yield (
            severity,
            _format_finding(path, severity, "no-self-hosted-runner", location, message),
        )


def _check_one_with_parse(
    path: Path,
) -> tuple[dict[str, Any] | None, list[tuple[str, str]]]:
    """Read `path`, run parse-time checks, return `(doc, findings)`.

    `doc` is the parsed workflow, or None if the file failed to parse or
    had duplicate keys. `findings` is a list of `(severity, formatted_line)`
    pairs for this path's parse-time findings (the per-workflow findings
    are emitted separately by `_check_one` once the caller has the doc).

    Splitting the parse from the per-workflow checks is what lets
    `main()` collect the parsed `doc` for the cross-workflow check
    without re-reading any file.
    """
    findings: list[tuple[str, str]] = []

    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        findings.append(("error", f"{path}:0: read-error: {exc}"))
        return None, findings

    try:
        doc = yaml.safe_load(text)
    except yaml.YAMLError as exc:
        findings.append(("error", f"{path}:0: parse-error: {exc}"))
        return None, findings

    if not isinstance(doc, dict):
        findings.append(
            (
                "error",
                f"{path}:0: shape-error: top-level YAML must be a mapping "
                f"(GitHub Actions workflow)",
            )
        )
        return None, findings

    return doc, findings


def _iter_workflows(root: Path) -> Iterator[Path]:
    if not root.is_dir():
        return
    for path in sorted(root.glob("*.yml")):
        if path.is_file():
            yield path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--root",
        default=str(WORKFLOWS_DIR),
        help="directory containing the workflow files (default: .github/workflows)",
    )
    args = parser.parse_args(argv)

    root = Path(args.root)
    if not root.is_dir():
        print(f"FAIL  workflows directory does not exist: {root}", file=sys.stderr)
        return 1

    violations = 0
    warnings = 0
    parsed: list[tuple[Path, dict[str, Any]]] = []

    for path in _iter_workflows(root):
        doc, parse_findings = _check_one_with_parse(path)
        for severity, line in parse_findings:
            print(line, file=sys.stderr)
            if severity == "warn":
                warnings += 1
            else:
                violations += 1
        if doc is not None:
            for severity, line in _check_one(path, doc):
                print(line, file=sys.stderr)
                if severity == "warn":
                    warnings += 1
                else:
                    violations += 1
            parsed.append((path, doc))

    # Cross-workflow checks (env.PYTHON_VERSION consistency) only see
    # workflows that parsed cleanly. A parse-broken file already
    # produced an error above; an absent doc is not a reason to skip the
    # cross-workflow check on the rest.
    if len(parsed) >= 2:
        for severity, location, message in _check_python_version_consistency(parsed):
            # All participating workflows are named in the message; the
            # `path` here is just the surface the on-disk line is
            # anchored to (the line number is meaningless for a
            # cross-workflow finding). Use the first parsed path as
            # the anchor.
            anchor = parsed[0][0]
            line = _format_finding(
                anchor, severity, "python-version-consistency", location, message
            )
            print(line, file=sys.stderr)
            if severity == "warn":
                warnings += 1
            else:
                violations += 1

    if violations == 0:
        files = sum(1 for _ in _iter_workflows(root))
        if warnings:
            print(f"OK    {files} workflow file(s) clean ({warnings} warning(s))")
        else:
            print(f"OK    {files} workflow file(s) clean")
        return 0

    print(
        f"FAIL  {violations} violation(s) across workflow files",
        file=sys.stderr,
    )
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
