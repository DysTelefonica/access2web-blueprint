#!/usr/bin/env python3
# ci-pattern asset — workflow policy gate: structural reader + HR-26 (DysTelefonica/team-skills#192)
"""Structural policy gate over GitHub Actions workflows (HR-3, HR-26, HR-30).

Reads each workflow with a deliberately small YAML subset, stdlib only:
block mappings and sequences, plain or quoted scalars, block scalars
(``|``/``>``), one-line flow sequences and comments. Anchors, aliases,
merge keys, tags, flow mappings and tab indentation fail closed (exit 2)
instead of being guessed: the tree is walked structurally, never by line
shapes.

HR-26 (pinned toolchain), on every job, required or not: every ``uses:``
action is pinned to a full 40-hex commit SHA, and ``docker://`` images
by digest (``@sha256:`` + 64 hex; local ``./`` actions are exempt);
``actions/setup-python`` / ``actions/setup-node`` declare an exact
``X.Y.Z`` version; ``pip install`` declares ``==`` inline, or a
requirements file (``-r``) with ``--require-hashes``. Constraint files
(``-c``) are always a finding: the pins inside are invisible to a
structural gate, and reporting OK without measuring them would be the
failure HR-3 forbids — declare the pins inline or in a hash-pinned
requirements file instead.

HR-3 (fail loud), on required jobs: no ``continue-on-error: true``
(workflow, job or step level; a conditional expression is a finding too)
and no masking statement in run blocks — ``|| true``, ``|| :``,
``|| exit 0``, ``set +e``, ``set +o errexit`` or a bare ``exit 0``.

HR-30 (required names from PR events only), on required jobs: the
workflow triggers on ``pull_request``; ``schedule`` and
``workflow_dispatch`` do not publish the required check name — the job
either skips them via a job-level ``if`` whose ``event_name ==
'<event>'`` comparisons prove it, or its name is event-qualified
(``name: x (${{ github.event_name }})``); ``push`` must be filtered to
the base branch. An ``if`` containing ``!=``, ``always()``, ``${{`` or
``needs.`` is not provable and is treated as running on every event.

The list of required jobs always comes from the required-jobs policy,
never from code. A required job named by the policy but absent from every
analyzed workflow is a finding: silence there would be a gate reporting
OK without having measured.

Exit codes: 0 no findings · 1 findings · 2 invalid input or empty subject.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import NamedTuple

SHA40 = re.compile(r"^[0-9a-f]{40}$")
DOCKER_DIGEST = re.compile(r"^sha256:[0-9a-f]{64}$")
EXACT_VERSION = re.compile(r"^\d+\.\d+\.\d+$")
ALIAS = re.compile(r"^[\w.-]+$")
KEY = re.compile(r"""^('[^']*'|"[^"]*"|[^:\s][^:]*?):(?:\s+(.*))?$""")
PIP_INSTALL = re.compile(r"\bpip3? install\b")
PIP_REQUIREMENTS = re.compile(r"\s(-r|--requirement)(\s|=)")
PIP_CONSTRAINTS = re.compile(r"\s(-c|--constraint)(\s|=)")
EVENT_IF = re.compile(r"event_name\s*==\s*['\"]([^'\"]+)['\"]")
MASK_TOKENS = ("|| true", "|| :", "|| exit 0")
NORUN_TOKENS = ("set +e", "set +o errexit")
SETUP_KEYS = {"actions/setup-python": "python-version", "actions/setup-node": "node-version"}


class InvalidInput(Exception):
    """A condition the gate refuses to interpret; maps to exit 2."""


class Line(NamedTuple):
    number: int
    indent: int
    text: str


def strip_comment(text: str) -> str:
    quote = None
    for i, c in enumerate(text):
        if quote:
            quote = None if c == quote else quote
        elif c in "'\"":
            quote = c
        elif c == "#" and (i == 0 or text[i - 1] in " \t"):
            return text[:i].rstrip()
    return text.rstrip()


def read_lines(text: str) -> list[Line]:
    lines = []
    for n, raw in enumerate(text.splitlines(), 1):
        content = strip_comment(raw.strip())
        if not content:
            continue
        if "\t" in raw[:len(raw) - len(raw.lstrip())]:
            raise InvalidInput(f"line {n}: tab in indentation")
        lines.append(Line(n, len(raw) - len(raw.lstrip()), content))
    return lines


def parse_scalar(v: str) -> object:
    if v[0] in "&!{" or (v[0] == "*" and ALIAS.fullmatch(v)):
        raise InvalidInput(f"unsupported YAML construct: {v!r}")
    if len(v) >= 2 and v[0] == v[-1] and v[0] in "'\"":
        if v[0] in v[1:-1]:
            raise InvalidInput(f"unterminated quoted scalar: {v!r}")
        return v[1:-1]
    return {"true": True, "false": False}.get(v, v)


def parse_flow_seq(text: str, n: int) -> list:
    if not (text.startswith("[") and text.endswith("]")):
        raise InvalidInput(f"line {n}: unsupported value: {text!r}")
    inner = text[1:-1].strip()
    if any(c in inner for c in "[{"):
        raise InvalidInput(f"line {n}: nested flow collection")
    return [parse_scalar(x.strip()) for x in inner.split(",")] if inner else []


def parse_node(lines: list[Line], i: int, indent: int):
    line = lines[i]
    if line.text == "-" or line.text.startswith("- "):
        seq = []
        while i < len(lines) and lines[i].indent == indent and \
                (lines[i].text == "-" or lines[i].text.startswith("- ")):
            rest = lines[i].text[1:].strip()
            n = lines[i].number
            if rest:
                lines[i] = Line(n, indent + 2, rest)
                val, i = parse_node(lines, i, indent + 2)
            else:
                i += 1
                val, i = parse_node(lines, i, lines[i].indent) \
                    if i < len(lines) and lines[i].indent > indent else (None, i)
            seq.append(val)
        return seq, i
    if not KEY.match(line.text):
        raise InvalidInput(f"line {line.number}: unsupported construct: {line.text!r}")
    out = {}
    while i < len(lines) and lines[i].indent == indent:
        m = KEY.match(lines[i].text)
        if not m:
            raise InvalidInput(f"line {lines[i].number}: unsupported construct: {lines[i].text!r}")
        key = parse_scalar(m.group(1))
        if key == "<<":
            raise InvalidInput(f"line {lines[i].number}: merge key")
        rest = m.group(2)
        n = lines[i].number
        i += 1
        if rest in (None, ""):
            if i < len(lines) and lines[i].indent > indent:
                out[key], i = parse_node(lines, i, lines[i].indent)
            else:
                out[key] = None
        elif rest[0] in "|>":
            while i < len(lines) and lines[i].indent > indent:
                out[key] = out.get(key, "") + lines[i].text + "\n"
                i += 1
        elif rest[0] == "[":
            out[key] = parse_flow_seq(rest, n)
        else:
            out[key] = parse_scalar(rest)
    return out, i


def parse_workflow(text: str, path: str) -> dict:
    lines = read_lines(text)
    if not lines:
        raise InvalidInput(f"{path}: empty workflow")
    doc, i = parse_node(lines, 0, lines[0].indent)
    if i < len(lines):
        raise InvalidInput(f"{path}: stray content at line {lines[i].number}")
    jobs = doc.get("jobs") if isinstance(doc, dict) else None
    if not isinstance(jobs, dict) or not jobs:
        raise InvalidInput(f"{path}: empty subject (no jobs block)")
    for jid, job in jobs.items():
        if not isinstance(job, dict):
            raise InvalidInput(f"{path}: job '{jid}' is not a mapping")
    return doc


def workflow_events(doc: dict) -> dict:
    raw = doc.get("on", doc.get(".on"))
    if raw is None:
        return {}
    if isinstance(raw, list):
        return {str(e): {} for e in raw}
    if isinstance(raw, dict):
        return {k: (v if isinstance(v, dict) else {}) for k, v in raw.items()}
    raise InvalidInput("unsupported 'on:' value")


def allowed_events(job: dict):
    """Events provably allowed by the job-level if; None when not provable."""
    cond = job.get("if")
    if not isinstance(cond, str) or not cond:
        return None
    if "!=" in cond or "always()" in cond or "${{" in cond or "needs." in cond:
        return None
    return set(EVENT_IF.findall(cond))


def scan_run(text: str) -> list[str]:
    found = []
    for raw in text.splitlines():
        line = raw.strip()
        for tok in MASK_TOKENS + NORUN_TOKENS:
            if tok in line and f"`{tok}`" not in found:
                found.append(f"`{tok}` masks failure (fail loud, HR-3)")
        if re.fullmatch(r"exit 0\s*;?", line) or line.endswith(("; exit 0", "; exit 0;")):
            found.append("bare `exit 0` masks failure (fail loud, HR-3)")
    return found


def pin_finding(uses: str) -> str | None:
    if uses.startswith(("./", "../")):
        return None
    if uses.startswith("docker://"):
        ref = uses.rsplit("@", 1)[1] if "@" in uses else ""
        if not DOCKER_DIGEST.fullmatch(ref):
            return f"`{uses}` is not pinned by digest `@sha256:<64 hex>` (HR-26)"
        return None
    ref = uses.rsplit("@", 1)[1] if "@" in uses else ""
    if not SHA40.fullmatch(ref):
        return f"`{uses}` is not pinned to a full 40-char commit SHA (HR-26)"
    return None


def check_job(wf: str, jid: str, job: dict, doc: dict, required: bool, base: str, out: list):
    def finding(msg):
        out.append(f"{wf} job '{jid}': {msg}")

    if required:
        if doc.get("continue-on-error") is True:
            finding("workflow-level `continue-on-error: true` masks failure (HR-3)")
        if job.get("continue-on-error") is True or isinstance(job.get("continue-on-error"), str):
            finding("`continue-on-error` masks failure (HR-3)")
        events = workflow_events(doc)
        if "pull_request" not in events:
            finding("required check name is never published from a PR event (HR-30)")
        allowed = allowed_events(job)
        qualified = isinstance(job.get("name"), str) and "github.event_name" in job["name"]
        for ev in ("schedule", "workflow_dispatch"):
            if ev in events and not qualified and (allowed is None or ev in allowed):
                finding(f"publishes its check name on '{ev}', which does not evaluate the PR (HR-30)")
        if "push" in events:
            branches = events["push"].get("branches")
            base_only = isinstance(branches, list) and branches and set(branches) == {base}
            if not qualified and not base_only and (allowed is None or "push" in allowed):
                finding("`push` trigger reaches branches other than the base (HR-30)")

    steps = job.get("steps") or []
    for st in steps:
        if not isinstance(st, dict):
            raise InvalidInput(f"{wf}: step of job '{jid}' is not a mapping")
        uses = st.get("uses")
        if isinstance(uses, str):
            msg = pin_finding(uses)
            if msg:
                finding(msg)
            for repo, key in SETUP_KEYS.items():
                if uses.startswith(repo + "@"):
                    with_block = st.get("with")
                    version = with_block.get(key) if isinstance(with_block, dict) else None
                    if not isinstance(version, str) or not EXACT_VERSION.fullmatch(version):
                        finding(f"`{repo}` does not declare an exact {key} X.Y.Z (HR-26)")
        if required and (st.get("continue-on-error") is True or
                         isinstance(st.get("continue-on-error"), str)):
            finding("step-level `continue-on-error` masks failure (HR-3)")
        run = st.get("run")
        if isinstance(run, str):
            if required:
                for msg in scan_run(run):
                    finding(msg)
            for raw in run.splitlines():
                if PIP_INSTALL.search(raw) and "==" not in raw:
                    if PIP_REQUIREMENTS.search(raw):
                        if "--require-hashes" not in raw:
                            finding("requirements-file install without `--require-hashes` "
                                    "cannot be verified as exact (HR-26)")
                    elif PIP_CONSTRAINTS.search(raw):
                        finding("constraint-file install cannot be verified as exact; declare "
                                "pins inline or in a hash-pinned requirements file (HR-26)")
                    else:
                        finding("`pip install` without an exact `==` pin (HR-26)")


def load_required(path: str) -> list[str]:
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise InvalidInput(f"required-jobs policy unreadable: {exc}") from exc
    jobs = data.get("required_jobs") if isinstance(data, dict) else None
    if not isinstance(jobs, list) or not jobs or \
            not all(isinstance(j, str) and j for j in jobs) or len(set(jobs)) != len(jobs):
        raise InvalidInput("required-jobs policy: `required_jobs` must be a non-empty "
                           "list of unique job names")
    return jobs


def main(argv):
    ap = argparse.ArgumentParser(description="Workflow policy gate (HR-3, HR-26, HR-30).")
    ap.add_argument("--workflow", action="append", required=True)
    ap.add_argument("--required-jobs", required=True)
    ap.add_argument("--base-branch", default="main")
    args = ap.parse_args(argv)
    try:
        required = load_required(args.required_jobs)
        docs = [(w, parse_workflow(Path(w).read_text(encoding="utf-8"), w)) for w in args.workflow]
    except (InvalidInput, OSError) as exc:
        print(f"WORKFLOW-POLICY ERROR: {exc}", file=sys.stderr)
        return 2
    findings, seen = [], set()
    for w, doc in docs:
        for jid, job in doc["jobs"].items():
            if jid in required:
                seen.add(jid)
            check_job(w, jid, job, doc, jid in required, args.base_branch, findings)
    for jid in required:
        if jid not in seen:
            findings.append(f"required job '{jid}' is absent from every analyzed workflow")
    if findings:
        print(f"WORKFLOW-POLICY FAIL: {len(findings)} finding(s)")
        for f in findings:
            print(f"  - {f}")
        return 1
    print(f"WORKFLOW-POLICY OK: {len(docs)} workflow(s), {len(seen)} required job(s) "
          "checked; no findings.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
