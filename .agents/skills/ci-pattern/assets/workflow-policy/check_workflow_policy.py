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
import fnmatch
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
                if not KEY.match(rest):
                    # Bare scalar list item (e.g. "- opened" under
                    # `types:`). The block form
                    #   types:
                    #     - opened
                    #     - synchronize
                    # is semantically equivalent to the flow form
                    #   types: [opened, synchronize]
                    # and the parser must accept both: parse as a
                    # scalar and move to the next line. The previous
                    # behaviour of recursing into parse_node here
                    # raised 'unsupported construct' for any bare
                    # scalar under a list, which is a YAML conformant
                    # form (#BP-857, #BP-951, gate regression noted in
                    # the consumer's PR-B review).
                    seq.append(parse_scalar(rest))
                    i += 1
                    continue
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
        # A job whose provable if excludes pull_request never publishes a PR check name,
        # so its schedule/dispatch runs cannot satisfy one (HR-30).
        never_on_pr = allowed is not None and "pull_request" not in allowed
        for ev in ("schedule", "workflow_dispatch"):
            if ev in events and not qualified and not never_on_pr and (allowed is None or ev in allowed):
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


def load_required(path: str) -> tuple[list[str], dict]:
    """Lee la politica del policy: de `required-jobs` y devuelve (jobs, data)
    para que el resto del gate acceda a otras claves declaradas (HR-39)."""
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise InvalidInput(f"required-jobs policy unreadable: {exc}") from exc
    jobs = data.get("required_jobs") if isinstance(data, dict) else None
    if not isinstance(jobs, list) or not jobs or \
            not all(isinstance(j, str) and j for j in jobs) or len(set(jobs)) != len(jobs):
        raise InvalidInput("required-jobs policy: `required_jobs` must be a non-empty "
                           "list of unique job names")
    return jobs, data


def _state_entry_valid(entry: object) -> str | None:
    """None si valida; mensaje si no. Shape invalido -> exit 2 (HR-3, HR-18)."""
    if not isinstance(entry, dict):
        return "entry is not a mapping"
    str_fields = ("name", "command", "pr_workflow", "pr_job",
                  "scheduled_workflow", "alert_command")
    for f in str_fields + ("manifests",):
        if f not in entry:
            return f"missing required field '{f}'"
        v = entry[f]
        if f == "manifests":
            if not isinstance(v, list) or not v or not all(isinstance(m, str) and m for m in v):
                return "field 'manifests' must be a non-empty list of strings"
        elif not isinstance(v, str) or not v:
            return f"field '{f}' must be a non-empty string"
    return None


def _job_runs_command(job: dict, command: str) -> bool:
    steps = job.get("steps") or []
    for step in steps:
        if not isinstance(step, dict):
            continue
        run = step.get("run")
        if isinstance(run, str) and command in run:
            return True
    return False


def _any_job_runs_command(doc: dict, command: str) -> bool:
    jobs = doc.get("jobs") or {}
    for job in jobs.values():
        if isinstance(job, dict) and _job_runs_command(job, command):
            return True
    return False


def _find_workflow(workflows: list[tuple[str, dict]], target: str):
    """Busca un workflow por igualdad de ruta normalizada. Sin red y sin
    adivinar: la ruta declarada en el policy debe coincidir con una de las
    que el consumer paso por --workflow."""
    t = str(Path(target))
    for path, doc in workflows:
        if str(Path(path)) == t:
            return path, doc
    return None, None


def _paths_cover(manifests: list[str], patterns: list[str]) -> list[str]:
    """Manifiestos declarados que ningun pattern de paths cubre (fnmatch)."""
    uncovered = []
    for m in manifests:
        if not any(fnmatch.fnmatch(m, p) for p in patterns if isinstance(p, str)):
            uncovered.append(m)
    return uncovered


def check_state_dependent_gates(workflows: list[tuple[str, dict]], data: dict,
                                out: list[str]) -> None:
    """HR-39: un gate cuyo veredicto depende de estado externo mutable va
    programado y solo bloquea PRs que tocan los manifiestos declarados.

    Activacion: la clave `state_dependent_gates` debe estar en el policy
    file. Si esta AUSENTE imprime notice por stderr y termina (el fail-open
    lo cierra por fuera el doctor de flota); si esta PRESENTE se exige, y
    la lista vacia es declaracion valida.
    """
    if "state_dependent_gates" not in data:
        print("HR-39: state_dependent_gates no declarado en el policy file; "
              "la regla todavia no se exige (se activa al declararse).",
              file=sys.stderr)
        return
    entries = data["state_dependent_gates"]
    if not isinstance(entries, list):
        raise InvalidInput("`state_dependent_gates` debe ser una lista de entradas")
    for entry in entries:
        err = _state_entry_valid(entry)
        if err is not None:
            raise InvalidInput(f"HR-39: malformed entry ({err})")
        name = entry["name"]
        command = entry["command"]
        pr_wf_path = entry["pr_workflow"]
        pr_job_id = entry["pr_job"]
        manifests = entry["manifests"]
        sched_path = entry["scheduled_workflow"]
        alert = entry["alert_command"]
        pr_path, pr_doc = _find_workflow(workflows, pr_wf_path)
        if pr_doc is None:
            out.append(f"HR-39: el gate '{name}' declara pr_workflow='{pr_wf_path}' "
                       "pero ese workflow no esta entre los inputs.")
            continue
        sched_real_path, sched_doc = _find_workflow(workflows, sched_path)
        if sched_doc is None:
            out.append(f"HR-39: el gate '{name}' declara scheduled_workflow='{sched_path}' "
                       "pero ese workflow no esta entre los inputs.")
            continue
        pr_events = workflow_events(pr_doc)
        if "pull_request" not in pr_events:
            out.append(f"HR-39: el gate '{name}' no se ejecuta en pull_request en "
                       f"{pr_path}; sin disparador de PR no puede gatear el PR.")
            continue
        pr_paths = pr_events["pull_request"].get("paths")
        if not isinstance(pr_paths, list) or not pr_paths:
            out.append(f"HR-39: el gate '{name}' corre en pull_request sin `paths:` "
                       f"en {pr_path}; bloquea PRs que no tocan los manifiestos.")
        else:
            for m in _paths_cover(manifests, pr_paths):
                out.append(f"HR-39: el gate '{name}' no cubre el manifiesto '{m}' "
                           f"en paths: {pr_paths} ({pr_path}).")
        jobs = pr_doc.get("jobs") or {}
        pr_job = jobs.get(pr_job_id)
        if not isinstance(pr_job, dict):
            out.append(f"HR-39: el gate '{name}' declara pr_job='{pr_job_id}' pero "
                       f"ese job no existe en {pr_path}.")
        elif not _job_runs_command(pr_job, command):
            out.append(f"HR-39: el gate '{name}' declara pr_job='{pr_job_id}' pero "
                       f"ese job no ejecuta '{command}' en {pr_path}.")
        sched_events = workflow_events(sched_doc)
        if "schedule" not in sched_events:
            out.append(f"HR-39: el gate '{name}' declara scheduled_workflow='{sched_real_path}' "
                       "pero ese workflow no tiene disparador schedule:.")
        if not _any_job_runs_command(sched_doc, command):
            out.append(f"HR-39: el gate '{name}' corre programado pero ningun job "
                       f"ejecuta '{command}' en {sched_real_path}.")
        if not _any_job_runs_command(sched_doc, alert):
            out.append(f"HR-39: el gate '{name}' corre programado pero no hay paso "
                       f"que ejecute alert_command='{alert}' en {sched_real_path}; "
                       "sin alerta, el rojo programado queda sin abrir issue.")


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main(argv):
    _pin_output_encoding()
    ap = argparse.ArgumentParser(description="Workflow policy gate "
                                "(HR-3, HR-26, HR-30, HR-39).")
    ap.add_argument("--workflow", action="append", required=True)
    ap.add_argument("--required-jobs", required=True)
    ap.add_argument("--base-branch", default="main")
    args = ap.parse_args(argv)
    try:
        required, data = load_required(args.required_jobs)
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
    try:
        check_state_dependent_gates(docs, data, findings)
    except InvalidInput as exc:
        print(f"WORKFLOW-POLICY ERROR: {exc}", file=sys.stderr)
        return 2
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
