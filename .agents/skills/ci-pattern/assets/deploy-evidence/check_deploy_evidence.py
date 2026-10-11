#!/usr/bin/env python3
# ci-pattern asset — deploy-evidence gate (HR-10; DysTelefonica/team-skills#253)
"""Deploy/battery evidence gate (HR-10).

The policy declares which workflow jobs publish deploy or battery
evidence. For each one, the gate reads the job's run scripts structurally
and requires the evidence to be published as a commit status or check run
OVER THE DEPLOYED SHA: the evidence call must carry an explicit SHA
expression (github.sha, GITHUB_SHA, needs.*.outputs.*sha*,
github.event.*.sha). HR-10 violations:

  - evidence published over the default branch or a plain ref
    (``github.ref``, ``--ref <rama>``, forbidden branch names);
  - the SHA hanging from a workflow-level global env variable;
  - a declared evidence job that publishes no commit evidence at all.

Empty ``evidence_jobs`` is a DECLARED state (a repo without deploys) and
exits 0 with an explicit line — not a silent vacuous pass (HR-18). Exit
codes: 0 clean/declared-empty, 1 findings, 2 invalid input. stdlib only,
read-only, no network.
"""
import argparse
import json
import re
import sys
from pathlib import Path

SHA_EXPRESSIONS = ("github.sha", "GITHUB_SHA", "github.event")
EVIDENCE_MARKERS = ("statuses/", "check-runs", "head_sha")
DEFAULT_BRANCH_RE = re.compile(r"github\.ref\b|--ref\b")
GLOBAL_ENV_USE_RE = re.compile(r"\$\{?\{?env\.([A-Za-z_][A-Za-z0-9_]*)|\$([A-Z][A-Z0-9_]{2,})\b")


def fail2(msg):
    print(f"fail-closed: {msg}", file=sys.stderr)
    return 2


def _job_block(workflow_text, job):
    """Extract one top-level job block (indentation-based), or None."""
    lines = workflow_text.splitlines()
    starts = [i for i, l in enumerate(lines)
              if re.match(rf"^  {re.escape(job)}:\s*(#.*)?$", l)]
    if len(starts) != 1:
        return None
    start = starts[0]
    end = len(lines)
    for i in range(start + 1, len(lines)):
        # el bloque acaba en el siguiente job hermano (misma indentación)
        if re.match(r"^  [^\s#].*:\s*(#.*)?$", lines[i]):
            end = i
            break
    return "\n".join(lines[start:end])


def _workflow_env_keys(workflow_text):
    """Keys declared in the workflow-level (column-0) env: block."""
    keys = set()
    in_env = False
    for line in workflow_text.splitlines():
        if re.match(r"^env:\s*$", line):
            in_env = True
            continue
        if in_env:
            if line and not line[0].isspace():
                in_env = False
                continue
            m = re.match(r"^\s+([A-Za-z_][A-Za-z0-9_]*):", line)
            if m:
                keys.add(m.group(1))
    return keys


def _sha_position_tokens(line):
    """Tokens que ocupan la posición del SHA tras statuses/, check-runs/ o
    en head_sha=."""
    tokens = []
    for marker in ("statuses/", "check-runs/", "head_sha="):
        idx = 0
        while True:
            i = line.find(marker, idx)
            if i < 0:
                break
            rest = line[i + len(marker):]
            token = re.split(r"[\s'\"\\]", rest)[0]
            if token:
                tokens.append(token)
            idx = i + len(marker)
    return tokens


def _global_sha_var(token, global_env_keys):
    """El token de la posición del SHA es una variable global del workflow."""
    clean = token.strip("${}$ ")
    return clean in global_env_keys


_NEEDS_SHA_RE = re.compile(r"needs\.[A-Za-z0-9_-]+\.outputs\.[A-Za-z0-9_-]*sha", re.IGNORECASE)


def _has_sha_expression(line):
    """El nombre de la salida debe contener sha (needs.build.outputs.sha ✓,
    needs.build.outputs.branch ✗) — igual que el docstring promete."""
    low = line.lower()
    return ("github.sha" in low or "github_sha" in low
            or _NEEDS_SHA_RE.search(low) is not None
            or ("github.event" in low and "sha" in low))


def _is_default_branch_evidence(line, forbidden_branches):
    if DEFAULT_BRANCH_RE.search(line):
        return True
    low = line.lower()
    return any(re.search(rf"[\s/'\"]{re.escape(b)}[\s/'\"']", low)
               for b in forbidden_branches)


def check_evidence_job(workflow_text, job, forbidden_branches):
    """Returns (ok, findings, evidence_lines) for one declared evidence job."""
    findings = []
    block = _job_block(workflow_text, job)
    if block is None:
        return False, [f"el job '{job}' no existe (o aparece más de una vez) "
                       "en el workflow declarado."], []
    # las variables GLOBALES viven en el env: de nivel de workflow (columna 0)
    global_env_keys = _workflow_env_keys(workflow_text)
    evidence_lines = [l.strip() for l in block.splitlines()
                      if any(m in l for m in EVIDENCE_MARKERS)]
    if not evidence_lines:
        findings.append(f"el job '{job}' no publica evidencia sobre commits "
                        "(ninguna línea referencia statuses/, check-runs o head_sha).")
        return False, findings, []
    with_sha = [l for l in evidence_lines if _has_sha_expression(l)]
    if not with_sha:
        findings.append(f"el job '{job}' publica evidencia sin un SHA explícito "
                        "(github.sha / GITHUB_SHA / needs.*.outputs / github.event.*.sha).")
    for line in evidence_lines:
        if _is_default_branch_evidence(line, forbidden_branches):
            findings.append(f"evidencia sobre la rama por defecto (github.ref o rama "
                            f"prohibida) en: {line.strip()[:80]}")
        for token in _sha_position_tokens(line):
            if _global_sha_var(token, global_env_keys):
                findings.append(f"evidencia sobre una variable global del workflow "
                                f"({token}) en: {line.strip()[:80]}")
    return (not findings), findings, evidence_lines


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Read-only deploy-evidence gate (HR-10, #253).")
    ap.add_argument("--policy", required=True)
    ap.add_argument("--repo", default=".")
    args = ap.parse_args(argv)
    try:
        policy = json.loads(Path(args.policy).read_text(encoding="utf-8"))
        if not isinstance(policy, dict) or not isinstance(policy.get("evidence_jobs"), list):
            raise KeyError("evidence_jobs debe ser una lista")
        forbidden = policy.get("forbidden_branches", ["main"])
        if not isinstance(forbidden, list):
            raise KeyError("forbidden_branches debe ser una lista")
    except (OSError, json.JSONDecodeError, KeyError, TypeError) as exc:
        return fail2(f"política ilegible o incompleta ({args.policy}): {exc}")

    evidence_jobs = policy["evidence_jobs"]
    if not evidence_jobs:
        print("DEPLOY EVIDENCE OK: sin jobs de evidencia declarados "
              "(repo sin deploys; estado declarado en la política).")
        return 0
    if not (Path(args.repo) / ".github" / "workflows").is_dir():
        return fail2(f"'{args.repo}/.github/workflows' no existe; sujeto ausente es duda (HR-3).")

    findings, checked = [], 0
    for entry in evidence_jobs:
        workflow, job = entry.get("workflow"), entry.get("job")
        if not isinstance(workflow, str) or not isinstance(job, str) or not workflow or not job:
            findings.append(f"entrada de evidence_jobs mal formada: {entry!r}.")
            continue
        wf_path = Path(args.repo) / workflow
        if not wf_path.is_file():
            findings.append(f"el workflow declarado '{workflow}' no existe.")
            continue
        ok, job_findings, _ = check_evidence_job(wf_path.read_text(encoding="utf-8"),
                                                 job, forbidden)
        checked += 1
        if ok:
            print(f"  ok  {workflow}#{job}: evidencia sobre el SHA desplegado")
        findings.extend(job_findings)
    if findings:
        print(f"FAIL: evidencia de deploy con {len(findings)} hallazgo(s):")
        for f in findings:
            print(f"  ✗ {f}")
        return 1
    print(f"DEPLOY EVIDENCE OK: {checked} job(s) de evidencia publican sobre el SHA desplegado.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
