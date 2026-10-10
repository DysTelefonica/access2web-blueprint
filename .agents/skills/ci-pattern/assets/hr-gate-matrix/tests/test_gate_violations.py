#!/usr/bin/env python3
# ci-pattern asset — HR-32 meta-gate: every matrix gate sees a violation fail (#245)
"""Meta-gate for HR-32: every ``gate`` entry in the HR→gate matrix carries a
test that exists AND a violation recipe proving the gate fails (exit != 0)
against a violation fixture. A gate whose violation check passes cleanly —
or a gate with no recipe at all — is the violation this gate reports.

The recipes live in this file as data + code: each one builds its violating
input locally (temp dirs, throwaway git repositories, no network) and runs
the real gate binary the matrix entry declares. Recipes are keyed by the
matrix entry's ``test`` field, so several HRs sharing one gate share one
recipe. This file's own matrix entry (HR-32) is the check itself and is
exempt from the registry by construction.

Exit codes: 0 every gate rejects its violation · 1 findings (a gate exists
that does not fail against its violation, or has no recipe) · 2 invalid
input (matrix unreadable).
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[3]
MATRIX = SKILL_ROOT / "references" / "hr-gate-matrix.json"




def run(cmd, **kw):
    return subprocess.run(cmd, capture_output=True, text=True, **kw)


def recipe_workflow_policy() -> int:
    """HR-3/26/30: a workflow with an unpinned action must fail."""
    with tempfile.TemporaryDirectory() as d:
        wf = Path(d, "wf.yml")
        wf.write_text("on:\n  pull_request:\njobs:\n  build:\n"
                      "    steps:\n      - uses: actions/checkout@v4\n", encoding="utf-8")
        pol = Path(d, "policy.json")
        pol.write_text(json.dumps({"required_jobs": ["build"]}), encoding="utf-8")
        gate = SKILL_ROOT / "assets/workflow-policy/check_workflow_policy.py"
        return run([sys.executable, str(gate), "--workflow", str(wf),
                    "--required-jobs", str(pol)]).returncode


def recipe_pr_contract() -> int:
    """HR-6/7/8/14/31: a PR without the required labels must fail."""
    with tempfile.TemporaryDirectory() as d:
        repo = Path(d, "repo")
        repo.mkdir()
        for args in (["init", "-q", "-b", "main"],
                     ["config", "user.name", "v"], ["config", "user.email", "v@v"],
                     ["commit", "-q", "--allow-empty", "-m", "chore: base"],
                     ["checkout", "-q", "-b", "feature/x"],
                     ["commit", "-q", "--allow-empty", "-m", "feat: x"]):
            run(["git", "-C", str(repo), *args])
        event = Path(d, "event.json")
        event.write_text(json.dumps({
            "action": "opened",
            "pull_request": {"number": 1, "base": {"ref": "main", "sha": "0" * 40},
                             "head": {"ref": "feature/x", "sha": "0" * 40},
                             "labels": [], "body": None, "title": "test PR"},
            "repository": {"default_branch": "main"},
        }), encoding="utf-8")
        gate = SKILL_ROOT / "assets/pr-contract/check_pr_contract.py"
        policy = SKILL_ROOT / "assets/pr-contract/pr-contract.policy.example.json"
        rc_labels = run([sys.executable, str(gate), "--event-file", str(event),
                         "--policy-file", str(policy), "--repo", str(repo)]).returncode
    # #334 (HR-53): un PR que cierra una issue y no nombra el test que la
    # prueba es violación del gate.
    with tempfile.TemporaryDirectory() as d:
        repo = Path(d, "repo")
        repo.mkdir()
        for args in (["init", "-q", "-b", "main"],
                     ["config", "user.name", "v"], ["config", "user.email", "v@v"],
                     ["commit", "-q", "--allow-empty", "-m", "chore: base"],
                     ["checkout", "-q", "-b", "feature/x"],
                     ["commit", "-q", "--allow-empty", "-m", "feat: x"]):
            run(["git", "-C", str(repo), *args])
        event = Path(d, "event.json")
        event.write_text(json.dumps({
            "action": "opened",
            "pull_request": {"number": 2, "base": {"ref": "main", "sha": "0" * 40},
                             "head": {"ref": "feature/x", "sha": "0" * 40},
                             "labels": [{"name": "type:feature"}],
                             "body": "Closes #42\n\nSin sección de tests.\n",
                             "title": "test PR"},
            "repository": {"default_branch": "main"},
        }), encoding="utf-8")
        rc_closure = run([sys.executable, str(gate), "--event-file", str(event),
                          "--policy-file", str(policy), "--repo", str(repo)]).returncode
    # #333 (HR-54): un PR de cadena sin la sección Chain Context es violación.
    with tempfile.TemporaryDirectory() as d:
        repo = Path(d, "repo")
        repo.mkdir()
        for args in (["init", "-q", "-b", "main"],
                     ["config", "user.name", "v"], ["config", "user.email", "v@v"],
                     ["commit", "-q", "--allow-empty", "-m", "chore: base"],
                     ["checkout", "-q", "-b", "feature/x"],
                     ["commit", "-q", "--allow-empty", "-m", "feat: x"]):
            run(["git", "-C", str(repo), *args])
        event = Path(d, "event.json")
        event.write_text(json.dumps({
            "action": "opened",
            "pull_request": {"number": 3, "base": {"ref": "main", "sha": "0" * 40},
                             "head": {"ref": "feature/x", "sha": "0" * 40},
                             "labels": [{"name": "type:feature"},
                                        {"name": "chain:partial"}],
                             "body": "Refs #333\n", "title": "test PR"},
            "repository": {"default_branch": "main"},
        }), encoding="utf-8")
        rc_chain = run([sys.executable, str(gate), "--event-file", str(event),
                        "--policy-file", str(policy), "--repo", str(repo)]).returncode
    for rc in (rc_labels, rc_closure, rc_chain):
        if rc == 0:
            return 0
    return 1


def recipe_ratchet() -> int:
    """HR-15/16: a new finding identity must fail."""
    with tempfile.TemporaryDirectory() as d:
        f = Path(d, "findings.json")
        f.write_text(json.dumps({"measured": True, "subjects": 1, "findings": [
            {"rule": "r", "path": "p", "fingerprint": "new-identity"}]}), encoding="utf-8")
        b = Path(d, "baseline.json")
        b.write_text(json.dumps({"baseline_version": 1, "entries": []}), encoding="utf-8")
        gate = SKILL_ROOT / "assets/ratchet/check_ratchet.py"
        return run([sys.executable, str(gate), "--findings", str(f),
                    "--baseline", str(b), "--today", "2026-10-04"]).returncode


def recipe_release_gate() -> int:
    """HR-19/20: a semver tag on a side branch must fail."""
    with tempfile.TemporaryDirectory() as d:
        repo = Path(d, "repo")
        repo.mkdir()
        for args in (["init", "-q", "-b", "main"],
                     ["config", "user.name", "v"], ["config", "user.email", "v@v"],
                     ["commit", "-q", "--allow-empty", "-m", "chore: base"],
                     ["checkout", "-q", "-b", "side"],
                     ["commit", "-q", "--allow-empty", "-m", "feat: side"],
                     ["tag", "-a", "v1.0.0", "-m", "release"]):
            run(["git", "-C", str(repo), *args])
        gate = SKILL_ROOT / "assets/release-gate/check_release_tag.py"
        return run([sys.executable, str(gate), "--repo", str(repo),
                    "--tag", "v1.0.0", "--default-branch", "main"]).returncode


def recipe_required_jobs() -> int:
    """HR-29: a failed required job must fail the verdict."""
    with tempfile.TemporaryDirectory() as d:
        pol = Path(d, "policy.json")
        pol.write_text(json.dumps({"required_jobs": ["lint"]}), encoding="utf-8")
        needs = Path(d, "needs.json")
        needs.write_text(json.dumps({"lint": {"result": "failure", "outputs": {}}}),
                         encoding="utf-8")
        gate = SKILL_ROOT / "assets/required-jobs/check_required_jobs.py"
        return run([sys.executable, str(gate), "--policy", str(pol),
                    "--event", "pull_request"], stdin=open(needs)).returncode


def recipe_host_readback() -> int:
    """HR-34: a declared label missing from the host snapshot must fail; and
    the runner-enforced class (#303) must fail where the host can enforce."""
    with tempfile.TemporaryDirectory() as d:
        gate = SKILL_ROOT / "assets/host-readback/check_host_drift.py"
        snaps = {"repo": {}, "labels": [{"name": "status:approved"}],
                 "branch-protection": {}, "rulesets": []}
        paths = []
        for kind, payload in snaps.items():
            path = Path(d, f"{kind}.json")
            path.write_text(json.dumps(payload), encoding="utf-8")
            paths += ["--snapshot", f"{kind}={path}"]
        contract = Path(d, "contract.json")
        contract.write_text(json.dumps({
            "contract_version": 1,
            "labels": [{"name": "status:approved", "class": "host-enforced"},
                       {"name": "status:blocked", "class": "host-enforced"}],
            "required_checks": [],
            "merge_methods": {},
            "protection": {},
            "rulesets": [],
        }), encoding="utf-8")
        rc_label = run([sys.executable, str(gate), "--contract", str(contract), *paths]).returncode
        # #303: fixture violador de la tercera clase — una capacidad que el
        # host aplica no admite una regla runner-enforced.
        contract.write_text(json.dumps({
            "contract_version": 1,
            "labels": [{"name": "status:approved", "class": "host-enforced"}],
            "required_checks": [],
            "merge_methods": {},
            "protection": {"enforce_admins": {"declared": False, "class": "runner-enforced"}},
            "rulesets": [],
            "runner": {"label": "cadete-oracle-arm64"},
        }), encoding="utf-8")
        rc_runner = run([sys.executable, str(gate), "--contract", str(contract), *paths]).returncode
        # HR-32 solo exige que el gate RECHACE su fixture: basta con que
        # alguno de los dos pase para que la receta sea verde (0).
        return 0 if rc_label == 0 or rc_runner == 0 else rc_label


def recipe_push_compliance() -> int:
    """HR-49: un push a rama protegida sin PR mergeado debe abrir incidente (1)."""
    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        sha, base = "1f4e9a2c7b8d3e6f5a0c9b8d7e6f5a4c3b2d1e0f", "0a1b2c3d4e5f60718293a4b5c6d7e8f90a1b2c3d"
        commit = {"id": sha, "message": "x", "author": {"name": "Dev", "email": "d@e"}}
        root.joinpath("event.json").write_text(json.dumps({
            "ref": "refs/heads/main", "before": base, "after": sha, "forced": False, "deleted": False,
            "commits": [commit], "head_commit": commit, "pusher": {"name": "dev"},
            "repository": {"full_name": "o/r", "default_branch": "main"}}), encoding="utf-8")
        root.joinpath("contract.json").write_text(json.dumps({
            "contract_version": 1, "protected_branches": ["main"], "required_checks": []}), encoding="utf-8")
        shim = root / "bin"
        shim.mkdir()
        gh = shim / "gh"
        gh.write_text("#!/usr/bin/env bash\n"
                      "case \"$*\" in\n"
                      "  *\" -X POST \"*) echo '{\"number\": 1, \"html_url\": \"https://github.com/o/r/issues/1\"}' ;;\n"
                      "  *compare*) echo '{\"total_commits\": 1, \"commits\": [{\"sha\": \"1f4e9a2c7b8d3e6f5a0c9b8d7e6f5a4c3b2d1e0f\", \"commit\": {\"author\": {\"name\": \"Dev\"}}}]}' ;;\n"
                      "  *) echo \"[]\" ;;\n"
                      "esac\n", encoding="utf-8")
        gh.chmod(0o755)
        env = {**os.environ, "PATH": f"{shim}{os.pathsep}{os.environ.get('PATH', '')}"}
        gate = SKILL_ROOT / "assets/runner-controls/check_push_compliance.py"
        return run([sys.executable, str(gate), "--event", str(root / "event.json"),
                    "--contract", str(root / "contract.json")], env=env, timeout=60).returncode


def recipe_governed_merge() -> int:
    """HR-50: un PR sin etiqueta type:* no debe fusionar (1)."""
    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        root.joinpath("contract.json").write_text(json.dumps({
            "contract_version": 1, "required_checks": []}), encoding="utf-8")
        shim = root / "bin"
        shim.mkdir()
        gh = shim / "gh"
        gh.write_text("#!/usr/bin/env bash\\n"
                      "case \"$*\" in\\n"
                      "  *check-runs*) echo '{\"total_count\": 0, \"check_runs\": []}' ;;\\n"
                      "  *\"/pulls/7\") echo '{\"number\": 7, \"state\": \"open\", \"head\": {\"sha\": \"abc\"}, \"labels\": [], \"additions\": 1, \"deletions\": 1, \"body\": \"\"}' ;;\\n"
                      "  *) echo \"unexpected: $*\" >&2; exit 1 ;;\\n"
                      "esac\\n", encoding="utf-8")
        gh.chmod(0o755)
        env = {**os.environ, "PATH": f"{shim}{os.pathsep}{os.environ.get('PATH', '')}"}
        gate = SKILL_ROOT / "assets/runner-controls/check_governed_merge.py"
        return run([sys.executable, str(gate), "--repo", "o/r", "--pr", "7",
                    "--contract", str(root / "contract.json")], env=env, timeout=60).returncode


def recipe_force_push_guard() -> int:
    """HR-51: un force-push a rama protegida debe abrir incidente (1)."""
    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        sha, base = "1f4e9a2c7b8d3e6f5a0c9b8d7e6f5a4c3b2d1e0f", "0a1b2c3d4e5f60718293a4b5c6d7e8f90a1b2c3d"
        commit = {"id": sha, "message": "x", "author": {"name": "Dev", "email": "d@e"}}
        root.joinpath("event.json").write_text(json.dumps({
            "ref": "refs/heads/main", "before": base, "after": sha, "forced": True, "deleted": False,
            "commits": [commit], "head_commit": commit, "pusher": {"name": "dev"},
            "repository": {"full_name": "o/r", "default_branch": "main"}}), encoding="utf-8")
        root.joinpath("contract.json").write_text(json.dumps(
            {"contract_version": 1, "protected_branches": ["main"]}), encoding="utf-8")
        shim = root / "bin"
        shim.mkdir()
        gh = shim / "gh"
        gh.write_text("#!/usr/bin/env bash\\n"
                      "case \"$*\" in\\n"
                      "  *\" -X POST \"*) echo '{\"number\": 1, \"html_url\": \"https://github.com/o/r/issues/1\"}' ;;\\n"
                      "  *) exit 1 ;;\\n"
                      "esac\\n", encoding="utf-8")
        gh.chmod(0o755)
        env = {**os.environ, "PATH": f"{shim}{os.pathsep}{os.environ.get('PATH', '')}"}
        gate = SKILL_ROOT / "assets/runner-controls/check_force_push_guard.py"
        return run([sys.executable, str(gate), "--event", str(root / "event.json"),
                    "--contract", str(root / "contract.json")], env=env, timeout=60).returncode


def recipe_runner_binding() -> int:
    """HR-52: una declaracion runner-enforced sin workflow debe fallar (1)."""
    with tempfile.TemporaryDirectory() as d:
        root = Path(d)
        root.joinpath("contract.json").write_text(json.dumps({
            "contract_version": 1,
            "protection": {"enforce_admins": {"declared": True, "class": "runner-enforced"}},
            "runner": {"label": "cadete"}}), encoding="utf-8")
        wdir = root / "workflows"
        wdir.mkdir()
        (wdir / "merge.yml").write_text(
            "name: merge\non: push\njobs:\n  j:\n    runs-on: [self-hosted, cadete]\n"
            "    steps:\n      - run: python3 x/check_governed_merge.py\n", encoding="utf-8")
        gate = SKILL_ROOT / "assets/runner-controls/check_runner_binding.py"
        return run([sys.executable, str(gate), "--contract", str(root / "contract.json"),
                    "--workflows", str(wdir)], timeout=60).returncode


def recipe_branch_name() -> int:
    """HR-35: a name the gate rejects must fail."""
    gate = SKILL_ROOT / "assets/branch-name.sh"
    return run([str(gate), "validate", "Not A Branch Name"]).returncode


def recipe_exemptions() -> int:
    """HR-33: an exemption without a verifiable identity must fail."""
    with tempfile.TemporaryDirectory() as d:
        doc = Path(d, "doc.json")
        doc.write_text(json.dumps({"exemption": {"reason": "r", "reference": "x"}}),
                       encoding="utf-8")
        gate = SKILL_ROOT / "assets/exemptions/check_exemptions.py"
        return run([sys.executable, str(gate), "--file", str(doc)]).returncode


# Recipe registry: test path (relative to skill root) -> callable that runs
# the gate against a violation fixture and returns the gate's exit code.
def recipe_local_ci_parity() -> int:
    """HR-4: a CI command without a local counterpart must fail."""
    with tempfile.TemporaryDirectory() as d:
        repo = Path(d, "repo")
        wf = repo / ".github" / "workflows"
        wf.mkdir(parents=True)
        (wf / "ci.yml").write_text(
            "jobs:\n  unit:\n    steps:\n"
            "      - run: python3 ghost.py\n", encoding="utf-8")
        ep = repo / "testing" / "run-unit-suites.sh"
        ep.parent.mkdir(parents=True)
        ep.write_text("if bash local.sh; then\n  echo PASS\nfi\n", encoding="utf-8")
        (repo / "local.sh").write_text("#!/bin/sh\n", encoding="utf-8")
        policy = Path(d, "policy.json")
        policy.write_text(json.dumps({
            "ci": {"workflow": ".github/workflows/ci.yml", "job": "unit"},
            "local": {"entrypoint": "testing/run-unit-suites.sh"},
            "exclusions": []}), encoding="utf-8")
        gate = SKILL_ROOT / "assets/local-ci-parity/check_local_ci_parity.py"
        return run([sys.executable, str(gate), "--repo", str(repo),
                    "--policy", str(policy)]).returncode


def recipe_env_isolation() -> int:
    """HR-28: a suite reading a local .env must fail."""
    with tempfile.TemporaryDirectory() as d:
        suite = Path(d, "testing", "suites", "a.sh")
        suite.parent.mkdir(parents=True, exist_ok=True)
        suite.write_text('#!/usr/bin/env bash\nset -a\nsource .env\nset +a\n',
                         encoding="utf-8")
        allow = Path(d, "allowlist.json")
        allow.write_text(json.dumps({"env_vars": ["TMPDIR"]}), encoding="utf-8")
        gate = SKILL_ROOT / "assets/env-isolation/check_env_isolation.py"
        rc_dotenv = run([sys.executable, str(gate), "--root", str(suite.parents[2]),
                         "--allowlist", str(allow)]).returncode
    # #315 (HR-28): un árbol con ficheros de código fuera del dominio
    # declarado del gate, y sin exclusión explícita, es hallazgo.
    with tempfile.TemporaryDirectory() as d:
        root = Path(d, "testing")
        (root / "suites").mkdir(parents=True)
        (root / "suites" / "a.sh").write_text("#!/bin/sh\necho ok\n", encoding="utf-8")
        (root / "suites" / "b.php").write_text("<?php echo 1;\n", encoding="utf-8")
        allow = Path(d, "allowlist.json")
        allow.write_text(json.dumps({"env_vars": ["TMPDIR"]}), encoding="utf-8")
        rc_domain = run([sys.executable, str(gate), "--root", str(root),
                         "--allowlist", str(allow)]).returncode
    for rc in (rc_dotenv, rc_domain):
        if rc == 0:
            return 0
    return 1


def recipe_structured_inputs() -> int:
    """HR-1: a gate scraping host text output must fail."""
    with tempfile.TemporaryDirectory() as d:
        skill = Path(d, "skill")
        (skill / "references").mkdir(parents=True)
        gate = skill / "assets" / "scraper.sh"
        gate.parent.mkdir(parents=True, exist_ok=True)
        # hr-1: fixture — el texto violador es dato del meta-gate HR-32, no una lectura real
        gate.write_text('title=$(gh pr view 123)\n', encoding="utf-8")  # hr-1: fixture — dato del meta-gate HR-32, no una lectura real
        matrix = skill / "references" / "hr-gate-matrix.json"
        matrix.write_text(json.dumps({"rules": [
            {"id": "HR-1", "enforcement": "gate", "asset": "assets/scraper.sh",
             "test": "assets/scraper.sh"}]}), encoding="utf-8")
        check = SKILL_ROOT / "assets/hr-gate-matrix/check_structured_inputs.py"
        return run([sys.executable, str(check), "--skill-root", str(skill)]).returncode


def recipe_deploy_policy() -> int:
    """HR-40/HR-41: an asymmetric bootstrap policy must fail."""
    with tempfile.TemporaryDirectory() as d:
        policy = Path(d, "policy.json")
        policy.write_text(json.dumps({
            "policy_version": 1,
            "resolution_source": {"kind": "health-endpoint",
                                  "location": "https://host/healthz"},
            "contexts": {
                "release/e2e-production": {"required": True,
                                           "bootstrap_exempt": True},
                "release/smoke-production": {"required": True,
                                             "bootstrap_exempt": False},
            }}), encoding="utf-8")
        gate = SKILL_ROOT / "assets/deploy-review/check_deploy_policy.py"
        return run([sys.executable, str(gate), "--policy", str(policy)]).returncode
def recipe_runbook_premises() -> int:
    """HR-42: a state premise without its runbook command must fail."""
    with tempfile.TemporaryDirectory() as d:
        rb = Path(d, "RELEASE-v1.0.0.md")
        rb.write_text("The e2e user is confirmed seeded.\n\n"
                      "```bash\n./deploy.sh\n```\n", encoding="utf-8")
        patterns = Path(d, "patterns.json")
        patterns.write_text(json.dumps(
            {"premise_patterns": ["confirmed seeded"]}), encoding="utf-8")
        gate = SKILL_ROOT / "assets/runbook-premises/check_runbook_premises.py"
        return run([sys.executable, str(gate), "--root", str(d),
                    "--patterns", str(patterns)]).returncode



def recipe_deploy_evidence() -> int:
    """HR-10: evidence published over the default branch must fail."""
    with tempfile.TemporaryDirectory() as d:
        repo = Path(d, "repo")
        wf = repo / ".github" / "workflows"
        wf.mkdir(parents=True)
        (wf / "deploy.yml").write_text(
            "jobs:\n  deploy:\n    runs-on: ubuntu-24.04\n    steps:\n"
            "      - run: gh api repos/$GITHUB_REPOSITORY/statuses/${{ github.ref }}"
            " --field state=success\n", encoding="utf-8")
        policy = Path(d, "policy.json")
        policy.write_text(json.dumps({
            "evidence_jobs": [{"workflow": ".github/workflows/deploy.yml",
                               "job": "deploy"}]}), encoding="utf-8")
        gate = SKILL_ROOT / "assets/deploy-evidence/check_deploy_evidence.py"
        return run([sys.executable, str(gate), "--repo", str(repo),
                    "--policy", str(policy)]).returncode


def recipe_adoption_check() -> int:
    """HR-45/HR-46: un contrato de adopción con disposición inválida (fase 0),
    un árbol con gobierno residual no inventariado (fase 3), un documento
    operativo con una etiqueta derivada (fase 5, #279), un target
    'adopted' fuera del catálogo canónico (fase 0, #306) y una etiqueta del
    host sin inventariar (fase 0, #312) y un script invocado por el
    gobierno sin inventariar (fase 0, #311/#312) deben fallar."""
    gate = SKILL_ROOT / "assets/bin/ci-pattern"
    with tempfile.TemporaryDirectory() as d:
        repo = Path(d, "repo")
        (repo / ".github" / "workflows").mkdir(parents=True)
        (repo / ".github" / "workflows" / "ci.yml").write_text("on: push\n", encoding="utf-8")
        contract = {
            "schema_version": 1,
            "phases": {
                "0": {
                    "inventory": [
                        {"artifact": ".github/workflows/ci.yml", "disposition": "kept",
                         "target": ".github/workflows/ci.yml"}
                    ],
                    "removal_confirmed": True
                }
            },
        }
        (repo / ".github" / "ci-pattern-adoption.json").write_text(
            json.dumps(contract), encoding="utf-8")
        rc45 = run([sys.executable, str(gate), "adoption", "check", "--phase", "0",
                    str(repo)]).returncode
    with tempfile.TemporaryDirectory() as d:
        repo = Path(d, "repo")
        (repo / ".github" / "workflows").mkdir(parents=True)
        (repo / ".github" / "workflows" / "ci.yml").write_text("on: push\n", encoding="utf-8")
        (repo / ".github" / "workflows" / "legacy.yml").write_text("on: push\n", encoding="utf-8")
        (repo / "docs" / "adoption").mkdir(parents=True)
        (repo / "tests").mkdir()
        (repo / "tests" / "test_smoke.sh").write_text("#!/bin/sh\necho ok\n", encoding="utf-8")
        (repo / ".github" / "env-isolation-allowlist.json").write_text(
            json.dumps({"env_vars": []}), encoding="utf-8")
        (repo / "docs/adoption/branch-protection.json").write_text("{}", encoding="utf-8")
        (repo / "docs/adoption/rulesets.json").write_text("[]", encoding="utf-8")
        (repo / "docs/adoption/phase1.md").write_text("evidencia\n", encoding="utf-8")
        (repo / "docs/adoption/phase2.md").write_text("evidencia\n", encoding="utf-8")
        (repo / "docs/adoption/phase3.md").write_text("evidencia\n", encoding="utf-8")
        # La identidad del patrón sale del manifiesto verificado y sus
        # rulesets del contrato del host (#277); se siembran válidos para que
        # el disparo de la receta sea el gobierno residual, no la duda.
        (repo / ".github" / "host-contract.json").write_text(
            json.dumps({"rulesets": [{"name": "branch-protection",
                                      "enforcement": "active", "class": "branch"}]}),
            encoding="utf-8")
        (repo / ".governance-manifest.json").write_text(
            json.dumps({"schema_version": 1, "files": {
                ".github/workflows/ci.yml": "sha256:" + __import__("hashlib").sha256(
                    b"on: push\n").hexdigest(),
                ".github/host-contract.json": "sha256:" + __import__("hashlib").sha256(
                    (repo / ".github" / "host-contract.json").read_bytes()).hexdigest(),
            }}),
            encoding="utf-8")
        contract = {
            "schema_version": 1,
            "phases": {
                "0": {
                    "inventory": [
                        {"artifact": ".github/workflows/ci.yml", "disposition": "adopted",
                         "target": ".github/workflows/ci.yml"}
                    ],
                    "removal_confirmed": True
                },
                "1": {"evidence": {"path": "docs/adoption/phase1.md"}},
                "2": {"env_isolation": {
                    "root": "tests",
                    "allowlist": ".github/env-isolation-allowlist.json"}},
                "3": {
                    "evidence": {"path": "docs/adoption/phase3.md"},
                    "pattern_paths": [".github/workflows/ci.yml"],
                    "host_snapshots": {
                        "branch-protection": "docs/adoption/branch-protection.json",
                        "rulesets": "docs/adoption/rulesets.json"
                    }
                },
            },
        }
        (repo / ".github" / "ci-pattern-adoption.json").write_text(
            json.dumps(contract), encoding="utf-8")
        rc46 = run([sys.executable, str(gate), "adoption", "check", "--phase", "3",
                    str(repo)]).returncode
    # #279: reutiliza los fixtures de la suite del gate para exigir que la
    # fase 5 RECHAZA un doc operativo con una etiqueta documentada derivada.
    import importlib.machinery
    import importlib.util
    spec = importlib.util.spec_from_loader(
        "adoption_gate_tests",
        importlib.machinery.SourceFileLoader(
            "adoption_gate_tests",
            str(SKILL_ROOT / "assets/adoption/tests/test_adoption_check.py")))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)

    def docs_repo(d):
        repo, _ = mod.make_contract_repo(d, contract=mod.complete_contract(upto=5))
        mod.seed_evidence(repo, 3)
        mod.seed_operating_doc(repo)
        return repo / "docs" / "ci-pattern-flow.md"

    with tempfile.TemporaryDirectory() as d:
        doc = docs_repo(d)
        text = doc.read_text(encoding="utf-8")
        assert "type:bug" in text
        doc.write_text(
            text.replace("`bug_report.yml` → `type:bug`",
                         "`bug_report.yml` → `type:chore`", 1), encoding="utf-8")
        rc79 = run([sys.executable, str(gate), "adoption", "check", "--phase", "5",
                    str(doc.parent)]).returncode
    # #305 (tabla 1:1): dirección 1 — job requerido sin fila en la tabla.
    with tempfile.TemporaryDirectory() as d:
        doc = docs_repo(d)
        text = doc.read_text(encoding="utf-8")
        row = next(l for l in text.splitlines() if l.startswith("| `unit` |"))
        doc.write_text(text.replace(row + "\n", "", 1), encoding="utf-8")
        rc_table_missing = run([sys.executable, str(gate), "adoption", "check",
                                "--phase", "5", str(doc.parent)]).returncode
    # #305 (tabla 1:1): dirección 2 — fila que documenta un check inexistente.
    with tempfile.TemporaryDirectory() as d:
        doc = docs_repo(d)
        text = doc.read_text(encoding="utf-8")
        row = next(l for l in text.splitlines() if l.startswith("| `unit` |"))
        doc.write_text(
            text.replace(row, row + "\n| `ghost-check` | inventado | no | no |", 1),
            encoding="utf-8")
        rc_table_unknown = run([sys.executable, str(gate), "adoption", "check",
                                "--phase", "5", str(doc.parent)]).returncode
    # #305 (bloque AGENTS.md): dirección 1 — bloque ausente.
    with tempfile.TemporaryDirectory() as d:
        doc = docs_repo(d)
        agents = doc.parent.parent / "AGENTS.md"
        text = agents.read_text(encoding="utf-8")
        begin = "<!-- ci-pattern-agents:begin -->"
        end = "<!-- ci-pattern-agents:end -->"
        agents.write_text(
            text[:text.index(begin)] + text[text.index(end) + len(end) + 1:],
            encoding="utf-8")
        rc_agents_missing = run([sys.executable, str(gate), "adoption", "check",
                                 "--phase", "5", str(doc.parent)]).returncode
    # #305 (bloque AGENTS.md): dirección 2 — copia alterada del mandato.
    with tempfile.TemporaryDirectory() as d:
        doc = docs_repo(d)
        agents = doc.parent.parent / "AGENTS.md"
        agents.write_text(
            agents.read_text(encoding="utf-8").replace(
                "no lo invente", "puede inferirse", 1), encoding="utf-8")
        rc_agents_altered = run([sys.executable, str(gate), "adoption", "check",
                                 "--phase", "5", str(doc.parent)]).returncode
    # #306: un target 'adopted' fuera del catálogo canónico del patrón es
    # hallazgo de la fase 0, aunque el fichero no exista todavía.
    with tempfile.TemporaryDirectory() as d:
        repo = Path(d, "repo")
        (repo / ".github").mkdir(parents=True)
        contract = {
            "schema_version": 1,
            "phases": {
                "0": {
                    "inventory": [
                        {"artifact": "legacy-build", "disposition": "adopted",
                         "target": "Jenkinsfile"}
                    ],
                    "removal_confirmed": True
                }
            },
        }
        (repo / ".github" / "ci-pattern-adoption.json").write_text(
            json.dumps(contract), encoding="utf-8")
        rc306 = run([sys.executable, str(gate), "adoption", "check", "--phase", "0",
                     str(repo)]).returncode
    # #312: una etiqueta del host sin entrada en el inventario es hallazgo
    # de fase 0 (cobertura de exhaustividad contra la instantánea declarada).
    with tempfile.TemporaryDirectory() as d:
        repo = Path(d, "repo")
        (repo / ".github" / "workflows").mkdir(parents=True)
        (repo / ".github" / "workflows" / "ci.yml").write_text("on: push\n", encoding="utf-8")
        (repo / ".github" / "host-labels-snapshot.json").write_text(
            json.dumps(["status:approved", "security"]), encoding="utf-8")
        contract = {
            "schema_version": 1,
            "phases": {
                "0": {
                    "inventory": [
                        {"artifact": ".github/workflows/ci.yml",
                         "disposition": "adopted",
                         "target": ".github/workflows/ci.yml"},
                        {"artifact": "labels status:approved",
                         "disposition": "adopted",
                         "target": ".github/host-contract.json"},
                    ],
                    "removal_confirmed": True,
                    "host_labels": ".github/host-labels-snapshot.json",
                }
            },
        }
        (repo / ".github" / "ci-pattern-adoption.json").write_text(
            json.dumps(contract), encoding="utf-8")
        rc312 = run([sys.executable, str(gate), "adoption", "check", "--phase", "0",
                     str(repo)]).returncode
    # #311/#312: un script que el gobierno INVOCA (el `run:` del workflow de CI)
    # y no tiene entrada de inventario es hallazgo de fase 0. Un script
    # operativo que nadie del gobierno invoca NO lo es (ver asserts de
    # adopción): el glob `scripts/*` retiraba scripts necesarios de la APP.
    with tempfile.TemporaryDirectory() as d:
        repo = Path(d, "repo")
        (repo / ".github" / "workflows").mkdir(parents=True)
        (repo / "scripts").mkdir()
        (repo / "scripts" / "gate.sh").write_text("#!/bin/sh\n", encoding="utf-8")
        (repo / "scripts" / "backup.ps1").write_text("# app\n", encoding="utf-8")
        (repo / ".github" / "workflows" / "ci.yml").write_text(
            "on: push\njobs:\n  x:\n    steps:\n      - run: bash scripts/gate.sh\n",
            encoding="utf-8")
        (repo / ".github" / "host-labels-snapshot.json").write_text("[]", encoding="utf-8")
        contract = {
            "schema_version": 1,
            "phases": {
                "0": {
                    "inventory": [
                        {"artifact": ".github/workflows/ci.yml",
                         "disposition": "adopted",
                         "target": ".github/workflows/ci.yml"},
                    ],
                    "removal_confirmed": True,
                    "host_labels": ".github/host-labels-snapshot.json",
                }
            },
        }
        (repo / ".github" / "ci-pattern-adoption.json").write_text(
            json.dumps(contract), encoding="utf-8")
        rc319 = run([sys.executable, str(gate), "adoption", "check", "--phase", "0",
                     str(repo)]).returncode
    # #319 (HR-48): un AGENTS.md que manda la documentación fuera del repo
    # es hallazgo de la fase 3, nombrando fichero y línea.
    with tempfile.TemporaryDirectory() as d:
        repo, _ = mod.make_contract_repo(d, contract=mod.complete_contract(upto=3))
        mod.seed_evidence(repo, 3)
        agents = repo / "AGENTS.md"
        agents.write_text(agents.read_text(encoding="utf-8")
                          + "Documentation belongs in `C:\\00repos\\documentacion`, "
                            "not in this repo.\n", encoding="utf-8")
        mod.seed_manifest(repo)
        rc_docs_location = run([sys.executable, str(gate), "adoption", "check",
                                "--phase", "3", str(repo)]).returncode
    # #322 (HR-21): un post-mortem con una credencial es hallazgo de la
    # fase 5, nombrando fichero y línea. Las IPs y los comandos internos
    # están permitidos: el registro operativo es el detalle técnico.
    with tempfile.TemporaryDirectory() as d:
        repo, _ = mod.make_contract_repo(d, contract=mod.complete_contract(upto=5))
        mod.seed_evidence(repo, 3)
        mod.seed_operating_doc(repo)
        (repo / "docs" / "postmortems").mkdir(parents=True)
        (repo / "docs" / "postmortems" / "2026-10-06-caida.md").write_text(
            "# Caída de producción\n\n- password: hunter2\n", encoding="utf-8")
        rc_incident = run([sys.executable, str(gate), "adoption", "check",
                           "--phase", "5", str(repo)]).returncode
    # #330: el 403 de plan del host vale como instantánea solo si el contrato
    # declara esa área `unavailable`; con `available` es drift (exit 1).
    with tempfile.TemporaryDirectory() as d:
        repo, _ = mod.make_contract_repo(d, contract=mod.complete_contract(upto=3))
        host_path = repo / ".github" / "host-contract.json"
        host = json.loads(host_path.read_text(encoding="utf-8"))
        host["host_capabilities"] = {"branch_protection": "available",
                                     "rulesets": "available"}
        host_path.write_text(json.dumps(host), encoding="utf-8")
        mod.seed_evidence(repo, 3)
        (repo / "docs" / "adoption" / "rulesets.json").write_text(json.dumps({
            "message": "Upgrade to GitHub Pro or make this repository public "
                       "to enable this feature.",
            "status": "403"}), encoding="utf-8")
        rc_plan_403 = run([sys.executable, str(gate), "adoption", "check",
                           "--phase", "3", str(repo)]).returncode
    # Cada fixture se verifica: devolver el primer no-cero dejaba sin
    # comprobar todos los escenarios siguientes. Non-cero = el gate rechaza
    # su violación (lo que HR-32 espera); 0 = lo ACEPTÓ y HR-32 falla.
    for rc in (rc45, rc46, rc79, rc_table_missing, rc_table_unknown,
               rc_agents_missing, rc_agents_altered, rc306, rc312, rc319,
               rc_docs_location, rc_incident, rc_plan_403):
        if rc == 0:
            return 0
    return 1


def recipe_version_bump() -> int:
    """HR-44: skill files changed without version bump must fail."""
    with tempfile.TemporaryDirectory() as d:
        skill = Path(d, "skill")
        skill.mkdir(parents=True)
        (skill / "SKILL.md").write_text(
            "---\nname: x\nmetadata:\n  version: \"0.3\"\n---\nbody\n",
            encoding="utf-8")
        import subprocess as sp
        sp.run(["git", "init", "-q", "-b", "main"], cwd=d, check=True)
        sp.run(["git", "config", "user.name", "t"], cwd=d, check=True)
        sp.run(["git", "config", "user.email", "t@t"], cwd=d, check=True)
        sp.run(["git", "add", "-A"], cwd=d, check=True)
        sp.run(["git", "commit", "-q", "-m", "v0.3"], cwd=d, check=True)
        (skill / "new.txt").write_text("content\n", encoding="utf-8")
        sp.run(["git", "add", "-A"], cwd=d, check=True)
        sp.run(["git", "commit", "-q", "-m", "no bump"], cwd=d, check=True)
        gate = SKILL_ROOT / "assets/version-bump/check_version_bump.py"
        return run([sys.executable, str(gate), "--repo", str(d),
                    "--base", "main~1", "--head", "main",
                    "--skill-dir", "skill"]).returncode

def recipe_agents_structure() -> int:
    """HR-55: un AGENTS sin bloque con agents_md declarado en el contrato
    de adopción debe fallar (la violación propia de HR-55, #341)."""
    with tempfile.TemporaryDirectory() as d:
        repo = Path(d)
        (repo / "AGENTS.md").write_text("# Titulo\n", encoding="utf-8")
        bin_path = SKILL_ROOT / "assets/bin/ci-pattern"
        script = (
            "import importlib.util, importlib.machinery, sys\n"
            "spec = importlib.util.spec_from_loader('cip', importlib.machinery."
            f"SourceFileLoader('cip', {str(bin_path)!r}))\n"
            "m = importlib.util.module_from_spec(spec)\n"
            "spec.loader.exec_module(m)\n"
            "from pathlib import Path\n"
            "code = m._check_agents_structure(Path(sys.argv[1]), "
            "{'agents_md': 'propagated'})\n"
            "sys.exit(1 if code else 0)\n")
        return run([sys.executable, "-c", script, str(repo)]).returncode


def recipe_slice_content() -> int:
    """HR-56: una ruta personal en lo que compone la propagación debe
    fallar (la violación propia de HR-56, #358)."""
    with tempfile.TemporaryDirectory() as d:
        repo = Path(d)
        (repo / "fleet").mkdir()
        (repo / "fleet" / "registry.json").write_text(json.dumps({
            "schema_version": 1, "maintainers": ["mantener-one"],
            "repositories": [], "local_layout": {"local_root": str(repo)}}),
            encoding="utf-8")
        target = repo / "slices" / "partials" / "gov.md"
        target.parent.mkdir(parents=True)
        target.write_text("Ver personal/ardelperal/ci-pattern/SKILL.md\n",
                          encoding="utf-8")
        bin_path = SKILL_ROOT / "assets/bin/ci-pattern"
        script = (
            "import importlib.util, importlib.machinery, sys\n"
            "spec = importlib.util.spec_from_loader('cip', importlib.machinery."
            f"SourceFileLoader('cip', {str(bin_path)!r}))\n"
            "m = importlib.util.module_from_spec(spec)\n"
            "spec.loader.exec_module(m)\n"
            "from pathlib import Path\n"
            "code = m._check_slice_content(Path(sys.argv[1]), [Path(sys.argv[2])])\n"
            "sys.exit(1 if code else 0)\n")
        return run([sys.executable, "-c", script, str(repo), str(target)]).returncode


RECIPES = {
    "assets/workflow-policy/tests/test_workflow_policy.py": recipe_workflow_policy,
    "assets/pr-contract/tests/test_pr_contract.py": recipe_pr_contract,
    "assets/ratchet/tests/test_ratchet.py": recipe_ratchet,
    "assets/release-gate/tests/test_release_tag.py": recipe_release_gate,
    "assets/required-jobs/tests/test_required_jobs.py": recipe_required_jobs,
    "assets/host-readback/tests/test_host_drift.py": recipe_host_readback,
    "assets/runner-controls/tests/test_push_compliance.py": recipe_push_compliance,
    "assets/runner-controls/tests/test_governed_merge.py": recipe_governed_merge,
    "assets/runner-controls/tests/test_force_push_guard.py": recipe_force_push_guard,
    "assets/runner-controls/tests/test_runner_binding.py": recipe_runner_binding,
    "assets/local-ci-parity/tests/test_local_ci_parity.py": recipe_local_ci_parity,
    "assets/deploy-evidence/tests/test_deploy_evidence.py": recipe_deploy_evidence,
    "assets/branch-name.sh self-test": recipe_branch_name,
    "assets/exemptions/tests/test_exemptions.py": recipe_exemptions,
    "assets/hr-gate-matrix/tests/test_structured_inputs.py": recipe_structured_inputs,
    "assets/env-isolation/tests/test_env_isolation.py": recipe_env_isolation,
    "assets/deploy-review/tests/test_deploy_policy.py": recipe_deploy_policy,
    "assets/runbook-premises/tests/test_runbook_premises.py": recipe_runbook_premises,
    "assets/adoption/tests/test_adoption_check.py": [
        recipe_adoption_check, recipe_agents_structure, recipe_slice_content],
    "assets/version-bump/tests/test_version_bump.py": recipe_version_bump,
}

def main(argv):
    try:
        rules = json.loads(MATRIX.read_text(encoding="utf-8"))["rules"]
    except (OSError, json.JSONDecodeError, KeyError) as exc:
        print(f"HR-32 ERROR: matrix unreadable: {exc}", file=sys.stderr)
        return 2
    self_path = str(Path(__file__).relative_to(SKILL_ROOT))
    gates = {}
    for rule in rules:
        if rule.get("enforcement") == "gate" and rule.get("test") != self_path:
            gates.setdefault(rule["test"], []).append(rule["id"])
    report = []
    for test_path, hrs in sorted(gates.items()):
        recipes = RECIPES.get(test_path)
        if recipes is None:
            report.append(f"gate test '{test_path}' ({', '.join(hrs)}) has no "
                          "violation recipe: HR-32 cannot see this control fail")
            continue
        if not isinstance(recipes, list):
            recipes = [recipes]
        for recipe in recipes:
            try:
                code = recipe()
            except Exception as exc:  # noqa: BLE001 — a recipe crash IS the finding
                report.append(f"gate test '{test_path}' ({', '.join(hrs)}): "
                              f"recipe {getattr(recipe, '__name__', recipe)} crashed: {exc}")
                continue
            if code == 0:
                report.append(f"gate test '{test_path}' ({', '.join(hrs)}): "
                              f"recipe {getattr(recipe, '__name__', recipe)} exited 0 "
                              "against its violation fixture — the control did not fail")
    if report:
        print(f"HR-32 FAIL: {len(report)} finding(s)")
        for line in report:
            print(f"  - {line}")
        return 1
    print(f"HR-32 OK: {len(gates)} gate test(s) reject their violation fixtures.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
