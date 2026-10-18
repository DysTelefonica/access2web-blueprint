#!/usr/bin/env python3
# ci-pattern asset — governed merge (DysTelefonica/team-skills#303)
"""Merge gobernado: el unico camino de merge sobre una rama protegida.

El workflow que corre en el runner propio invoca este asset con el numero
de PR; verifica el contrato antes de fusionar: los ``required_checks`` del
contrato en verde sobre la cabeza del PR, exactamente una etiqueta
``type:*`` y ``Closes #N`` solo en el punta de cadena (un PR
``chain:partial`` lo lleva en la violacion). Un check requerido que aun no
existe o no ha terminado se espera hasta ``--wait-seconds`` (#787). El tamano
no se repite aqui: lo mide ``pr-size`` (excluye lo vendorizado), y ``required``
ya lo incluye; ``--max-lines`` se acepta e ignora por compatibilidad. Si algo falta no fusiona y
devuelve 1 con el hallazgo; el entorno insalvable falla cerrado con 2
(HR-3). Con todo en verde fusiona por la API con la identidad del bot que
ejecuta el workflow y devuelve 0. Un PR atrasado respecto de su base se bloquea antes (1): el arbol a fusionar
no seria el que reviso el CI (#787) y, si ambos lados cambiaron el mismo fichero
de ``.github/workflows/``, ``GITHUB_TOKEN`` no tiene el permiso ``workflows`` (#783).
Runner-enforced (HR-34); lo que no pase
por aqui lo detecta el detective post-push (HR-49). Solo stdlib; cada
llamada a ``gh api`` lleva timeout.
"""
import argparse
import json
import re
import sys
import time

from check_push_compliance import (
    FailClosedError,
    check_verdict,
    fail,
    gh_json,
    load_json,
    protected_branches,
    required_checks,
)
from github_rest import GitHubApiError, call

CLOSING = re.compile(r"\b(closes|fixes|resolves)\s+#?\d+", re.IGNORECASE)
WORKFLOWS_DIR = ".github/workflows/"
FORBIDDEN = 403
WAITABLE = (": missing", ": pending")  # el CI aun no publico o no termino; un fallo no se espera
DEFAULT_WAIT, DEFAULT_POLL = 300.0, 10.0
WORKFLOWS_HINT = ("the token lacks the 'workflows' permission, which GITHUB_TOKEN cannot "
                  "have: GitHub refuses a merge that creates new workflow-file content; "
                  "update the PR branch onto its base or merge it as the maintainer")


def gh_merge(repo: str, number: object, sha: str) -> str:
    # El sha verificado viaja con el merge: si el head cambia entre la
    # comprobación y el PUT, la API responde 409 y el fallo es cerrado (2) —
    # nunca se fusiona una cabeza que nadie verificó (TOCTOU).
    try:
        text, _ = call("PUT", f"repos/{repo}/pulls/{number}/merge",
                       {"commit_title": f"Merge PR #{number} (governed merge)", "sha": sha})
    except GitHubApiError as exc:
        hint = f" ({WORKFLOWS_HINT})" if exc.status == FORBIDDEN and "workflow" in str(exc).lower() else ""
        fail(f"the merge call failed (fail closed): {exc}{hint}", exc)
    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        fail(f"the merge API returned no valid JSON body: {exc}", exc)
    if not isinstance(data, dict) or data.get("merged") is not True \
            or not isinstance(data.get("sha"), str):
        fail(f"the merge response does not report a merged PR: {text.strip()!r}")
    return data["sha"]


def is_pull_shape(pull: object) -> bool:
    return (isinstance(pull, dict) and isinstance(pull.get("head"), dict)
            and isinstance(pull["head"].get("sha"), str) and isinstance(pull.get("labels"), list)
            and isinstance(pull.get("base"), dict) and isinstance(pull["base"].get("ref"), str)
            and isinstance(pull.get("additions"), int) and isinstance(pull.get("deletions"), int))


def compare(repo: str, left: str, right: str) -> dict:
    """Respuesta del compare de la API, validada en forma (`files` y `behind_by`)."""
    data = gh_json(f"repos/{repo}/compare/{left}...{right}")
    files = data.get("files") if isinstance(data, dict) else None
    if not isinstance(files, list) or not all(
            isinstance(item, dict) and isinstance(item.get("filename"), str) for item in files):
        fail(f"compare {left}...{right} is not the API compare shape")
    return data


def paths_of(data: dict) -> set[str]:
    return {name for item in data["files"] for name in (item["filename"], item.get("previous_filename"))
            if isinstance(name, str)}


def freshness_findings(args: argparse.Namespace, pull: dict) -> list[str]:
    """La rama debe estar al dia con su base: el arbol a fusionar es el que reviso el CI.

    Con la base movida (`behind_by` > 0) el merge produce un arbol que ninguna
    ejecucion de CI vio, y el deploy lo rechaza («main moved»): se bloquea antes,
    con la accion a tomar (#787). Si ademas ambos lados cambiaron el mismo
    workflow, el merge crearia contenido de workflow nuevo y GITHUB_TOKEN no
    puede (sin permiso `workflows`, #783): GitHub respondería 403 al PUT.
    """
    base, head = pull["base"]["ref"], pull["head"]["sha"]
    ahead = compare(args.repo, base, head)
    behind = ahead.get("behind_by")
    if not isinstance(behind, int) or isinstance(behind, bool) or behind < 0:
        fail(f"compare {base}...{head} has no valid 'behind_by'")
    if behind == 0:
        return []
    mine = {name for name in paths_of(ahead) if name.startswith(WORKFLOWS_DIR)}
    shared = sorted(mine & paths_of(compare(args.repo, head, base))) if mine else []
    action = f"update the PR branch onto '{base}', let CI go green and run the governed merge again"
    if shared:
        return [f"PR #{args.pr} and '{base}' both changed {', '.join(shared)}: merging would create "
                "new workflow content, which the workflow token cannot do (no 'workflows' permission); "
                + action]
    return [f"PR #{args.pr} is behind '{base}' by {behind} commit(s): the merged tree would differ "
            "from the tree CI reviewed and the deploy evidence refuses it; " + action]


def verdict_findings(args: argparse.Namespace, sha: str, required: list[str]) -> list[str]:
    """Veredicto de los checks requeridos, esperando (acotado) a los que aun no existen.

    Un merge lanzado antes de que el CI publique `required` no es un contrato
    incumplido, es una carrera (#787): se reintenta cada `--poll-seconds` hasta
    `--wait-seconds`. Un check en fallo, o cualquier otro hallazgo, no se espera.
    """
    deadline = time.monotonic() + args.wait_seconds
    while True:
        findings = check_verdict(args.repo, args.pr, sha, required)
        waiting = [item for item in findings if item.endswith(WAITABLE)]
        if not waiting or len(waiting) != len(findings) or time.monotonic() >= deadline:
            return findings
        print(f"waiting for required checks: {'; '.join(waiting)}", file=sys.stderr)
        time.sleep(min(args.poll_seconds, max(0.0, deadline - time.monotonic())))


def gate_findings(args: argparse.Namespace, pull: dict, contract: dict,
                  required: list[str]) -> list[str]:
    body = pull.get("body") if isinstance(pull.get("body"), str) else ""
    labels = {label.get("name") for label in pull["labels"]
              if isinstance(label, dict) and isinstance(label.get("name"), str)}
    findings = []
    if pull.get("state") != "open":
        findings.append(f"PR #{args.pr} is not open (state: {pull.get('state')!r})")
    if pull.get("draft") is True:
        findings.append(f"PR #{args.pr} is a draft: the governed path merges ready PRs only")
    findings += verdict_findings(args, pull["head"]["sha"], required)
    type_labels = sorted(name for name in labels if name.startswith("type:"))
    if len(type_labels) != 1:
        findings.append(f"expected exactly one type:* label, found {type_labels or 'none'}")
    if "chain:partial" in labels and CLOSING.search(body):
        findings.append("Closes in an intermediate PR: only the chain tip may carry closing words")
    base_ref = pull["base"]["ref"]
    if base_ref not in protected_branches(contract, None) and "chain:partial" not in labels:
        findings.append(f"PR #{args.pr} targets '{base_ref}': not a branch protected by the "
                        "contract and the PR is not a chain intermediate (chain:partial)")
    if not findings:
        findings += freshness_findings(args, pull)
    return findings


def run(args: argparse.Namespace) -> int:
    if args.repo.count("/") != 1 or not args.pr.isdigit() or int(args.pr) < 1:
        fail("--repo must be OWNER/REPO and --pr a positive number")
    if args.wait_seconds < 0 or args.poll_seconds <= 0:
        fail("--wait-seconds must be zero or more and --poll-seconds positive")
    contract = load_json(args.contract, "contract")
    if not isinstance(contract, dict):
        fail("contract must be a JSON object")
    required = required_checks(contract)
    pull = gh_json(f"repos/{args.repo}/pulls/{args.pr}")
    if not is_pull_shape(pull):
        fail(f"PR #{args.pr} is not the API pull shape")
    findings = gate_findings(args, pull, contract, required)
    if findings:
        for finding in findings:
            print(f"  BLOCKED  {finding}")
        print(f"VERDICT: PR #{args.pr} not merged: contract not satisfied")
        return 1
    sha = gh_merge(args.repo, args.pr, pull["head"]["sha"])
    print(f"MERGED: PR #{args.pr} merged by the governed path as {sha}")
    return 0


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main() -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description="governed merge vs the host contract")
    parser.add_argument("--repo", required=True, help="OWNER/REPO")
    parser.add_argument("--pr", required=True, help="pull request number")
    parser.add_argument("--contract", required=True)
    parser.add_argument("--max-lines", type=int, default=None,
                        help="deprecated and ignored: pr-size owns the review budget (#787)")
    parser.add_argument("--wait-seconds", type=float, default=DEFAULT_WAIT,
                        help="how long to wait for required checks not yet published")
    parser.add_argument("--poll-seconds", type=float, default=DEFAULT_POLL)
    args = parser.parse_args()
    try:
        code = run(args)
    except FailClosedError as exc:
        print(f"fail-closed: {exc}", file=sys.stderr)
        return 2
    return code


if __name__ == "__main__":
    sys.exit(main())
