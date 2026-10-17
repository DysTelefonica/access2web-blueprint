#!/usr/bin/env python3
# ci-pattern asset — governed merge (DysTelefonica/team-skills#303)
"""Merge gobernado: el unico camino de merge sobre una rama protegida.

El workflow que corre en el runner propio invoca este asset con el numero
de PR; verifica el contrato antes de fusionar: los ``required_checks`` del
contrato en verde sobre la cabeza del PR, exactamente una etiqueta
``type:*``, el tamano dentro del presupuesto de revision (``--max-lines``,
400 por defecto) y ``Closes #N`` solo en el punta de cadena (un PR
``chain:partial`` lo lleva en la violacion). Si algo falta no fusiona y
devuelve 1 con el hallazgo; el entorno insalvable falla cerrado con 2
(HR-3). Con todo en verde fusiona por la API con la identidad del bot que
ejecuta el workflow y devuelve 0. Runner-enforced (HR-34); lo que no pase
por aqui lo detecta el detective post-push (HR-49). Solo stdlib; cada
llamada a ``gh api`` lleva timeout.
"""
import argparse
import json
import re
import sys

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


def gh_merge(repo: str, number: object, sha: str) -> str:
    # El sha verificado viaja con el merge: si el head cambia entre la
    # comprobación y el PUT, la API responde 409 y el fallo es cerrado (2) —
    # nunca se fusiona una cabeza que nadie verificó (TOCTOU).
    try:
        text, _ = call("PUT", f"repos/{repo}/pulls/{number}/merge",
                       {"commit_title": f"Merge PR #{number} (governed merge)", "sha": sha})
    except GitHubApiError as exc:
        fail(f"the merge call failed (fail closed): {exc}", exc)
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
    findings += check_verdict(args.repo, args.pr, pull["head"]["sha"], required)
    type_labels = sorted(name for name in labels if name.startswith("type:"))
    if len(type_labels) != 1:
        findings.append(f"expected exactly one type:* label, found {type_labels or 'none'}")
    size = pull["additions"] + pull["deletions"]
    if size > args.max_lines and "size:exception" not in labels:
        findings.append(f"PR size {size} lines exceeds the review budget of {args.max_lines} "
                        "without the 'size:exception' label applied by the maintainer (A1)")
    if "chain:partial" in labels and CLOSING.search(body):
        findings.append("Closes in an intermediate PR: only the chain tip may carry closing words")
    base_ref = pull["base"]["ref"]
    if base_ref not in protected_branches(contract, None) and "chain:partial" not in labels:
        findings.append(f"PR #{args.pr} targets '{base_ref}': not a branch protected by the "
                        "contract and the PR is not a chain intermediate (chain:partial)")
    return findings


def run(args: argparse.Namespace) -> int:
    if args.repo.count("/") != 1 or not args.pr.isdigit() or int(args.pr) < 1:
        fail("--repo must be OWNER/REPO and --pr a positive number")
    if args.max_lines < 1:
        fail("--max-lines must be positive")
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


def main() -> int:
    parser = argparse.ArgumentParser(description="governed merge vs the host contract")
    parser.add_argument("--repo", required=True, help="OWNER/REPO")
    parser.add_argument("--pr", required=True, help="pull request number")
    parser.add_argument("--contract", required=True)
    parser.add_argument("--max-lines", type=int, default=400)
    args = parser.parse_args()
    try:
        code = run(args)
    except FailClosedError as exc:
        print(f"fail-closed: {exc}", file=sys.stderr)
        return 2
    return code


if __name__ == "__main__":
    sys.exit(main())
