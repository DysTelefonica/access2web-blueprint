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
import subprocess
import sys
from pathlib import Path

from check_push_compliance import (FailClosed, check_verdict, gh_json, load_json,
                                   protected_branches, required_checks)

# HR-8 vive UNA sola vez: el merge gobernado importa el validador del gate
# de contrato de PR en proceso (nunca lo copia ni invoca la CLI).
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "pr-contract"))
from check_pr_contract import size_exception_reason  # noqa: E402

GH_TIMEOUT = 30
CLOSING = re.compile(r"\b(closes|fixes|resolves)\s+#?\d+", re.IGNORECASE)


def gh_merge(repo: str, number: object, sha: str) -> str:
    # El sha verificado viaja con el merge: si el head cambia entre la
    # comprobación y el PUT, la API responde 409 y el fallo es cerrado (2) —
    # nunca se fusiona una cabeza que nadie verificó (TOCTOU).
    try:
        proc = subprocess.run(["gh", "api", "-X", "PUT", f"repos/{repo}/pulls/{number}/merge",
                               "-F", f"commit_title=Merge PR #{number} (governed merge)",
                               "-F", f"sha={sha}"],
                              capture_output=True, text=True, timeout=GH_TIMEOUT)
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise FailClosed(f"the merge call failed: {exc}") from exc
    if proc.returncode != 0:
        raise FailClosed(f"the merge call failed (fail closed): {proc.stderr.strip()}")
    try:
        data = json.loads(proc.stdout)
    except json.JSONDecodeError as exc:
        raise FailClosed(f"gh returned no valid JSON body: {exc}") from exc
    if not isinstance(data, dict) or data.get("merged") is not True or not isinstance(data.get("sha"), str):
        raise FailClosed(f"the merge response does not report a merged PR: {proc.stdout.strip()!r}")
    return data["sha"]


def main() -> int:
    parser = argparse.ArgumentParser(description="governed merge vs the host contract")
    parser.add_argument("--repo", required=True, help="OWNER/REPO")
    parser.add_argument("--pr", required=True, help="pull request number")
    parser.add_argument("--contract", required=True)
    parser.add_argument("--max-lines", type=int, default=400)
    args = parser.parse_args()
    try:
        if args.repo.count("/") != 1 or not args.pr.isdigit() or int(args.pr) < 1:
            raise FailClosed("--repo must be OWNER/REPO and --pr a positive number")
        if args.max_lines < 1:
            raise FailClosed("--max-lines must be positive")
        contract = load_json(args.contract, "contract")
        if not isinstance(contract, dict):
            raise FailClosed("contract must be a JSON object")
        required = required_checks(contract)
        pull = gh_json(f"repos/{args.repo}/pulls/{args.pr}")
        if not isinstance(pull, dict) or not isinstance(pull.get("head"), dict) \
                or not isinstance(pull["head"].get("sha"), str) or not isinstance(pull.get("labels"), list) \
                or not isinstance(pull.get("base"), dict) or not isinstance(pull["base"].get("ref"), str) \
                or not isinstance(pull.get("additions"), int) or not isinstance(pull.get("deletions"), int):
            raise FailClosed(f"PR #{args.pr} is not the API pull shape")
        body = pull.get("body") if isinstance(pull.get("body"), str) else ""
        labels = {l.get("name") for l in pull["labels"] if isinstance(l, dict) and isinstance(l.get("name"), str)}
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
        if size > args.max_lines and size_exception_reason(body) is None:
            findings.append(f"PR size {size} lines exceeds the review budget of {args.max_lines} "
                            "without a valid size-exception-reason field (HR-8)")
        if "chain:partial" in labels and CLOSING.search(body):
            findings.append("Closes in an intermediate PR: only the chain tip may carry closing words")
        base_ref = pull["base"]["ref"]
        if base_ref not in protected_branches(contract, None) and "chain:partial" not in labels:
            findings.append(f"PR #{args.pr} targets '{base_ref}': not a branch protected by the "
                            "contract and the PR is not a chain intermediate (chain:partial)")
        if findings:
            for finding in findings:
                print(f"  BLOCKED  {finding}")
            print(f"VERDICT: PR #{args.pr} not merged: contract not satisfied")
            return 1
        sha = gh_merge(args.repo, args.pr, pull["head"]["sha"])
        print(f"MERGED: PR #{args.pr} merged by the governed path as {sha}")
        return 0
    except FailClosed as exc:
        print(f"fail-closed: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
