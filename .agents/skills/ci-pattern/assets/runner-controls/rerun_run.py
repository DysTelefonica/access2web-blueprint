#!/usr/bin/env python3
# ci-pattern asset -- relanzar un run caido con un solo comando (DysTelefonica/team-skills#1630)
"""Relanza un run de GitHub Actions cancelado o caido, solo si sigue valiendo la pena.

Lo ejecuta el coordinador del repositorio con su propia App: el colaborador no
tiene ``actions:write`` y le pide el relanzamiento. Cerrar y reabrir la PR no es
la via de relanzar. Antes de relanzar comprueba:

- que el run pertenece a ``--repo``;
- que termino como ``cancelled``, ``failure`` o ``timed_out``;
- que su SHA sigue siendo la cabeza de su PR (o de su rama si no tiene PR): si
  no, hay que empujar, porque el rerun repetiria un rojo ya corregido.

Si todo cuadra ejecuta ``gh run rerun <id> --failed -R <repo>`` (``--all-jobs``
relanza el run entero). ``gh`` es ``$CI_PATTERN_GH`` o, si no, el del PATH (#1609).
La ultima linea dice el resultado. Salida: 0 relanzado; 1 GitHub rechazo el
rerun; 2 rechazado por una comprobacion o lectura fallida (falla cerrado, HR-3).
Solo stdlib.
"""
from __future__ import annotations

import argparse
import subprocess
import sys
from typing import NoReturn
from urllib.parse import quote

from safe_merge import GH_TIMEOUT, MergeFailClosedError, gh_json, run_command

RERUNNABLE = ("cancelled", "failure", "timed_out")
STALE = "ran on {0}, but the head of {1} is now {2}: push instead; a rerun repeats the old SHA"


class RefusedError(Exception):
    """A check refused the rerun; the message is the reason (exit 2)."""


def refuse(reason: str) -> NoReturn:
    """Stop with exit 2; the reason goes through a name so ``raise`` stays TRY003-clean."""
    raise RefusedError(reason)


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def read(path: str, what: str) -> dict:
    try:
        data = gh_json(path)
    except MergeFailClosedError as exc:
        reason = f"{what} could not be read: {exc}"
        raise RefusedError(reason) from exc
    if not isinstance(data, dict):
        refuse(f"{what} is not the API {what.split()[0]} shape")
    return data


def current_heads(repo: str, run: dict) -> list[tuple[str, str]]:
    """``(label, sha)`` of each PR of the run, or of its branch when it has none."""
    pulls = run.get("pull_requests") or []
    heads = []
    for pull in pulls:
        number = pull.get("number") if isinstance(pull, dict) else None
        detail = read(f"repos/{repo}/pulls/{number}", f"pull PR #{number}")
        heads.append((f"PR #{number}", (detail.get("head") or {}).get("sha")))
    if not heads:
        branch = run.get("head_branch")
        owner = (run.get("head_repository") or {}).get("full_name") or repo
        ref = read(f"repos/{owner}/git/ref/heads/{quote(str(branch), safe='/')}", f"ref of branch '{branch}'")
        heads.append((f"branch '{branch}'", (ref.get("object") or {}).get("sha")))
    return heads


def check(repo: str, run_id: int) -> dict:
    run = read(f"repos/{repo}/actions/runs/{run_id}", f"run {run_id} of {repo}")
    owner = (run.get("repository") or {}).get("full_name")
    if not isinstance(owner, str) or owner.lower() != repo.lower():
        refuse(f"run {run_id} belongs to {owner}, not to {repo}")
    status, conclusion = run.get("status"), run.get("conclusion")
    if status != "completed" or conclusion not in RERUNNABLE:
        refuse(f"run {run_id} is {status}/{conclusion}; only a completed run that "
                           f"ended {', '.join(RERUNNABLE[:-1])} or {RERUNNABLE[-1]} is rerun")
    sha = run.get("head_sha")
    if not isinstance(sha, str) or not sha:
        refuse(f"run {run_id} has no head SHA")
    heads = current_heads(repo, run)
    if not any(head == sha for _, head in heads):
        label, head = heads[0]
        refuse(f"run {run_id} " + STALE.format(str(sha)[:12], label, str(head)[:12]))
    return run


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Rerun a cancelled or failed run whose SHA is still the head (#1630).")
    parser.add_argument("--repo", required=True, help="owner/name of the repository")
    parser.add_argument("--run", required=True, type=int, help="workflow run id")
    parser.add_argument("--all-jobs", action="store_true", help="rerun every job, not only the failed ones")
    return parser.parse_args()


def main() -> int:
    _pin_output_encoding()
    args = parse_args()
    try:
        run = check(args.repo, args.run)
    except RefusedError as exc:
        print(f"REFUSED: {exc}")
        return 2
    argv = ["gh", "run", "rerun", str(args.run), *([] if args.all_jobs else ["--failed"]), "-R", args.repo]
    try:
        proc = run_command(argv, GH_TIMEOUT)
    except (OSError, subprocess.TimeoutExpired) as exc:
        print(f"FAILED: gh run rerun {args.run} could not run: {exc}")
        return 1
    if proc.returncode != 0:
        reason = " | ".join(line.strip() for line in proc.stderr.splitlines() if line.strip())
        print(f"FAILED: gh run rerun {args.run} exited {proc.returncode}: {reason}")
        return 1
    mode = "every job" if args.all_jobs else "--failed"
    print(f"RERUN: run {args.run} of {args.repo} ({run['conclusion']}, {run['head_sha'][:12]}) relaunched with {mode}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
