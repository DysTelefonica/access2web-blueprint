#!/usr/bin/env python3
# ci-pattern asset -- relanzar la canalizacion nocturna sobre main (DysTelefonica/team-skills#1627, slice 3)
"""Relanza el workflow nocturno sobre main, sin duplicarlo.

Antes de lanzar lee la cabeza de main y las ejecuciones nocturnas de ese SHA: si ya hay
una activa (``queued``, ``in_progress``, ``waiting``...), sale con SKIP y no relanza. Así
dos llamadas seguidas dan un solo ``gh workflow run``. ``--dry-run`` no llama a gh.

Salida: la ultima linea dice el resultado. 0 relanzado, SKIP o ensayo; 1 GitHub rechazo el
relanzamiento; 2 REFUSED (repo mal formado o estado ilegible: no se relanza). El ``gh`` es
``$CI_PATTERN_GH`` o, si no, el del PATH (#1609). Solo stdlib.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys

WORKFLOW = "nightly-pipeline.yml"
ACTIVE = ("queued", "in_progress", "waiting", "requested", "pending", "action_required")
REPO_SHAPE = re.compile(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+")


def valid_repo(repo):
    """OWNER/NAME con caracteres admitidos, sin salto final y sin segmentos solo de puntos."""
    if not REPO_SHAPE.fullmatch(repo):
        return False
    return not any(set(segment) <= {"."} for segment in repo.split("/"))
GH_TIMEOUT = 60


class ReadError(Exception):
    """A read of the nightly state failed or had another shape (refuse, never dispatch)."""


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def gh_command() -> list[str]:
    """``gh`` a ejecutar: ``$CI_PATTERN_GH`` si esta fijada (una ruta ``.py`` corre con este Python), si no el del PATH."""
    explicit = os.environ.get("CI_PATTERN_GH")
    if explicit and explicit.endswith(".py"):
        return [sys.executable, explicit]
    return [explicit or "gh"]


def gh_api(path: str) -> object:
    try:
        proc = subprocess.run([*gh_command(), "api", path], capture_output=True, text=True, check=False,  # noqa: S603 - fixed argv
                              encoding="utf-8", timeout=GH_TIMEOUT)
    except (OSError, subprocess.TimeoutExpired) as exc:
        reason = f"gh api could not run: {exc}"
        raise ReadError(reason) from exc
    if proc.returncode != 0:
        reason = f"gh api {path} failed ({proc.returncode})"
        raise ReadError(reason)
    try:
        return json.loads(proc.stdout)
    except ValueError as exc:
        reason = f"gh api {path} did not return JSON"
        raise ReadError(reason) from exc


def active_run(repo: str, sha: str) -> dict | None:
    """La ejecucion nocturna activa para ``sha``, o None."""
    listed = gh_api(f"repos/{repo}/actions/workflows/{WORKFLOW}/runs?head_sha={sha}&per_page=100")
    runs = listed.get("workflow_runs") if isinstance(listed, dict) else None
    if not isinstance(runs, list):
        reason = "the runs list is not in the API shape"
        raise ReadError(reason)
    return next((run for run in runs if isinstance(run, dict) and run.get("status") in ACTIVE), None)


def head_sha(repo: str) -> str:
    """La cabeza de main; si no se puede leer, levanta ReadError."""
    ref = gh_api(f"repos/{repo}/git/ref/heads/main")
    obj = ref.get("object") if isinstance(ref, dict) else None
    sha = obj.get("sha") if isinstance(obj, dict) else None
    if not isinstance(sha, str) or not sha:
        reason = "main has no head SHA"
        raise ReadError(reason)
    return sha


def dispatch(repo: str) -> tuple[int, str]:
    """Lanza el workflow sobre main (una sola llamada a gh workflow run)."""
    argv = [*gh_command(), "workflow", "run", WORKFLOW, "--repo", repo, "--ref", "main"]
    try:
        proc = subprocess.run(argv, capture_output=True, text=True, check=False,  # noqa: S603 - fixed argv
                              encoding="utf-8", timeout=GH_TIMEOUT)
    except (OSError, subprocess.TimeoutExpired) as exc:
        return 1, f"FAILED: gh workflow run could not run: {exc}"
    if proc.returncode != 0:
        reason = " | ".join(line.strip() for line in proc.stderr.splitlines() if line.strip())
        return 1, f"FAILED: gh workflow run {WORKFLOW} exited {proc.returncode}: {reason}"
    return 0, f"REDISPATCHED: {WORKFLOW} on main of {repo}; follow it with gh run list --workflow {WORKFLOW}"


def redispatch(repo: str, dry_run: bool) -> tuple[int, str]:
    """Devuelve (codigo, linea final): SKIP si ya hay una nocturna activa para main, si no relanza."""
    if not valid_repo(repo):
        return 2, f"REFUSED: --repo must be owner/name, got {repo!r}"
    if dry_run:
        return 0, f"DRY-RUN: would dispatch {WORKFLOW} on main of {repo} (no gh call made)"
    try:
        sha = head_sha(repo)
        running = active_run(repo, sha)
    except ReadError as exc:
        return 2, f"REFUSED: could not read the nightly state, not dispatching: {exc}"
    if running is not None:
        return 0, (f"SKIP: nightly run {running.get('id')} for {sha[:12]} is {running.get('status')}; "
                   f"not dispatching again")
    return dispatch(repo)


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description="Relanza la canalizacion nocturna sobre main (#1627, slice 3).")
    parser.add_argument("--repo", required=True, help="owner/name")
    parser.add_argument("--dry-run", action="store_true", help="no llama a gh; solo dice lo que haria")
    args = parser.parse_args(sys.argv[1:] if argv is None else argv)
    code, line = redispatch(args.repo, args.dry_run)
    print(line)
    return code


if __name__ == "__main__":
    sys.exit(main())
