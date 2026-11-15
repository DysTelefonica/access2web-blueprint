#!/usr/bin/env python3
# ci-pattern asset -- esperar a la CI de una PR tras el push (DysTelefonica/team-skills#1502)
"""Espera a que la CI de una PR aparezca y termine, y dice el resultado.

`gh pr checks --watch` sale con 1 y `no checks reported` justo despues del
push, y `--required` responde `no required checks reported` mientras el
agregador no se ha lanzado. Cada agente escribia su propio bucle y algunos
daban la CI por verde con checks que aun no habian aparecido. Este helper
nunca devuelve exito con checks ausentes o pendientes:

- espera a que la cabeza tenga check runs (`--appear-seconds`);
- espera mientras haya check suites de GitHub Actions en cola o en curso
  dentro de esa misma ventana: GitHub crea la suite antes que sus runs, y un
  run ya verde no significa que no vengan mas (team-skills #1500);
- espera a cada `--check` nombrado (p. ej. el agregador `required`);
- con todo terminado, verde si cada run acabo en success, skipped o neutral.

Salida: 0 verde; 1 rojo, sin checks o fuera de tiempo; 2 uso o API (falla
cerrado, HR-3). Lee la API por HTTP (github_rest.py), sin `gh`: funciona
igual en Windows y en Linux. Solo stdlib.
"""
from __future__ import annotations

import argparse
import sys
import time
from datetime import UTC, datetime

from check_push_compliance import FailClosedError, fail, fetch_check_runs, gh_json

PASSING = {"success", "skipped", "neutral"}
ACTIONS_APP = "github-actions"


def _age_seconds(stamp: object, now: datetime) -> float | None:
    if not isinstance(stamp, str):
        return None
    try:
        created = datetime.fromisoformat(stamp.replace("Z", "+00:00"))
    except ValueError:
        return None
    return (now - created).total_seconds()


def starting_suites(repo: str, sha: str, appear_seconds: float) -> int:
    """Check suites de GitHub Actions aun en cola o en curso, creadas dentro de la ventana."""
    data = gh_json(f"repos/{repo}/commits/{sha}/check-suites?per_page=100")
    suites = data.get("check_suites") if isinstance(data, dict) else None
    if not isinstance(suites, list):
        return 0
    now = datetime.now(UTC)
    count = 0
    for suite in suites:
        if not isinstance(suite, dict) or suite.get("status") == "completed":
            continue
        if (suite.get("app") or {}).get("slug") != ACTIONS_APP:
            continue
        age = _age_seconds(suite.get("created_at"), now)
        if age is None or age <= appear_seconds:
            count += 1
    return count


def head_sha(repo: str, pr: int) -> str:
    pull = gh_json(f"repos/{repo}/pulls/{pr}")
    sha = (pull.get("head") or {}).get("sha") if isinstance(pull, dict) else None
    if not isinstance(sha, str):
        fail(f"PR #{pr} has no head sha in the API answer (fail closed)")
    return sha


def snapshot(repo: str, pr: int, sha: str, checks: list[str], appear_seconds: float) -> tuple[str, str]:
    """(estado, detalle): estado es waiting, green o red."""
    runs = [run for run in fetch_check_runs(repo, pr, sha) if isinstance(run, dict)]
    if not runs:
        return "waiting", f"no check runs yet for {sha[:12]}"
    by_name = {run.get("name"): run for run in runs}
    absent = [name for name in checks if name not in by_name]
    if absent:
        return "waiting", "required check(s) not reported yet: " + ", ".join(absent)
    pending = sorted(str(run.get("name")) for run in runs if run.get("status") != "completed")
    if pending:
        return "waiting", f"{len(pending)} check run(s) pending: " + ", ".join(pending)
    starting = starting_suites(repo, sha, appear_seconds)
    if starting:
        return "waiting", f"{starting} GitHub Actions check suite(s) still starting for {sha[:12]}"
    failed = sorted(f"{run.get('name')} ({run.get('conclusion')})" for run in runs
                    if run.get("conclusion") not in PASSING)
    if failed:
        return "red", ", ".join(failed)
    return "green", f"{len(runs)} check run(s) of {sha[:12]} passed"


def wait(args: argparse.Namespace) -> int:
    start = time.monotonic()
    sha = head_sha(args.repo, args.pr)
    last = ""
    seen_runs = False
    while True:
        current = head_sha(args.repo, args.pr)
        if current != sha:
            print(f"NOTE: the head of PR #{args.pr} moved from {sha[:12]} to {current[:12]}; "
                  "waiting on the new head")
            sha, seen_runs, last = current, False, ""
        state, detail = snapshot(args.repo, args.pr, sha, args.check, args.appear_seconds)
        if state == "green":
            print(f"GREEN: {detail}")
            return 0
        if state == "red":
            print(f"RED: {detail}")
            return 1
        seen_runs = seen_runs or not detail.startswith("no check runs yet")
        if detail != last:
            print(f"WAITING: {detail}", flush=True)
            last = detail
        elapsed = time.monotonic() - start
        if not seen_runs and elapsed >= args.appear_seconds:
            print(f"NOT STARTED: no check runs appeared for {sha[:12]} within {args.appear_seconds:g} s; "
                  "a base change does not start CI: close and reopen the PR or push a commit")
            return 1
        if elapsed >= args.timeout_seconds:
            print(f"TIMEOUT: {detail} after {args.timeout_seconds:g} s")
            return 1
        time.sleep(args.poll_seconds)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Wait for a PR's CI to appear and finish.")
    parser.add_argument("--repo", required=True, help="OWNER/REPO")
    parser.add_argument("--pr", required=True, type=int, help="pull request number")
    parser.add_argument("--check", action="append", default=[],
                        help="check run that must be reported (repeatable), e.g. the 'required' aggregator")
    parser.add_argument("--poll-seconds", type=float, default=30.0)
    parser.add_argument("--appear-seconds", type=float, default=600.0,
                        help="how long the first check runs may take to appear")
    parser.add_argument("--timeout-seconds", type=float, default=3600.0)
    args = parser.parse_args(argv)
    if args.pr <= 0 or min(args.poll_seconds, args.appear_seconds, args.timeout_seconds) < 0:
        parser.error("--pr must be positive and the seconds must not be negative")
    return args


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    args = parse_args(sys.argv[1:] if argv is None else argv)
    try:
        return wait(args)
    except FailClosedError as exc:
        print(f"ERROR: {exc}")
        return 2


if __name__ == "__main__":
    sys.exit(main())
