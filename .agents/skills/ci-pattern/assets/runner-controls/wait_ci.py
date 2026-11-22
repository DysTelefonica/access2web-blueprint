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
from pathlib import Path
from typing import NamedTuple

sys.path.insert(0, str(Path(__file__).resolve().parent))
from advisory_checks import (  # noqa: E402
    CONTRACT_PATH,
    AdvisoryContractError,
    read_contract,
    split_runs,
)
from check_push_compliance import FailClosedError, fail, fetch_check_runs, gh_json  # noqa: E402

PASSING = {"success", "skipped", "neutral"}
ACTIONS_APP = "github-actions"


class WaitPolicy(NamedTuple):
    """Como se espera: la ventana de aparicion y los checks que no bloquean (#1762)."""

    appear_seconds: float
    advisory: tuple[str, ...]
    protected: frozenset[str]


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


def snapshot(repo: str, pr: int, sha: str, checks: list[str], policy: WaitPolicy) -> tuple[str, str]:
    """(estado, detalle): estado es waiting, green o red.

    #1762: los checks advisory no cuentan para el veredicto; un check requerido
    (por `--check` o por el contrato) nunca queda tapado por un patron."""
    runs = [run for run in fetch_check_runs(repo, pr, sha) if isinstance(run, dict)]
    by_name = {run.get("name"): run for run in runs}
    absent = [name for name in checks if name not in by_name]
    blocking, ignored = split_runs(runs, policy.advisory, required=policy.protected)
    if absent and runs:
        return "waiting", "required check(s) not reported yet: " + ", ".join(absent)
    if not blocking:
        # No hay checks, o solo hay advisory: no hay nada que juzgar todavia, no es verde.
        detalle = "no check runs yet" if not runs else "no blocking check run yet"
        return "waiting", f"{detalle} for {sha[:12]}"
    pending = sorted(str(run.get("name")) for run in blocking if run.get("status") != "completed")
    if pending:
        return "waiting", f"{len(pending)} check run(s) pending: " + ", ".join(pending)
    starting = starting_suites(repo, sha, policy.appear_seconds)
    if starting:
        return "waiting", f"{starting} GitHub Actions check suite(s) still starting for {sha[:12]}"
    failed = sorted(f"{run.get('name')} ({run.get('conclusion')})" for run in blocking
                    if run.get("conclusion") not in PASSING)
    if failed:
        return "red", ", ".join(failed)
    detail = f"{len(blocking)} check run(s) of {sha[:12]} passed"
    if ignored:
        detail += "; advisory check(s) ignored: " + ", ".join(
            sorted(str(run.get("name")) for run in ignored))
    return "green", detail


def wait(args: argparse.Namespace) -> int:
    start = time.monotonic()
    try:
        declared = read_contract(args.contract)
    except AdvisoryContractError as exc:
        raise FailClosedError(str(exc)) from exc
    # Un check requerido nunca queda tapado por un patron advisory: lo es el que
    # nombra la linea de comandos o el que declara el contrato del consumer. Los
    # del contrato protegen del patron, pero no se esperan: los espera el guard
    # que mergea, no este.
    protected = frozenset(args.check) | frozenset(declared.required)
    sha = head_sha(args.repo, args.pr)
    last = ""
    seen_runs = False
    while True:
        current = head_sha(args.repo, args.pr)
        if current != sha:
            print(f"NOTE: the head of PR #{args.pr} moved from {sha[:12]} to {current[:12]}; "
                  "waiting on the new head")
            sha, seen_runs, last = current, False, ""
        state, detail = snapshot(args.repo, args.pr, sha, args.check,
                       WaitPolicy(args.appear_seconds, declared.patterns, protected))
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
                  "a base change does not start CI: push an empty commit")
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
    parser.add_argument("--contract", default=CONTRACT_PATH,
                        help="host-contract del consumer, donde se declaran los checks advisory "
                             f"(por defecto {CONTRACT_PATH}; si no existe, no hay advisory)")
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
