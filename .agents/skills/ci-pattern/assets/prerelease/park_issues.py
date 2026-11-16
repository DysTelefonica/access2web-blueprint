#!/usr/bin/env python3
# ci-pattern asset -- aparcar issues durante una noche en rojo (DysTelefonica/team-skills#1627, slice 5)
"""Aparca las issues normales en una noche roja (ver park-issues.md).

Habla con la API de GitHub por `runner-controls/github_rest.py` (urllib, stdlib):
el runner del consumer no tiene el CLI `gh` (#1691).
"""
from __future__ import annotations

import argparse
import importlib.util
import json
import sys
import time
import urllib.parse
from collections.abc import Callable
from pathlib import Path
from typing import NamedTuple

FIX_LABEL = "prerelease:night-red"
PARK_LABEL = "prerelease:parked"
NIGHTLY_WORKFLOW = "nightly"
RED = ("failure", "cancelled", "timed_out")
DEFAULT_LIST_LIMIT = 5000
DEPLOY_ENVIRONMENT = "production"
ISSUE_PAGE = 100  # el maximo de `per_page` que acepta la API
LINK_PARTS = 2  # <url>; rel="..."


class ReadError(Exception):
    """A GitHub read failed or returned another shape (fail closed)."""


def _load_github_rest():
    """Carga el cliente compartido `runner-controls/github_rest.py` por ruta (#1592).

    Sin él, el gate no arranca: sale con 2 (HR-3), como `issue-spec`.
    """
    rest_path = Path(__file__).resolve().parents[1] / "runner-controls" / "github_rest.py"
    if not rest_path.is_file():
        return None
    spec = importlib.util.spec_from_file_location("park_issues_github_rest", str(rest_path))
    if spec is None or spec.loader is None:
        return None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


_REST = _load_github_rest()
if _REST is None:
    print("park-issues: runner-controls/github_rest.py no esta en el catalogo "
          "(HR-3 fail-closed)", file=sys.stderr)
    sys.exit(2)


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8: la salida no debe depender del locale."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def night_state(run: object) -> str:
    """``green`` (success), ``red`` (failure, cancelled, timed_out) o ``unknown`` (no terminada o ilegible)."""
    if not isinstance(run, dict) or run.get("status") != "completed":
        return "unknown"
    conclusion = run.get("conclusion")
    if conclusion == "success":
        return "green"
    if conclusion in RED:
        return "red"
    return "unknown"


def workflow_stem(run: dict) -> str:
    """Nombre del workflow de la ejecucion: el de su fichero (``nightly.yml`` -> ``nightly``)."""
    path = run.get("path")
    if not isinstance(path, str) or not path:
        return ""
    return Path(path).stem


def _read_page(path: str) -> tuple[object, str]:
    """Un GET con reintento: (cuerpo JSON, cabecera Link). Un fallo o un cuerpo que no es JSON levanta ReadError."""
    try:
        text, link = _REST.call_con_reintento("GET", path)
    except _REST.GitHubApiError as exc:
        reason = f"GET {path} failed: {exc}"
        raise ReadError(reason) from exc
    try:
        return json.loads(text), link
    except ValueError as exc:
        reason = f"GET {path} did not return JSON"
        raise ReadError(reason) from exc


def _read_json(path: str) -> object:
    """Un GET con reintento; el cuerpo parseado."""
    return _read_page(path)[0]


def _next_page(link: str | None) -> str | None:
    """La ruta (con query) de la pagina siguiente segun la cabecera Link; None si no hay."""
    for part in (link or "").split(","):
        segments = part.strip().split(";")
        if len(segments) == LINK_PARTS and segments[1].strip() == 'rel="next"':
            target = urllib.parse.urlsplit(segments[0].strip().strip("<>"))
            return f"{target.path}?{target.query}" if target.query else target.path
    return None


def read_night(repo: str, run_id: int) -> object:
    try:
        return _read_json(f"repos/{repo}/actions/runs/{run_id}")
    except ReadError:
        return None


def deploy_ok(repo: str, environment: str, day: str) -> bool:
    """Hay un deploy a ``environment`` creado el dia UTC ``day`` cuyo estado mas reciente es success."""
    wanted = urllib.parse.quote(environment, safe="")
    deployments = _read_json(f"repos/{repo}/deployments?environment={wanted}&per_page=100")
    if not isinstance(deployments, list):
        reason = "the deployments did not return a list"
        raise ReadError(reason)
    for dep in deployments:
        if not isinstance(dep, dict) or dep.get("environment") != environment:
            continue
        if str(dep.get("created_at", ""))[:10] != day:
            continue
        statuses = _read_json(f"repos/{repo}/deployments/{dep.get('id')}/statuses?per_page=1")
        if isinstance(statuses, list) and statuses and isinstance(statuses[0], dict) \
                and statuses[0].get("state") == "success":
            return True
    return False


def list_issues(repo: str, limit: int) -> list[dict]:
    """Las issues del repo (no los PRs) hasta ``limit``; si hay mas, la lista esta incompleta.

    La API mezcla issues y PRs: se pagina hasta tener ``limit`` issues o hasta que
    la cabecera Link diga que no hay mas paginas.
    """
    listed: list = []
    path = f"repos/{repo}/issues?state=all&per_page={ISSUE_PAGE}"
    while True:
        chunk, link = _read_page(path)
        if not isinstance(chunk, list):
            reason = "the issue list did not return a list"
            raise ReadError(reason)
        listed += [item for item in chunk
                   if isinstance(item, dict) and isinstance(item.get("number"), int)
                   and "pull_request" not in item]
        if len(listed) > limit:
            reason = f"the issue list returned more than {limit} issues: the list is incomplete"
            raise ReadError(reason)
        nxt = _next_page(link)
        if nxt is None:
            break
        path = nxt
    return listed


def has_label(issue: dict, name: str) -> bool:
    return any(isinstance(label, dict) and label.get("name") == name for label in issue.get("labels") or [])


def park_targets(issues: list[dict], fix: str, park: str) -> list[int]:
    return [i["number"] for i in issues
            if str(i.get("state", "")).upper() == "OPEN" and not has_label(i, fix) and not has_label(i, park)]


def unpark_targets(issues: list[dict], park: str) -> list[int]:
    """Issues de cualquier estado que aún tienen la etiqueta de aparcado."""
    return [i["number"] for i in issues if has_label(i, park)]


def edit_label(repo: str, number: int, add: bool, label: str) -> bool:
    """Una edición de etiqueta; True si la API la acepta.

    ``add`` True la pone (``POST .../labels``), False la quita
    (``DELETE .../labels/<nombre>``).
    """
    if add:
        path = f"repos/{repo}/issues/{number}/labels"
        method, body = "POST", {"labels": [label]}
    else:
        path = f"repos/{repo}/issues/{number}/labels/{urllib.parse.quote(label, safe='')}"
        method, body = "DELETE", None
    try:
        _REST.call(method, path, body)
    except _REST.GitHubApiError:
        return False
    return True


class Action(NamedTuple):
    """Un comando: el estado nocturno que lo permite, la edicion, su verbo y sus issues."""
    wanted: str
    add: bool
    verb: str
    targets_of: Callable[[list, str, str], list[int]]
    needs_deploy: bool = False


def _refusal(args: argparse.Namespace, action: Action) -> str | None:
    run = read_night(args.repo, args.night_run)
    if run is None:
        return f"HOLD: could not read nightly run {args.night_run}: {action.verb} needs a readable run"
    if workflow_stem(run) != args.workflow:
        return (f"HOLD: run {args.night_run} is from workflow {workflow_stem(run) or 'unknown'}, "
                f"not {args.workflow}: {action.verb} only on the nightly workflow")
    state = night_state(run)
    if state != action.wanted:
        return (f"HOLD: nightly run {args.night_run} is {state}, not {action.wanted}: "
                f"{action.verb} only on a {action.wanted} night")
    if action.needs_deploy:
        return _deploy_refusal(args, action.verb)
    return None


def _deploy_refusal(args: argparse.Namespace, verb: str) -> str | None:
    day = time.strftime("%Y-%m-%d", time.gmtime())
    try:
        deployed = deploy_ok(args.repo, args.environment, day)
    except ReadError as exc:
        return f"HOLD: could not read the deployments: {exc}"
    if not deployed:
        return f"HOLD: no successful deploy to {args.environment} today ({day} UTC): {verb} waits for one"
    return None


def apply(args: argparse.Namespace, action: Action) -> int:
    refusal = _refusal(args, action)
    if refusal is not None:
        print(refusal)
        return 2
    try:
        targets = action.targets_of(list_issues(args.repo, args.list_limit), args.fix_label, args.park_label)
    except ReadError as exc:
        print(f"HOLD: could not list the issues: {exc}")
        return 2
    if args.dry_run:
        print(f"DRY-RUN: would {action.verb} {len(targets)} issues: {', '.join(f'#{n}' for n in targets) or 'none'} (no edit made)")
        return 0
    failed = [n for n in targets if not edit_label(args.repo, n, action.add, args.park_label)]
    if failed:
        done = len(targets) - len(failed)
        print(f"FAILED: could not {action.verb} {', '.join(f'#{n}' for n in failed) or 'none'}; {done} done; rerun is idempotent")
        return 1
    print(f"{'PARKED' if action.verb == 'park' else 'UNPARKED'}: {len(targets)} issues")
    return 0


def cmd_park(args: argparse.Namespace) -> int:
    return apply(args, Action("red", True, "park", park_targets))


def positive_int(text: str) -> int:
    value = int(text) if text.lstrip("-").isdigit() else 0
    if value < 1:
        reason = f"must be a positive integer: {text!r}"
        raise argparse.ArgumentTypeError(reason)
    return value


def cmd_unpark(args: argparse.Namespace) -> int:
    return apply(args, Action("green", False, "unpark",
                              lambda issues, _fix, park: unpark_targets(issues, park), needs_deploy=True))


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Aparcar issues en una noche roja (#1627, slice 5).")
    sub = parser.add_subparsers(dest="command", required=True)
    for name, helptext in (("park", "aparca en una noche roja"), ("unpark", "desaparca tras una noche verde y un deploy")):
        cmd = sub.add_parser(name, help=helptext)
        cmd.add_argument("--repo", required=True, help="owner/name")
        cmd.add_argument("--night-run", required=True, type=int, help="id de la ejecucion nocturna")
        cmd.add_argument("--workflow", default=NIGHTLY_WORKFLOW,
                         help=f"fichero del workflow nocturno, sin extension (por defecto {NIGHTLY_WORKFLOW})")
        cmd.add_argument("--fix-label", default=FIX_LABEL, help=f"etiqueta de las issues de arreglo (por defecto {FIX_LABEL})")
        cmd.add_argument("--park-label", default=PARK_LABEL, help=f"etiqueta de aparcado (por defecto {PARK_LABEL})")
        cmd.add_argument("--list-limit", type=positive_int, default=DEFAULT_LIST_LIMIT, help="issues maximas leidas")
        cmd.add_argument("--dry-run", action="store_true", help="solo lee y dice lo que haria")
        if name == "unpark":
            cmd.add_argument("--environment", default=DEPLOY_ENVIRONMENT,
                             help=f"entorno del deploy exigido (por defecto {DEPLOY_ENVIRONMENT})")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    args = parse_args(sys.argv[1:] if argv is None else argv)
    return cmd_park(args) if args.command == "park" else cmd_unpark(args)


if __name__ == "__main__":
    sys.exit(main())
