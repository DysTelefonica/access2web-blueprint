#!/usr/bin/env python3
"""main_red.py: productor del rojo de main (DysTelefonica/team-skills#1708, parte A).

Run en rojo: reutiliza o abre UNA issue abierta con la etiqueta. Run en verde: la cierra.
SKIP si el run no es completado o no es de la rama por defecto. Exit 2 ante cualquier
lectura o escritura fallida (idempotente, se relanza). Sin `gh`: API por `github_rest.py`.
"""

import argparse
import json
import re
import sys
from urllib.parse import quote

from github_rest import GitHubApiError, call, get_json

DEFAULT_LABEL = "prerelease:main-red"
RED_CONCLUSIONS = ("failure", "timed_out")
GREEN_CONCLUSION = "success"
ISSUES_PER_PAGE = 100
REPO_SHAPE = re.compile(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+")


class MainRedError(Exception):
    """Una lectura o escritura falló; el productor falla cerrado."""


def _rerun(repo, run_id):
    return f"relance: main_red.py sync --repo {repo} --run {run_id} (es idempotente)"


def _read(path):
    try:
        return get_json(path)
    except GitHubApiError as exc:
        reason = f"lectura {path}: {exc}"
        raise MainRedError(reason) from exc


def _write(method, path, body):
    try:
        text, _ = call(method, path, body)
    except GitHubApiError as exc:
        reason = f"escritura {method} {path}: {exc}"
        raise MainRedError(reason) from exc
    return text


def _open_labelled(repo, label):
    """Números de las issues ABIERTAS con la etiqueta, sin pull requests, ascendentes."""
    items = _read(f"repos/{repo}/issues?labels={quote(label, safe='')}&state=open&per_page={ISSUES_PER_PAGE}")
    if not isinstance(items, list):
        reason = f"el listado de issues con {label} no es una lista"
        raise MainRedError(reason)
    numbers = [item["number"] for item in items
               if isinstance(item, dict) and "pull_request" not in item and isinstance(item.get("number"), int)]
    return sorted(numbers)


def _merged_pr(repo, sha):
    """Número de la primera PR asociada al commit, o None."""
    pulls = _read(f"repos/{repo}/commits/{sha}/pulls")
    if not isinstance(pulls, list):
        reason = f"la lista de PR de {sha[:7]} no es una lista"
        raise MainRedError(reason)
    first = next((pull for pull in pulls if isinstance(pull, dict) and isinstance(pull.get("number"), int)), None)
    return first["number"] if first else None


def _run_url_comment(run):
    return f"Run en rojo: {run.get('html_url') or 'sin URL'} (SHA {str(run.get('head_sha'))[:7]})."


def _open_red(repo, run, label, dry_run):
    issues = _open_labelled(repo, label)
    if issues:
        number = issues[0]
        if dry_run:
            return f"DRY-RUN: would comment on issue #{number} with the red run"
        _write("POST", f"repos/{repo}/issues/{number}/comments", {"body": _run_url_comment(run)})
        return f"RED: reused issue #{number}, commented the red run"
    sha = str(run.get("head_sha"))
    pr = _merged_pr(repo, sha)
    title = f"main en rojo: {run.get('name') or 'workflow'} {sha[:7]}"
    pr_text = f"#{pr}" if pr is not None else "desconocido"
    body = (f"El run {run.get('html_url') or 'sin URL'} de main terminó en rojo.\n\n"
            f"- SHA: {sha}\n- PR fusionada: {pr_text}\n\n"
            "Se cierra sola con el siguiente run verde de main.")
    if dry_run:
        return f"DRY-RUN: would open an issue «{title}» labelled {label}"
    text = _write("POST", f"repos/{repo}/issues", {"title": title, "body": body, "labels": [label]})
    created = json.loads(text).get("number") if text else None
    if not isinstance(created, int):
        reason = "la API no devolvió el número de la issue creada"
        raise MainRedError(reason)
    return f"RED: opened issue #{created} «{title}»"


def _close_green(repo, run, label, dry_run):
    issues = _open_labelled(repo, label)
    if not issues:
        return "GREEN: nothing open with " + label
    if dry_run:
        return "DRY-RUN: would close " + ", ".join(f"#{number}" for number in issues)
    for number in issues:
        _write("POST", f"repos/{repo}/issues/{number}/comments", {"body": _run_url_comment(run)})
        _write("PATCH", f"repos/{repo}/issues/{number}", {"state": "closed", "state_reason": "completed"})
    return "GREEN: closed " + ", ".join(f"#{number}" for number in issues)


def sync(repo, run_id, label, dry_run):
    if not REPO_SHAPE.fullmatch(repo):
        reason = f"--repo debe ser owner/name, no {repo!r}"
        raise MainRedError(reason)
    run = _read(f"repos/{repo}/actions/runs/{run_id}")
    if not isinstance(run, dict):
        reason = f"el run {run_id} no es un objeto"
        raise MainRedError(reason)
    repo_doc = _read(f"repos/{repo}")
    default = repo_doc.get("default_branch") if isinstance(repo_doc, dict) else None
    if not isinstance(default, str) or not default:
        reason = f"el repo {repo} no informa su rama por defecto"
        raise MainRedError(reason)
    if run.get("status") != "completed" or run.get("head_branch") != default:
        return f"SKIP: run {run_id} is not a completed run on the default branch"
    conclusion = run.get("conclusion")
    if conclusion == GREEN_CONCLUSION:
        return _close_green(repo, run, label, dry_run)
    if conclusion in RED_CONCLUSIONS:
        return _open_red(repo, run, label, dry_run)
    return f"SKIP: run {run_id} ended «{conclusion}», not red or green"


def _pin_output_encoding():
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")


def main(argv=None):
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description="Productor del rojo de main (#1708, parte A).")
    sub = parser.add_subparsers(dest="command", required=True)
    cmd = sub.add_parser("sync", help="abre, reutiliza o cierra la issue de main en rojo")
    cmd.add_argument("--repo", required=True, help="owner/name")
    cmd.add_argument("--run", required=True, type=int, help="id del run del CI en la rama por defecto")
    cmd.add_argument("--label", default=DEFAULT_LABEL, help="etiqueta de la issue de main en rojo")
    cmd.add_argument("--dry-run", action="store_true", help="dice lo que haría; no escribe")
    args = parser.parse_args(argv)
    try:
        print(sync(args.repo, args.run, args.label, args.dry_run))
    except MainRedError as exc:
        print(f"FAILED: {exc}. {_rerun(args.repo, args.run)}")
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
