#!/usr/bin/env python3
"""build_night_snapshot.py: instantánea JSON para night_report.py (#1627, slice 4b).

Lee de GitHub con `gh api` (solo GET). Incluye los PRs fusionados desde la última noche
buena con sus issues, los jobs de la noche y el deploy a producción del día UTC (un
GitHub Deployment en `production` con algún estado `success`). Sin `--out` ni `--dry-run`
sale con 2; cualquier lectura fallida sale con 2 sin escribir (fallo cerrado).
"""

import argparse
import json
import os
import re
import subprocess
import sys
from pathlib import Path

PAGE = 100  # tamaño de página de la API; una página llena sin acotar falla cerrado
DEFAULT_EXPENSIVE = r"(?i)e2e|integra|upgrade|trivy|security|seguridad"
GRAPHQL = (
    "query($owner:String!,$name:String!,$pr:Int!){repository(owner:$owner,name:$name)"
    "{pullRequest(number:$pr){closingIssuesReferences(first:50){nodes{number}}}}}"
)


class SnapshotError(Exception):
    """Lectura o acotado imposible: la instantánea no se escribe."""


def gh_command():
    """`gh` a ejecutar: `$CI_PATTERN_GH` (un `.py` corre con este Python) o el del PATH."""
    explicit = os.environ.get("CI_PATTERN_GH")
    if explicit and explicit.endswith(".py"):
        return [sys.executable, explicit]
    return [explicit or "gh"]


def gh_api(*args):
    """Ejecuta `gh api ...` y devuelve el JSON; cualquier fallo es SnapshotError."""
    proc = subprocess.run(  # noqa: S603 - argv fijo, sin shell
        [*gh_command(), "api", *args], capture_output=True, text=True,
        encoding="utf-8", check=False,
    )
    if proc.returncode != 0:
        message = f"no se puede leer {args[0]}: {proc.stderr.strip() or 'gh falló'}"
        raise SnapshotError(message)
    try:
        return json.loads(proc.stdout)
    except json.JSONDecodeError as exc:
        message = f"respuesta no válida de {args[0]}: {exc}"
        raise SnapshotError(message) from exc


def utc_day(stamp):
    return (stamp or "")[:10]


def night_run(repo, workflow, branch):
    runs = gh_api(f"repos/{repo}/actions/workflows/{workflow}/runs"
                  f"?branch={branch}&event=schedule&per_page=1").get("workflow_runs") or []
    if not runs:
        raise SnapshotError("no hay ejecución nocturna de " + workflow)
    run = runs[0]
    return run, {"run_id": run["id"], "conclusion": run.get("conclusion"),
                 "url": run.get("html_url"), "date": utc_day(run.get("created_at"))}


def last_good_night(repo, workflow, branch):
    runs = gh_api(f"repos/{repo}/actions/workflows/{workflow}/runs"
                  f"?branch={branch}&event=schedule&status=success&per_page=1").get("workflow_runs") or []
    if not runs:
        return None
    return {"run_id": runs[0]["id"], "finished_at": runs[0]["updated_at"]}


def merged_prs(repo, branch, cutoff):
    """PRs fusionados tras `cutoff`. Una página llena que no llega al corte falla cerrado."""
    items = gh_api(f"repos/{repo}/pulls?state=closed&base={branch}&sort=updated"
                   f"&direction=desc&per_page={PAGE}")
    if len(items) >= PAGE and (cutoff is None or items[-1]["updated_at"] >= cutoff):
        message = f"más de {PAGE} PRs desde la última noche buena: no se puede acotar"
        raise SnapshotError(message)
    selected = [p for p in items if p.get("merged_at") and (cutoff is None or p["merged_at"] > cutoff)]
    return [{"number": p["number"], "title": p["title"], "merged_at": p["merged_at"],
             "merge_sha": p["merge_commit_sha"], "issues": linked_issues(repo, p["number"])}
            for p in selected]


def linked_issues(repo, number):
    owner, name = repo.split("/")
    data = gh_api("graphql", "-f", f"query={GRAPHQL}", "-F", f"owner={owner}",
                  "-F", f"name={name}", "-F", f"pr={number}")
    nodes = data["data"]["repository"]["pullRequest"]["closingIssuesReferences"]["nodes"]
    return [node["number"] for node in nodes]


def night_jobs(repo, run_id, expensive):
    jobs = gh_api(f"repos/{repo}/actions/runs/{run_id}/jobs?per_page={PAGE}")["jobs"]
    return [{"name": j["name"], "expensive": bool(re.search(expensive, j["name"])),
             "conclusion": j.get("conclusion"), "run_url": j.get("html_url")} for j in jobs]


def production_deploy(repo, day, environment):
    """Deploy con éxito: un GitHub Deployment creado ese día UTC con algún estado `success`."""
    deployments = gh_api(f"repos/{repo}/deployments?environment={environment}&per_page=50")
    for deployment in deployments:
        if utc_day(deployment.get("created_at")) != day:
            continue
        statuses = gh_api(f"repos/{repo}/deployments/{deployment['id']}/statuses?per_page=1")
        for status in statuses:
            if status.get("state") == "success":
                return {"deployed": True, "at": status.get("created_at"),
                        "environment": environment, "deployment_id": deployment["id"]}
    return {"deployed": False, "at": None, "environment": environment, "deployment_id": None}


def build(args):
    run, night = night_run(args.repo, args.workflow, args.branch)
    good = last_good_night(args.repo, args.workflow, args.branch)
    cutoff = good["finished_at"] if good else None
    day = args.date or night["date"]
    return {
        "repository": args.repo,
        "night": night,
        "last_good_night": good,
        "prs": merged_prs(args.repo, args.branch, cutoff),
        "jobs": night_jobs(args.repo, run["id"], args.expensive),
        "deploy": production_deploy(args.repo, day, args.environment),
    }


def parse(argv):
    parser = argparse.ArgumentParser(description="instantánea JSON de la noche (#1627)")
    parser.add_argument("--repo", required=True, help="OWNER/NAME")
    parser.add_argument("--workflow", default="tests.yml", help="workflow nocturno")
    parser.add_argument("--branch", default="main", help="rama de la noche")
    parser.add_argument("--environment", default="production", help="entorno de deploy")
    parser.add_argument("--expensive", default=DEFAULT_EXPENSIVE,
                        help="regex de nombres de job caros")
    parser.add_argument("--date", help="día UTC del deploy (por defecto, el de la noche)")
    parser.add_argument("--out", type=Path, help="fichero de salida")
    parser.add_argument("--dry-run", action="store_true", help="imprime, no escribe")
    args = parser.parse_args(argv)
    if args.out is None and not args.dry_run:
        parser.error("indique --out o --dry-run")
    return args


def _pin_output_encoding():
    """Fija stdout y stderr a UTF-8 para que los mensajes con tildes salgan íntegros."""
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")


def main(argv=None):
    _pin_output_encoding()
    args = parse(sys.argv[1:] if argv is None else argv)
    try:
        snapshot = build(args)
    except SnapshotError as exc:
        print(f"fail-closed: {exc}", file=sys.stderr)
        return 2
    text = json.dumps(snapshot, ensure_ascii=False, indent=2, sort_keys=True) + "\n"
    if args.dry_run:
        sys.stdout.write(text)
    else:
        args.out.write_text(text, encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
