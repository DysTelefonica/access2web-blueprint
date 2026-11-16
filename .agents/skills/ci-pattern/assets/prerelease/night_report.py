#!/usr/bin/env python3
"""Informe nocturno en HTML del perfil pre-release (#1627, slice 4).

Lee una instantánea validada por ``snapshot.py`` y escribe una página HTML con: los PRs
desde la última noche buena, agrupados y enlazados a sus issues; el resultado de cada
job caro; y si hubo deploy.

Sin red ni ``gh``. Idempotente: la misma entrada da el mismo fichero. ``--dry-run``
imprime el HTML y no escribe nada. Salida: 0 correcto, 2 instantánea o uso no válidos.
"""

from __future__ import annotations

import argparse
import html
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from snapshot import SnapshotError, load, when  # noqa: E402


def prs_since_last_good(snap: dict) -> list[dict]:
    """PRs fusionados después de la última noche buena, por número ascendente."""
    good = snap["last_good_night"]
    cutoff: datetime | None = when(good["finished_at"], "finished_at", "last_good_night") if good else None
    selected = [pr for pr in snap["prs"]
                if cutoff is None or when(pr["merged_at"], "merged_at", "PR") > cutoff]
    return sorted(selected, key=lambda pr: pr["number"])


def group_by_issue(prs: list[dict]) -> list[tuple[int | None, list[dict]]]:
    """Agrupa por issue vinculada; un PR con varias issues aparece en cada grupo.

    Los PRs sin issue quedan en el grupo ``None``, que va el último.
    """
    groups: dict[int | None, list[dict]] = {}
    for pr in prs:
        for issue in pr["issues"] or [None]:
            groups.setdefault(issue, []).append(pr)
    ordered = sorted(k for k in groups if k is not None)
    if None in groups:
        ordered.append(None)
    return [(k, groups[k]) for k in ordered]


def _esc(value: object) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def _pr_link(repo: str, number: int) -> str:
    return f"https://github.com/{repo}/pull/{number}"


def _issue_link(repo: str, number: int) -> str:
    return f"https://github.com/{repo}/issues/{number}"


def render(snap: dict) -> str:
    repo = snap["repository"]
    prs = prs_since_last_good(snap)
    good = snap["last_good_night"]
    since = _esc(good["finished_at"]) if good else "sin noche buena previa"
    rows = []
    for issue, group in group_by_issue(prs):
        head = (f'<a href="{_issue_link(repo, issue)}">#{issue}</a>' if issue is not None
                else "Sin issue vinculada")
        items = "".join(
            f'<li><a href="{_pr_link(repo, p["number"])}">#{p["number"]}</a> {_esc(p["title"])}</li>'
            for p in group)
        rows.append(f"<section><h2>{head}</h2><ul>{items}</ul></section>")
    jobs = []
    for job in snap["jobs"]:
        if not job["expensive"]:
            continue
        conclusion = job["conclusion"] or "sin resultado"
        name = _esc(job["name"])
        link = f'<a href="{_esc(job["run_url"])}">{name}</a>' if job["run_url"] else name
        jobs.append(f"<tr><td>{link}</td><td>{_esc(conclusion)}</td></tr>")
    deploy = snap["deploy"]
    deploy_text = "desplegado" if deploy["deployed"] else "no desplegado"
    deploy_at = f' ({_esc(deploy["at"])})' if deploy.get("at") else ""
    return (
        "<!doctype html>\n<html lang=\"es\"><head><meta charset=\"utf-8\">"
        "<title>Informe nocturno</title></head><body>\n"
        f"<h1>Informe nocturno {_esc(repo)}</h1>\n"
        f"<p>Noche: {_esc(snap['night']['started_at'])} · desde la última noche buena: {since}</p>\n"
        f"<p>Deploy: {deploy_text}{deploy_at}</p>\n"
        "<h2>Trabajos caros</h2>\n"
        f"<table><tr><th>Job</th><th>Resultado</th></tr>{''.join(jobs)}</table>\n"
        f"<h2>PRs desde la última noche buena ({len(prs)})</h2>\n"
        + "\n".join(rows) + "\n</body></html>\n"
    )


def _pin_output_encoding() -> None:
    """Fija stdout y stderr a UTF-8: un título con emoji no debe tumbar el script (#1627)."""
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--snapshot", required=True, type=Path)
    parser.add_argument("--out", type=Path)
    parser.add_argument("--dry-run", action="store_true",
                        help="imprime el HTML y no escribe ningún fichero")
    args = parser.parse_args(argv)
    if args.out is None and not args.dry_run:
        print("indique --out o --dry-run", file=sys.stderr)
        return 2
    try:
        page = render(load(args.snapshot))
    except SnapshotError as error:
        print(f"instantánea no válida: {error}", file=sys.stderr)
        return 2
    if args.dry_run:
        sys.stdout.write(page)
        return 0
    args.out.write_text(page, encoding="utf-8", newline="\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
