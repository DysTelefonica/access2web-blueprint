#!/usr/bin/env python3
"""Generate the migration progress report from GitHub issues (issue #1794).

Reads the phase definition (``docs/07-migracion/seguimiento/fases.json``) and the
issues labelled for phase 1, and writes ``avance.md`` (full checklist) and
``avance.html`` (aggregates for management). Done/pending is never edited by
hand: a closed issue counts as done; static tasks carry their own flag.

Everything is validated before anything is written: an issue from a repository
that no group maps fails the run and leaves the output directory untouched.

Exit codes:
    0  report written
    2  invalid input (unmapped repository, truncated search, gh or JSON error)
"""

from __future__ import annotations

import argparse
import datetime
import html
import json
import subprocess
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
TRACKING_DIR = ROOT / "docs" / "07-migracion" / "seguimiento"
SEARCH_LIMIT = 1000
COMMAND = "python scripts/migration_progress.py"
STATUS_PILL = {"revisada": "rag-green", "hecha": "rag-green", "en curso": "rag-amber"}

Issue = dict[str, Any]
Group = dict[str, Any]


class ProgressError(Exception):
    """Input that makes the report impossible to generate faithfully."""


def fetch_issues(gh: str, label: str) -> list[Issue]:
    # Without --state, gh search returns open and closed issues alike.
    command = [
        gh,
        "search",
        "issues",
        "--owner",
        "DysTelefonica",
        "--label",
        label,
        "--limit",
        str(SEARCH_LIMIT),
        "--json",
        "repository,number,state,title,url",
    ]
    try:
        result = subprocess.run(
            command, capture_output=True, text=True, encoding="utf-8", check=False
        )
    except OSError as exc:
        raise ProgressError(f"cannot run gh: {exc}") from exc
    if result.returncode != 0:
        raise ProgressError(f"gh search failed: {result.stderr.strip()}")
    issues: list[Issue] = json.loads(result.stdout)
    if len(issues) >= SEARCH_LIMIT:
        raise ProgressError(f"gh search returned {len(issues)} issues: result truncated")
    return issues


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise ProgressError(f"cannot read {path}: {exc}") from exc


def is_done(issue: Issue) -> bool:
    return str(issue["state"]).lower() == "closed"


def repo_index(groups: list[Group]) -> dict[str, int]:
    return {repo: position for position, group in enumerate(groups) for repo in group["repos"]}


def classify(issues: list[Issue], spec: dict[str, Any]) -> list[Group]:
    """Split issues into Bloque A plus one group per application, in spec order."""
    block_a = spec["block_a"]
    index = repo_index(spec["groups"])
    buckets: list[list[Issue]] = [[] for _ in spec["groups"]]
    first: list[Issue] = []
    unmapped: set[str] = set()
    for issue in issues:
        repo = issue["repository"]["name"]
        if block_a["title_contains"] in issue["title"]:
            first.append(issue)
        elif repo in index:
            buckets[index[repo]].append(issue)
        else:
            unmapped.add(repo)
    if unmapped:
        raise ProgressError("repositories not mapped to any group: " + ", ".join(sorted(unmapped)))
    names = [block_a["name"]] + [group["name"] for group in spec["groups"]]
    return [
        {"name": name, "issues": sort_issues(items)}
        for name, items in zip(names, [first, *buckets], strict=True)
    ]


def sort_issues(issues: list[Issue]) -> list[Issue]:
    return sorted(issues, key=lambda issue: (issue["repository"]["name"].lower(), issue["number"]))


def counts(items: list[Issue], tasks: list[dict[str, Any]]) -> tuple[int, int]:
    done = sum(1 for issue in items if is_done(issue)) + sum(1 for task in tasks if task["done"])
    return done, len(items) + len(tasks)


def build_phases(definition: dict[str, Any], issues: list[Issue]) -> list[dict[str, Any]]:
    phases = []
    for phase in definition["phases"]:
        groups = classify(issues, phase["issues"]) if "issues" in phase else []
        every_issue = [issue for group in groups for issue in group["issues"]]
        done, total = counts(every_issue, phase.get("tasks", []))
        phases.append({**phase, "groups": groups, "done": done, "total": total})
    return phases


def status_text(phase: dict[str, Any]) -> str:
    note = phase.get("status_note")
    return f"{phase['status']} ({note})" if note else str(phase["status"])


def issue_ref(issue: Issue) -> str:
    return f"{issue['repository']['name']}#{issue['number']}"


def box(done: bool) -> str:
    return "[x]" if done else "[ ]"


def markdown_phase(phase: dict[str, Any]) -> list[str]:
    lines = [
        f"## Fase {phase['id']}. {phase['name']}",
        "",
        f"**Estado:** {status_text(phase)} · **Hechas:** {phase['done']}/{phase['total']}",
        "",
    ]
    if phase.get("tasks"):
        if phase["groups"]:
            lines += ["### Preparación", ""]
        lines += [f"- {box(task['done'])} {task['title']}" for task in phase["tasks"]] + [""]
    for group in phase["groups"]:
        done, total = counts(group["issues"], [])
        lines += [f"### {group['name']} ({done}/{total})", ""]
        lines += [
            f"- {box(is_done(i))} [{issue_ref(i)}]({i['url']}) {i['title']}"
            for i in group["issues"]
        ]
        lines += [""]
    return lines


def render_markdown(phases: list[dict[str, Any]], date: str) -> str:
    lines = [
        "# Avance de la migración por fases",
        "",
        "> Fichero generado. No edite este fichero a mano: se regenera con "
        f"`{COMMAND}` a partir de `fases.json` y de las issues de GitHub. "
        f"Fecha de corte: {date}.",
        "",
        "| Fase | Nombre | Estado | Hechas | Total |",
        "|---|---|---|---|---|",
    ]
    lines += [
        f"| {p['id']} | {p['name']} | {status_text(p)} | {p['done']} | {p['total']} |"
        for p in phases
    ]
    lines += [""]
    for phase in phases:
        lines += markdown_phase(phase)
    return "\n".join(lines).rstrip("\n") + "\n"


def progress_row(name: str, done: int, total: int, extra: str = "") -> str:
    percent = round(100 * done / total) if total else 0
    return (
        f'<tr><th scope="row">{html.escape(name)}</th>{extra}'
        f'<td><progress value="{done}" max="{total}">{percent} %</progress></td>'
        f'<td class="num">{done}/{total}</td><td class="num">{percent} %</td></tr>'
    )


def html_phase_table(phases: list[dict[str, Any]]) -> list[str]:
    rows = []
    for phase in phases:
        # A status note qualifies the status ("revisada" design, not yet executed): grey.
        pill = (
            "rag-grey" if phase.get("status_note") else STATUS_PILL.get(phase["status"], "rag-grey")
        )
        status = f'<td><span class="rag-pill {pill}">{html.escape(status_text(phase))}</span></td>'
        rows.append(
            progress_row(f"{phase['id']}. {phase['name']}", phase["done"], phase["total"], status)
        )
    head = "<tr><th>Fase</th><th>Estado</th><th>Avance</th><th>Hechas</th><th>%</th></tr>"
    return [
        '<table class="data-table progress-table">',
        f"<thead>{head}</thead><tbody>",
        *rows,
        "</tbody></table>",
    ]


def html_group(group: Group) -> list[str]:
    done, total = counts(group["issues"], [])
    items = [
        f'<li class="{"done" if is_done(i) else "open"}"><a href="{html.escape(i["url"])}">'
        f"{html.escape(issue_ref(i))}</a> {html.escape(i['title'])}</li>"
        for i in group["issues"]
    ]
    return [
        progress_row(group["name"], done, total),
        f'<tr class="detail"><td colspan="4"><details><summary>Issues</summary><ul>'
        f"{''.join(items)}</ul></details></td></tr>",
    ]


def html_groups(phases: list[dict[str, Any]]) -> list[str]:
    lines = []
    for phase in (p for p in phases if p["groups"]):
        lines += [
            f'<section class="card"><h2>Fase {phase["id"]}: avance por aplicación</h2>',
            '<table class="data-table progress-table"><thead><tr><th>Grupo</th>'
            "<th>Avance</th><th>Hechas</th><th>%</th></tr></thead><tbody>",
        ]
        for group in phase["groups"]:
            lines += html_group(group)
        lines += ["</tbody></table></section>"]
    return lines


def render_html(phases: list[dict[str, Any]], date: str) -> str:
    lines = [
        "<!doctype html>",
        '<html lang="es">',
        "<head>",
        '<meta charset="utf-8" />',
        '<meta name="viewport" content="width=device-width, initial-scale=1" />',
        '<link rel="stylesheet" href="../../uat/assets/style-base.css" />',
        '<link rel="stylesheet" href="avance.css" />',
        "<title>Access2Web: avance de la migración por fases</title>",
        "</head>",
        "<body>",
        '<div class="page">',
        '<header class="hero" role="banner">',
        f'<div class="eyebrow">Épica #1545 · Fecha de corte: {html.escape(date)}</div>',
        "<h1>Avance de la migración por fases</h1>",
        '<p class="lead">Generado a partir de las issues de GitHub: una issue cerrada cuenta como '
        'hecha. La lista completa de subtareas está en <a href="avance.md">avance.md</a>.</p>',
        "</header>",
        '<section class="card"><h2>Avance por fase</h2>',
        *html_phase_table(phases),
        "</section>",
        *html_groups(phases),
        f'<p class="generated">Generado con <code>{COMMAND}</code>. No se edita a mano.</p>',
        "</div>",
        "</body>",
        "</html>",
    ]
    return "\n".join(lines) + "\n"


def build_parser(description: str | None) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=description, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("--gh", default="gh", help="gh executable (default: gh)")
    parser.add_argument("--from-json", type=Path, help="read the gh search JSON from this file")
    parser.add_argument(
        "--date", default=datetime.date.today().isoformat(), help="cut date YYYY-MM-DD"
    )
    parser.add_argument("--phases", type=Path, default=TRACKING_DIR / "fases.json")
    parser.add_argument("--out-dir", type=Path, default=TRACKING_DIR)
    return parser


def issue_labels(definition: dict[str, Any]) -> list[str]:
    return [p["issues"]["label"] for p in definition["phases"] if "issues" in p]


def write_report(phases: list[dict[str, Any]], date: str, out_dir: Path) -> None:
    outputs = {
        "avance.md": render_markdown(phases, date),
        "avance.html": render_html(phases, date),
    }
    for name, content in outputs.items():
        (out_dir / name).write_text(content, encoding="utf-8", newline="\n")


def load_issues(args: argparse.Namespace, label: str) -> list[Issue]:
    if args.from_json is not None:
        issues: list[Issue] = read_json(args.from_json)
        return issues
    return fetch_issues(args.gh, label)


def generate(args: argparse.Namespace) -> list[dict[str, Any]]:
    datetime.date.fromisoformat(args.date)
    definition = read_json(args.phases)
    issues = [issue for label in issue_labels(definition) for issue in load_issues(args, label)]
    phases = build_phases(definition, issues)
    write_report(phases, args.date, args.out_dir)
    return phases


def main(argv: list[str]) -> int:
    args = build_parser(__doc__).parse_args(argv)
    try:
        phases = generate(args)
    except (ProgressError, ValueError, KeyError) as exc:
        print(f"migration_progress: {exc}", file=sys.stderr)
        return 2
    for phase in phases:
        print(f"Fase {phase['id']}: {phase['done']}/{phase['total']}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
