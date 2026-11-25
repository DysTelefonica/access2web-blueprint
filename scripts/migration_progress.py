#!/usr/bin/env python3
"""Generate the migration progress report from GitHub issues (issue #1794).

Reads the phase definition (``docs/07-migracion/seguimiento/fases.json``) and the
issues labelled for phase 1, and writes ``avance.md`` (full checklist).
Done/pending is never edited by hand: a closed issue counts as done; static
tasks carry their own flag.

Everything is validated before anything is written: an issue from a repository
that no group maps fails the run and leaves the output directory untouched.

Exit codes:
    0  report written
    2  invalid input (unmapped repository, truncated search, gh or JSON error)
"""

from __future__ import annotations

import argparse
import datetime
import json
import subprocess
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
TRACKING_DIR = ROOT / "docs" / "07-migracion" / "seguimiento"
SEARCH_LIMIT = 1000
COMMAND = "python scripts/migration_progress.py"

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


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("--gh", default="gh", help="gh executable (default: gh)")
    parser.add_argument("--from-json", type=Path, help="read the gh search JSON from this file")
    parser.add_argument(
        "--date", default=datetime.date.today().isoformat(), help="cut date YYYY-MM-DD"
    )
    parser.add_argument("--phases", type=Path, default=TRACKING_DIR / "fases.json")
    parser.add_argument("--out-dir", type=Path, default=TRACKING_DIR)
    return parser.parse_args(argv)


def load_issues(args: argparse.Namespace, label: str) -> list[Issue]:
    if args.from_json is not None:
        issues: list[Issue] = read_json(args.from_json)
        return issues
    return fetch_issues(args.gh, label)


def generate(args: argparse.Namespace) -> list[dict[str, Any]]:
    datetime.date.fromisoformat(args.date)
    definition = read_json(args.phases)
    labels = [p["issues"]["label"] for p in definition["phases"] if "issues" in p]
    issues = [issue for label in labels for issue in load_issues(args, label)]
    phases = build_phases(definition, issues)
    outputs = {"avance.md": render_markdown(phases, args.date)}
    for name, content in outputs.items():
        (args.out_dir / name).write_text(content, encoding="utf-8", newline="\n")
    return phases


def main(argv: list[str]) -> int:
    args = parse_args(argv)
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
