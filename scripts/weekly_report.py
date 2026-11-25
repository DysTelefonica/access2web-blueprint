#!/usr/bin/env python3
"""Report the status of each migration milestone (issue #1802).

Reads the milestones (``hitos``) and phases of
``docs/07-migracion/seguimiento/fases.json`` and the issues labelled for phase 1,
writes ``avance.md`` and ``avance.html`` as ``migration_progress.py`` does, writes
the snapshot of the cut to ``semanal/AAAA-MM-DD.json`` (one line per record, no
issue bodies), and prints one line per milestone: done/total, estimated week and
status.

A milestone counts its issues, the static tasks of its ``phases`` and its own
``tasks``. It is declared closed only when all its work is done and the cut
reaches its estimated week (``semana_estimada``, or the week of its last
``replanificaciones`` entry); past its first estimated week with work pending it
is delayed. Every phase-1 group must belong to exactly one milestone; everything
is validated before anything is written.

Exit codes:
    0  report written
    2  invalid input (group in no milestone or in several, gh or JSON error)
"""

from __future__ import annotations

import datetime
import json
import re
import sys
from typing import Any

import migration_progress as mp

WEEK = datetime.timedelta(days=7)
MOTIVO_MAX = 200
NO_MOTIVO = "sin motivo indicado"
PROBLEM_HEADING = "### Problema y contexto"
MONTHS = (
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
)  # fmt: skip

Issue = mp.Issue
Week = datetime.date | None


# --- input -----------------------------------------------------------------------------------


def load_input(args: Any, definition: dict[str, Any]) -> list[Issue]:
    """Phase-1 issues, from ``--from-json`` or from GitHub."""
    if args.from_json is not None:
        issues: list[Issue] = mp.read_json(args.from_json)["issues"]
        return issues
    return [i for label in mp.issue_labels(definition) for i in mp.fetch_issues(args.gh, label)]


def label_names(issue: Issue) -> list[str]:
    return sorted(label["name"] for label in issue.get("labels", []))


def first_sentence(text: str) -> str:
    flat = " ".join(text.split())
    match = re.match(r"(.+?[.!?])(?:\s|$)", flat)
    sentence = match.group(1) if match else flat
    return sentence if len(sentence) <= MOTIVO_MAX else sentence[: MOTIVO_MAX - 1] + "…"


def motivo(issue: Issue) -> str:
    parts = (issue.get("body") or "").split(PROBLEM_HEADING, 1)
    if len(parts) < 2:
        return NO_MOTIVO
    section = re.split(r"^#{1,3} ", parts[1], maxsplit=1, flags=re.MULTILINE)[0]
    return first_sentence(section) or NO_MOTIVO


# --- milestones ------------------------------------------------------------------------------


def phase1_groups(definition: dict[str, Any]) -> list[str]:
    spec = next(phase["issues"] for phase in definition["phases"] if "issues" in phase)
    return [spec["block_a"]["name"]] + [group["name"] for group in spec["groups"]]


def validate_hitos(definition: dict[str, Any]) -> None:
    names = phase1_groups(definition)
    hitos = definition["hitos"]
    problems = [
        f"{name} (in {sum(name in h['groups'] for h in hitos)} hitos)"
        for name in names
        if sum(name in h["groups"] for h in hitos) != 1
    ]
    problems += [f"{g} (unknown group)" for h in hitos for g in h["groups"] if g not in names]
    if problems:
        raise mp.ProgressError("each phase-1 group must be in one hito: " + ", ".join(problems))


def planned_weeks(hito: dict[str, Any]) -> list[Week]:
    """The estimated week, then the new week of each replanning, in order."""
    replans = hito.get("replanificaciones", [])
    raws = [hito.get("semana_estimada"), *(r["semana_estimada"] for r in replans)]
    return [datetime.date.fromisoformat(raw) if raw else None for raw in raws]


def hito_status(item: dict[str, Any], cut: datetime.date) -> str:
    complete = 0 < item["total"] == item["done"]
    if complete and item["week"] is not None and cut >= item["week"]:
        return "cerrado"
    if not complete and any(week and cut >= week + WEEK for week in item["planned"]):
        return "retrasado"
    if cut < item["start"]:
        return "en preparación"
    return "en plazo"


def build_hito(
    hito: dict[str, Any],
    groups: dict[str, list[Issue]],
    phases: list[dict[str, Any]],
    cut: datetime.date,
) -> dict[str, Any]:
    items = [issue for name in hito["groups"] for issue in groups[name]]
    tasks = [t for p in phases if p["id"] in hito.get("phases", []) for t in p.get("tasks", [])]
    done, total = mp.counts(items, [*tasks, *hito.get("tasks", [])])
    planned = planned_weeks(hito)
    item = {
        "id": hito["id"],
        "start": datetime.date.fromisoformat(hito["start"]),
        "done": done,
        "total": total,
        "percent": 100 * done / total if total else 0.0,
        "planned": planned,
        "week": planned[-1],
        "phases": hito.get("phases", []),
        "tasks": hito.get("tasks", []),
        "refs": {mp.issue_ref(issue) for issue in items},
    }
    return {**item, "status": hito_status(item, cut)}


def build_hitos(
    definition: dict[str, Any], phases: list[dict[str, Any]], cut: datetime.date
) -> list[dict[str, Any]]:
    groups = {g["name"]: g["issues"] for p in phases for g in p["groups"]}
    hitos = [build_hito(hito, groups, phases, cut) for hito in definition["hitos"]]
    return sorted(hitos, key=lambda hito: hito["start"])


# --- snapshot -------------------------------------------------------------------------------


def issue_entry(issue: Issue, hito: str | None) -> dict[str, Any]:
    return {
        "key": mp.issue_ref(issue),
        "state": str(issue["state"]).lower(),
        "hito": hito,
        "title": issue["title"],
        "url": issue["url"],
        "createdAt": issue.get("createdAt", ""),
        # gh search reports the zero time as the closedAt of an open issue.
        "closedAt": issue.get("closedAt") if mp.is_done(issue) else None,
        "labels": label_names(issue),
        "motivo": motivo(issue),
    }


def make_snapshot(
    date: str, phases: list[dict[str, Any]], hitos: list[dict[str, Any]], issues: list[Issue]
) -> dict[str, Any]:
    hito_of = {ref: hito["id"] for hito in hitos for ref in hito["refs"]}
    task_hito = {phase: hito["id"] for hito in hitos for phase in hito["phases"]}
    tasks = [
        {"phase": p["id"], "hito": task_hito.get(p["id"]), "title": t["title"], "done": t["done"]}
        for p in phases
        for t in p.get("tasks", [])
    ]
    tasks += [
        {"phase": None, "hito": h["id"], "title": t["title"], "done": t["done"]}
        for h in hitos
        for t in h["tasks"]
    ]
    entries = [issue_entry(i, hito_of.get(mp.issue_ref(i))) for i in mp.sort_issues(issues)]
    return {"date": date, "issues": entries, "tasks": tasks}


def compact_json(snapshot: dict[str, Any]) -> str:
    """One record per line, so a weekly snapshot stays a few hundred lines."""
    parts = []
    for key, value in snapshot.items():
        if isinstance(value, list):
            rows = ",\n".join("    " + json.dumps(row, ensure_ascii=False) for row in value)
            parts.append(f"  {json.dumps(key)}: [\n{rows}\n  ]" if rows else f'  "{key}": []')
        else:
            parts.append(f"  {json.dumps(key)}: {json.dumps(value, ensure_ascii=False)}")
    return "{\n" + ",\n".join(parts) + "\n}\n"


# --- text ------------------------------------------------------------------------------------


def long_date(date: datetime.date) -> str:
    return f"{date.day} de {MONTHS[date.month - 1]} de {date.year}"


def week_of(week: Week) -> str:
    return f"del {long_date(week)}" if week else "por definir"


def week_text(week: Week) -> str:
    return f"semana {week_of(week)}"


# --- report ----------------------------------------------------------------------------------


def generate(args: Any) -> list[dict[str, Any]]:
    cut = datetime.date.fromisoformat(args.date)
    definition = mp.read_json(args.phases)
    validate_hitos(definition)
    issues = load_input(args, definition)
    phases = mp.build_phases(definition, issues)
    hitos = build_hitos(definition, phases, cut)
    snapshot = make_snapshot(args.date, phases, hitos, issues)
    mp.write_report(phases, args.date, args.out_dir)
    weekly_dir = args.out_dir / "semanal"
    weekly_dir.mkdir(exist_ok=True)
    (weekly_dir / f"{args.date}.json").write_text(
        compact_json(snapshot), encoding="utf-8", newline="\n"
    )
    return hitos


def main(argv: list[str]) -> int:
    args = mp.build_parser(__doc__).parse_args(argv)
    try:
        hitos = generate(args)
    except (mp.ProgressError, ValueError, KeyError) as exc:
        print(f"weekly_report: {exc}", file=sys.stderr)
        return 2
    for h in hitos:
        print(
            f"Hito {h['id']}: {h['done']}/{h['total']} ({round(h['percent'])} %), "
            f"{week_text(h['week'])}, {h['status']}"
        )
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
