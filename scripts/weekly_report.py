#!/usr/bin/env python3
"""Report the status of each migration milestone (issue #1802).

Reads the milestones (``hitos``) and phases of
``docs/07-migracion/seguimiento/fases.json`` and the issues labelled for phase 1,
and writes:

- ``semanal/AAAA-MM-DD.json``: the snapshot of the cut, one line per record, no
  issue bodies;
- ``semanal/AAAA-MM-DD.html``: one self-contained page for management;
- ``avance.md`` and ``avance.html``, the internal view of ``migration_progress.py``.

It prints one line per milestone: done/total, estimated week and status.

A milestone counts its issues, the static tasks of its ``phases`` and its own
``tasks``. It is declared closed only when all its work is done and the cut
reaches its estimated week (``semana_estimada``, or the week of its last
``replanificaciones`` entry); past its first estimated week with work pending it
is delayed. The management view is a funnel: the current milestone is the first
one not declared closed, and future milestones show no figures. Every phase-1
group must belong to exactly one milestone; everything is validated before
anything is written.

Exit codes:
    0  report written
    2  invalid input (group in no milestone or in several, gh or JSON error)
"""

from __future__ import annotations

import datetime
import html
import json
import re
import sys
from typing import Any

import migration_progress as mp

WEEK = datetime.timedelta(days=7)
MOTIVO_MAX = 200
NO_MOTIVO = "sin motivo indicado"
PROBLEM_HEADING = "### Problema y contexto"
NO_DELAY_MOTIVO = "motivo pendiente de indicar"
FIRST_CUT = "Primer corte, sin comparación."
STYLE_SOURCES = (
    mp.ROOT / "docs" / "uat" / "assets" / "style-base.css",
    mp.TRACKING_DIR / "avance.css",
    mp.TRACKING_DIR / "semanal.css",
)
MONTHS = (
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
)  # fmt: skip
STATUS_PILL = {
    "cerrado": "rag-green",
    "en plazo": "rag-green",
    "en preparación": "rag-grey",
    "retrasado": "rag-red",
}

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
    start = datetime.date.fromisoformat(hito["start"])
    planned = planned_weeks(hito)
    item = {
        "id": hito["id"],
        "start": start,
        "done": done,
        "total": total,
        "percent": 100 * done / total if total else 0.0,
        "planned": planned,
        "week": planned[-1],
        "phases": hito.get("phases", []),
        "tasks": hito.get("tasks", []),
        "refs": {mp.issue_ref(issue) for issue in items},
        "replans": hito.get("replanificaciones", []),
        "label": f"{MONTHS[start.month - 1].capitalize()} {start.year}: {hito['name']}",
    }
    return {**item, "status": hito_status(item, cut)}


def build_hitos(
    definition: dict[str, Any], phases: list[dict[str, Any]], cut: datetime.date
) -> list[dict[str, Any]]:
    groups = {g["name"]: g["issues"] for p in phases for g in p["groups"]}
    hitos = [build_hito(hito, groups, phases, cut) for hito in definition["hitos"]]
    return sorted(hitos, key=lambda hito: hito["start"])


def current_hito(hitos: list[dict[str, Any]]) -> dict[str, Any] | None:
    return next((hito for hito in hitos if hito["status"] != "cerrado"), None)


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


# --- HTML ------------------------------------------------------------------------------------


def esc(text: str) -> str:
    return html.escape(text)


def pill(text: str, css: str) -> str:
    return f'<span class="rag-pill {css}">{esc(text)}</span>'


def card(title: str, body: list[str]) -> list[str]:
    return [f'<section class="card"><h2>{esc(title)}</h2>', *body, "</section>"]


def item_list(items: list[str] | None) -> list[str]:
    """Items are HTML already escaped; ``None`` means there is no previous snapshot."""
    if items is None:
        return [f'<p class="empty">{esc(FIRST_CUT)}</p>']
    return ['<ul class="items">', *(f"<li>{item}</li>" for item in items), "</ul>"]


def long_date(date: datetime.date) -> str:
    return f"{date.day} de {MONTHS[date.month - 1]} de {date.year}"


def week_of(week: Week) -> str:
    return f"del {long_date(week)}" if week else "por definir"


def week_text(week: Week) -> str:
    return f"semana {week_of(week)}"


def replan_lines(hito: dict[str, Any]) -> list[str]:
    weeks = hito["planned"]
    return [
        esc(f"Replanificado: de la semana {week_of(weeks[n])} a la {week_of(weeks[n + 1])}")
        + f" — {esc(replan['motivo'])}"
        for n, replan in enumerate(hito["replans"])
    ]


def delay_lines(hito: dict[str, Any]) -> list[str]:
    if hito["status"] != "retrasado":
        return []
    replans = hito["replans"]
    reason = replans[-1]["motivo"] if replans else NO_DELAY_MOTIVO
    week = hito["week"] if replans else None
    return [
        f'<p class="delay">Motivo del retraso: {esc(reason)}<br />'
        f"Nueva semana prevista: {esc(week_text(week))}</p>"
    ]


def current_section(hito: dict[str, Any] | None) -> list[str]:
    if hito is None:
        return card("Hito actual", ['<p class="empty">Todos los hitos están cerrados.</p>'])
    head = (
        "<tr><th>Hito</th><th>Estado</th><th>Semana estimada</th>"
        "<th>Avance</th><th>Hechas</th><th>%</th></tr>"
    )
    extra = (
        f"<td>{pill(hito['status'], STATUS_PILL[hito['status']])}</td>"
        f"<td>{esc(week_text(hito['week']))}</td>"
    )
    replans = replan_lines(hito)
    return card(
        "Hito actual",
        [
            '<table class="data-table progress-table">',
            f"<thead>{head}</thead><tbody>",
            mp.progress_row(hito["label"], hito["done"], hito["total"], extra),
            "</tbody></table>",
            *delay_lines(hito),
            *(item_list(replans) if replans else []),
        ],
    )


def summary_row(hito: dict[str, Any], visible: set[str]) -> str:
    """A future milestone shows only its name, month and estimated week."""
    status = pill(hito["status"], STATUS_PILL[hito["status"]]) if hito["id"] in visible else ""
    return (
        f'<tr><th scope="row">{esc(hito["label"])}</th><td>{status}</td>'
        f"<td>{esc(week_text(hito['week']))}</td></tr>"
    )


def summary_table(hitos: list[dict[str, Any]], visible: set[str]) -> list[str]:
    return [
        '<table class="data-table progress-table">',
        "<thead><tr><th>Hito</th><th>Estado</th><th>Semana estimada</th></tr></thead><tbody>",
        *(summary_row(hito, visible) for hito in hitos),
        "</tbody></table>",
    ]


def render_weekly(hitos: list[dict[str, Any]], cut: datetime.date) -> str:
    css = "\n".join(path.read_text(encoding="utf-8") for path in STYLE_SOURCES)
    current = current_hito(hitos)
    visible = {hito["id"] for hito in hitos if hito["status"] == "cerrado"}
    visible |= {current["id"]} if current else set()
    lines = [
        "<!doctype html>",
        '<html lang="es">',
        "<head>",
        '<meta charset="utf-8" />',
        '<meta name="viewport" content="width=device-width, initial-scale=1" />',
        f"<style>\n{css}</style>",
        "<title>Access2Web: seguimiento semanal de la migración</title>",
        "</head>",
        "<body>",
        '<div class="page">',
        '<header class="hero" role="banner">',
        f'<div class="eyebrow">Fecha de corte: {esc(long_date(cut))}</div>',
        "<h1>Seguimiento semanal de la migración</h1>",
        '<p class="lead">Avance del hito en curso a partir de las tareas del proyecto: '
        "una tarea cerrada cuenta como hecha.</p>",
        "</header>",
        *current_section(current),
        *card("Todos los hitos", summary_table(hitos, visible)),
        *card("Hecho esta semana", item_list(None)),
        *card("Nuevas esta semana", item_list(None)),
        '<p class="generated">Informe generado a partir de las tareas; no se edita a mano.</p>',
        "</div>",
        "</body>",
        "</html>",
    ]
    return "\n".join(lines) + "\n"


# --- report ----------------------------------------------------------------------------------


def generate(args: Any) -> list[dict[str, Any]]:
    cut = datetime.date.fromisoformat(args.date)
    definition = mp.read_json(args.phases)
    validate_hitos(definition)
    issues = load_input(args, definition)
    phases = mp.build_phases(definition, issues)
    hitos = build_hitos(definition, phases, cut)
    snapshot = make_snapshot(args.date, phases, hitos, issues)
    page = render_weekly(hitos, cut)
    mp.write_report(phases, args.date, args.out_dir)
    weekly_dir = args.out_dir / "semanal"
    weekly_dir.mkdir(exist_ok=True)
    (weekly_dir / f"{args.date}.json").write_text(
        compact_json(snapshot), encoding="utf-8", newline="\n"
    )
    (weekly_dir / f"{args.date}.html").write_text(page, encoding="utf-8", newline="\n")
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
