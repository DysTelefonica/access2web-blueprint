---
name: a2w-seguimiento-semanal
description: 'Trigger: informe semanal, reunión semanal con la gerencia, avance del hito, seguimiento semanal, retrasos, bloqueos, riesgos. Prepara el informe semanal de avance por hito de la migración a partir de las issues, con su foto semanal, y lo entrega por PR. Skill propia de este repositorio; no se propaga a team-skills.'
license: Apache-2.0
metadata:
  author: ardelperal
  version: "1.1"
  last_verified: 2026-11-25
---

# access2web weekly progress report

## Activation Contract

Use this skill to prepare the weekly meeting with management on the migration (epic #1545): the milestone in progress, delays with their reason and new week, what was done and what appeared this week and why, blockers, risks and the decisions management has to take (#1802). The report is generated from GitHub issues by `scripts/weekly_report.py`; only blockers, risks and decisions are marked by hand, as labels, and the estimated week of each milestone is set by the owners in `docs/07-migracion/seguimiento/fases.json`. It is owned by this repository and is not propagated from team-skills nor installed in a user's global skills.

There are two views. The **management view** (`semanal/AAAA-MM-DD.html`) is a funnel: management receives the information as agreed, milestone by milestone, even when later work is already done. The **internal view** (`avance.md`, `avance.html`) shows every phase, group and issue as it is.

## Hard Rules

- **HR-1** Never hand-edit a generated file: `semanal/AAAA-MM-DD.json`, `semanal/AAAA-MM-DD.html`, `avance.md` and `avance.html` are rewritten by the script. Change the issues, `fases.json` or the script instead.
- **HR-2** Never report as done something that is not done: an issue counts as done only when it is closed, a static task only when its `done` is `true`.
- **HR-3** Never hide a delay: a milestone past its estimated week with work pending is shown as «retrasado» even after a replanning, with the reason and the new week.
- **HR-4** Set and replan the estimated week in `fases.json` before generating, never in the report: `semana_estimada` is the Monday (`AAAA-MM-DD`) of the estimated finish week, `null` until agreed («semana por definir»); a replanning appends `{"fecha", "semana_estimada", "motivo"}` to `replanificaciones` and keeps the original week.
- **HR-5** Never put technical detail of the cleanup in the management report (tables, fields, VBA, SQL): it shows counts, percentages, statuses and titles only. A title that only makes sense to the team is fixed in the issue, not in the HTML.
- **HR-6** Mark the issue, not the report: a blocker is an open issue labelled `bloqueado`, a risk one labelled `riesgo` (with `riesgo:alto`, `riesgo:medio` or `riesgo:bajo`), a decision one labelled `decision:gerencia`. Blockers may live in any repository of the cleanup; risks and decisions live in `DysTelefonica/access2web-blueprint`.
- **HR-7** Write the reason where the script reads it: the first sentence of `### Problema y contexto` in the issue body, and for a blocker its latest comment. An issue without that section shows «sin motivo indicado».
- **HR-8** Run every `gh` call with the profile of the admin who runs the session (`GH_TOKEN= GITHUB_TOKEN=`), and deliver the snapshot and the HTML through a PR per `skills/a2w-issue-flow/SKILL.md`; never push them to `main`.

## Decision Gates

| Condition | Action |
| --- | --- |
| A milestone has all its work done and the cut reaches its estimated week | It is declared closed («cerrado») and the next milestone becomes the current one |
| The current milestone has all its work done before its estimated week | It stays «en plazo»; it closes in its week, not earlier |
| The current milestone has not reached its start month | It shows «en preparación» with its real progress, work done earlier included |
| The current milestone passed its estimated week with work pending | Before generating, append a replanning with its reason and new week; without one the report says «motivo pendiente de indicar» and «semana por definir» |
| A future milestone has work done | Expected: the management view shows only its name, month and estimated week, and its issues stay out of «Hecho» and «Nuevas» |
| Blockers, risks or decisions belong to a future milestone | Expected: they are always shown |
| The script exits 2 naming a group or a repository | Fix `hitos` or the groups in `fases.json`, then run again |
| A fix is applied but its issue is still open | Close the issue before running: an open issue counts as pending |
| There is no earlier snapshot in `semanal/` | Expected: the report says «Primer corte, sin comparación.» |

## Execution Steps

1. **Day before the meeting**: review the open issues of the cleanup, close the ones done, and label `bloqueado`, `riesgo` (with its level) and `decision:gerencia` as needed. Agree with the other owner the estimated week of the current milestone and any replanning, and write it in `fases.json`.
2. **Generate** from a worktree of `main`, with the meeting's cut date: `GH_TOKEN= GITHUB_TOKEN= python scripts/weekly_report.py --gh "/c/Program Files/GitHub CLI/gh.exe" --date <AAAA-MM-DD>` (on Linux, `--gh gh`). The console shows one internal line per milestone: done/total, estimated week and status.
3. **Review** `docs/07-migracion/seguimiento/semanal/<AAAA-MM-DD>.html` in a browser: current milestone, all milestones, done and new this week, blockers, risks, decisions.
4. **Deliver**: commit `fases.json` if it changed, the new `semanal/<AAAA-MM-DD>.json` and `.html` and the refreshed `avance.md` and `avance.html` in one PR (`docs(platform): informe semanal <AAAA-MM-DD>`), following `skills/a2w-issue-flow/SKILL.md`. The JSON is the comparison base of next week.

## Output Contract

| Key | Type | Description |
| --- | --- | --- |
| `cut_date` | string | Cut date of the report, `AAAA-MM-DD`. |
| `report` | string | Path of the generated HTML. |
| `milestones` | string[] | The console lines of the run, one per milestone. |
| `replanned` | string[] | Milestones replanned in `fases.json` for this cut, or empty. |
| `pr` | string | PR that delivers the snapshot and the HTML, or `none`. |

## References

- `scripts/weekly_report.py`, `scripts/migration_progress.py` — the generators; `scripts/README.md` lists them.
- `docs/07-migracion/seguimiento/fases.json` — phases and milestones (`hitos`) with their groups, tasks, estimated week and replannings.
- `docs/07-migracion/presentacion/cronograma-hitos.html` — the milestones as presented to management.
