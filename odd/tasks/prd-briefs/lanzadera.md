# Brief — PRD de la aplicación `lanzadera` (issue #1074)

You are a documentation agent working UNATTENDED for the access2web-blueprint coordinator. Do not ask the human questions; if you hit a real blocker, write it in your final report and stop that sub-part.

## Skills to load before work (read each SKILL.md fully first)
- /home/ubuntu/repos/access2web-blueprint/.agents/skills/ai-slop-discipline/SKILL.md
- /home/ubuntu/repos/access2web-blueprint/.agents/skills/governed-delivery/SKILL.md
- /home/ubuntu/.agents/skills/access-vba-capability-docs/SKILL.md
- /home/ubuntu/.agents/skills/codegraph-usage/SKILL.md
- /home/ubuntu/.agents/skills/chained-pr/SKILL.md
- /home/ubuntu/.agents/skills/documentation-alan-style/SKILL.md
- /home/ubuntu/.agents/skills/work-unit-commits/SKILL.md

## Contract (read in this order)
1. Issue: `gh issue view 1074 -R DysTelefonica/access2web-blueprint` — its 6 sections are your acceptance contract.
2. Roadmap and PRD template (PR #1072, maybe not merged yet): read them from
   /home/ubuntu/repos/access2web-blueprint-worktrees/roadmap-capacidades/docs/07-migracion/roadmap-descubrimiento-prd.md and .../plantilla-prd-aplicacion.md. Follow the template exactly.
3. Existing context (not authority): docs/03-aplicaciones/lanzadera/ (epic.md, walkthrough-*.json, ERD/schema.sql, data-model.md).

## Source of truth
- Legacy repo DysTelefonica/Lanzadera pinned at `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17`. Source = `src/` only. Engram/epics/walkthroughs are context; contradictions become "Discrepancias abiertas".
- Clone read-only, sparse, into YOUR worktree's ignored dir:
  `git clone -q --filter=blob:none --no-checkout --sparse https://github.com/DysTelefonica/Lanzadera.git .atl/sources/lanzadera && cd .atl/sources/lanzadera && git sparse-checkout set src && git checkout -q 1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` (resolve the full SHA first if short).
- Index it: `codegraph init .` inside that clone (codegraph-vba 2.1.1 indexes VBA; event handlers show as `event-handler`). Use `codegraph query|explore|callers|node` there for control/event → handler → calls chains; confirm with the source lines. Never use the blueprint's own CodeGraph (Python) as legacy evidence.
- Every behavioural claim cites `src/...:line`. No citation ⇒ mark it pendiente. Never invent tables, columns, rules or relations. No business data, credentials or network paths.

## Workspace and delivery
- Worktree (first branch): `git -C /home/ubuntu/repos/access2web-blueprint fetch -q origin && git -C /home/ubuntu/repos/access2web-blueprint worktree add -b docs/prd-lanzadera-1 /home/ubuntu/repos/access2web-blueprint-worktrees/prd-lanzadera origin/main`. Work ONLY there. Never touch /home/ubuntu/repos/access2web-blueprint (main checkout) nor other worktrees.
- Output only under `docs/03-aplicaciones/lanzadera/prd/`. Do not edit epics, walkthroughs, other apps, CI, scripts, .agents/ or AGENTS.md.
- Language of the PRD: Castellano peninsular formal (usted), sentence case, no emojis. Commits/PR titles: English Conventional Commits with scope `lanzadera` (e.g. `docs(lanzadera): add PRD inventory and form cards G1`), no AI attribution.
- PR slices, each ≤400 changed lines (`git diff --numstat origin/main`), each a coherent slice: PR1 = prd/README.md (pin, reconciled inventory skeleton with every form/object listed as pendiente) + first form cards; next PRs = more form cards; then erd-origen.md, matriz-campos.md (may need several PRs), coexistencia.md, issues-propuestas.md, and a final PR updating README totals/delta.
- Independent PRs, not stacked: EVERY branch `docs/prd-lanzadera-<k>` starts from a fresh `origin/main` (git fetch first) and touches files DISJOINT from your other open PRs, so each PR diff against main stays ≤400 lines. Only prd/README.md is shared: create it in PR1 and modify it again only when PR1 is merged (check `gh pr view <n> --json state`); otherwise defer README edits to a later PR. To avoid broken links (docs-links gate), a file may only link to files already on main or in the same PR; the README index links to form cards only in a PR created after those cards are merged. Never rebase, --amend or force-push a published branch. Use one worktree per branch (e.g. prd-lanzadera-<k>) or switch branches in your worktree only when the tree is clean.
- PR body: follow .github/PULL_REQUEST_TEMPLATE.md. Use `Refs #1074` and label `chain:partial` on every PR except the final tip, which uses `Closes #1074` and full Chain Context. Add labels `type:docs` and `app/lanzadera`. Before each push: `git diff --check` and `python3 .agents/skills/ci-pattern/assets/docs-links/run_docs_links.py . --config .github/docs-links.json`.
- PR base: always `main` (never another PR branch).
- You MUST NOT: apply status:approved or any governance label other than type:docs/app/lanzadera/chain:partial, merge PRs, create issues, close issues, delete branches, touch Docker/containers, modify the legacy repo or any .accdb. The coordinator reviews and merges.
- If CI fails on your PR, fix it on the same branch with a new commit.

## Quality bar (this is what "perfect PRD" means)
- Inventory reconciles 100% of `src/forms/*.cls`, their `.form.txt` controls, modules, classes and saved queries: analizado / no aplicable (motivo) / pendiente. Pendientes are never counted as done.
- Every form card: controls+events table with handler line, calls, data effects, capability; normal/alternative/error flows; rules/validations/calculations; reads/writes with transaction notes; coexistence notes; three separated lists (comportamiento actual / posibles defectos / propuestas web pendientes de decisión); acceptance criteria traced to rows.
- Delta vs previous study: each capability añadida / modificada / retirada / sin cambios, with evidence.
- Work through ALL forms over successive PRs; do not stop after the first PR. Keep going until the tip PR closes #1074 or you hit a genuine blocker.

## Reporting
After each PR is opened, print one line: `PRD-lanzadera PR #<num> <k>/<total-estimate> <lines> lines`. At the end print a final report: PR list, totals analizado/no aplicable/pendiente, open discrepancies, blockers.
