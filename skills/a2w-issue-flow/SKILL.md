---
name: a2w-issue-flow
description: 'Trigger: trabajar una issue de access2web, abrir la PR, mergear, cadena de PRs, size:exception, cerrar la issue. Take an access2web issue end to end, from the issue to the governed merge, without waiting on anyone.'
license: Apache-2.0
metadata:
  author: ardelperal
  version: "1.0"
---

# access2web issue flow

## Activation Contract

Use this skill for any access2web issue, from the moment it is picked up until its PR is merged and the local worktree is gone. It applies the owner's standing order in the local part of `AGENTS.md` (#1772): the agent does everything itself, acting on GitHub with the profile of the repository admin who runs the session (ardelperal or juanfelixrt23). No role, bot, App or infra team takes part. It is the same flow as DysTelefonica/cadete (`cadete-issue-flow`), adapted to access2web: GitHub-hosted runners and the governed merge with the workflow token (#1770). It is owned by this repository and is not propagated from team-skills.

## Hard Rules

- **HR-1** No wait on anyone: no coordinator, reviewer, maintainer or infra role is asked for approval, and the agent coordination skills are not loaded here. Where a propagated skill or block says otherwise (`github-roles`, `governed-delivery`, «decisión humana»), the standing order wins.
- **HR-2** Every `gh` command runs with the profile of the admin who runs the session: `GH_TOKEN= GITHUB_TOKEN= gh ...` (on Windows, `"/c/Program Files/GitHub CLI/gh.exe"`). Never print or copy a token; secret values never go through the agent. A collaborator with write access stops at the PR.
- **HR-3** Merge through `gh workflow run governed-merge.yml -R DysTelefonica/access2web-blueprint -f pr=<N>`, once every check of the PR's current head has finished green. Never `--admin`, never force-push, rebase or `--amend` on a published branch.
- **HR-4** Edit, `git add`, commit and push are sequential steps. Push by SHA (`git push origin <sha>:refs/heads/<branch>`) and check that the PR's `headRefOid` moved: `git push origin HEAD:<branch>` from a worktree has reported «Everything up-to-date» without pushing (cadete#1191 F18).
- **HR-5** The repository is public: CI runs on GitHub's runners only (`docs/10-runners.md`, `scripts/check_workflows.py`). Never add a self-hosted runner to a PR job.

## Execution Steps

1. **Issue**: read it with `gh issue view`; a new one carries the six sections of `.github/ISSUE_TEMPLATE/issue-canonical.yml` and a `type:*` label.
2. **Worktree**: `git worktree add ../access2web-blueprint-worktrees/<n>-<topic> -b <type>/<n>-<slug> origin/main` (branch pattern `P01_branch_name_pattern` of `ci-pattern.yaml`).
3. **Build test-first**; run `make verify` (the local preflight) before pushing.
4. **PR**: push, then `gh pr create` with `Closes #<n>`, the table criterion → files → test under `## Tests que prueban el cierre`, and `## Asistencia de IA`. Over 400 changed lines: a chain, each intermediate link `chain:partial` with `Refs #<n>`, the tip with `Closes`.
5. **CI**: wait until every check of the current head has finished (`gh pr checks <N>`, no `pending`), all green.
6. **Merge**: governed merge (HR-3), then `gh run watch <id> --exit-status` and `gh pr view <N> --json state`. In a chain, after each merge retarget the next link (`gh pr edit <M> --base main`), merge `origin/main` into its branch, push by SHA and wait for its full CI.
7. **Cleanup**: once the PR is `MERGED`, `git worktree remove` and `git branch -D` of the local branch; the remote branch stays.
8. **Progress report**: when the closed issue, in any repository, carries `saneamiento:coexistencia` or changes a phase or milestone of `docs/07-migracion/seguimiento/fases.json`, run `python scripts/migration_progress.py` from an up-to-date `main` and merge the new cut through its own `docs(migracion)` PR the same day (#1825). The owner requires the progress of every phase to be current in `main`.

## Decision Gates

| Situation | Action |
| --- | --- |
| The governed merge exits non-zero for a real finding (red check, PR contract, closing words in an intermediate link) | Fix what it names in the branch and dispatch again |
| The governed merge cannot merge the PR that fixes the governed merge itself, or reports checks `missing` that the API shows green on the head | Merge that PR once with the admin's profile: `gh pr merge <N> --merge --subject "Merge PR #<N> (governed merge)"`, with every check green. The owner's standing authorization (2026-11-19); `push-compliance` accepts the marker |
| A link cannot be split below 400 lines without breaking `main` | Answer the four questions of `## Excepción de tamaño` in the PR body, then label it `size:exception` yourself |
| A check is red | Fix it in the branch; a red check is never merged |
| An issue of the sanitation closes in another repository (Expedientes, Automatización…) | The agent of that repository tells the owner; the access2web session regenerates the report (step 8) |
| `migration_progress.py` exits 2 | An issue sits in a repository that no group of `fases.json` maps: map it in `fases.json` in the same PR, never edit `avance.md` by hand |

## Output Contract

Report per issue: the PRs and their merge commits, the CI result of each, the governed merge run ids (or the fallback merge and why), and what was cleaned up.

## References

- `AGENTS.md`, local part — the owner's standing order (#1772).
- `docs/agents-repo-guide.md` — delivery and governed merge rules.
- `.github/workflows/governed-merge.yml`, `.github/host-contract.json` — the governed merge with the workflow token (#1770).
- `docs/10-runners.md` — why CI stays on GitHub's runners.
