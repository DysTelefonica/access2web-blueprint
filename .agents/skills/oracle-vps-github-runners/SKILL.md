---
name: oracle-vps-github-runners
description: Trigger: runner self-hosted, runner en Coolify, VPS Oracle, jobs encolados, runner offline, segundo runner, migrar repo a runners propios. Provision, harden and diagnose self-hosted GitHub Actions runners, and migrate repos onto them.
license: Apache-2.0
metadata:
  author: Andrés Román
  version: 2.1
  scope: ['universal', 'infra']
  last_verified: 2026-10-05
  auto_invoke: ['provisioning an Oracle VPS GitHub runner']
  tiers: ['universal', 'infra']
---



## Activation Contract

Load when provisioning, repairing, scaling or diagnosing a self-hosted runner
on the Oracle VPS, or migrating a repository onto that pool.

## Hard Rules

Rationale lives in the references; these are the constraints.

1. **Never let a failure become an answer.** Exit 127, a missing input, an empty
   context — behind `|| echo 0`, a default or a `skip()`, each becomes a
   plausible verdict. Fail the step. A long silent step is the same defect with a
   human reading it, and their failure mode is to cancel: heartbeat it.
2. **Never `RUNNER_TOKEN`** — it expires in an hour. `ACCESS_TOKEN`, minted at start.
3. **Never restart or watchdog a runner still on `RUNNER_TOKEN`.** Migrate first.
4. **One `_work` per runner**, identical path inside and outside the container.
5. **No second runner** while any job publishes a fixed host port.
6. **Identical labels, distinct `RUNNER_NAME`.** Labels route, not identify.
7. **Every job carries `timeout-minutes`**, sized on THIS runner. Default 360.
8. **No `gh`, no `python` on this image.** Add `setup-python`; use curl + jq.
9. **Read a service port in a step, never `jobs.<id>.env`** — empty, silently.
10. **Ship every requirement with a gate**, prove it fails on the real tree, and
    make it ignore comments.
11. **Neither `status=online` nor the API's runner count is proof.**
    Deregister and restart; cross-check every name to a container.
12. **Green PR checks are not a merge.** Read the post-merge runs.
13. **Pool size = the job graph's widest tier**, computed, not guessed.
14. **Visibility before migration.** Routing `pull_request`-reachable jobs to a
    runner that mounts the production `/var/run/docker.sock` is a critical risk
    while the repo is PUBLIC (any outsider's PR → root on the prod host, next to
    Coolify/postgres/adguard/vaultwarden) and only collaborator-level risk once
    private. Decide the visibility question FIRST, then migrate. Repo-side
    enforcement: `scripts/check_workflows.py::check_runner_isolation`
    (APAP_WEB issue #782).
15. **Migration is the PRIVATE-repo path.** Public = hosted minutes are free
    (`netAmount: 0`) → keep PR-reachable jobs hosted; private = minutes bill
    → migrate, via the seven-step recipe in `repo-migration.md`. Visibility
    is decided BEFORE the migration, never after — HR-14 is why.
16. **Every migrated job declares its full tooling contract and proves it
    on the runner that runs it.** Each binary its `run:` steps and
    containers invoke is declared as job data, and a fail-loud preflight
    step at the top of the job verifies each binary exists on the runner
    executing it (`command -v <bin> || exit 1`, naming the missing one).
    The canonical preflight is
    `personal/ardelperal/ci-pattern/assets/tooling-preflight/tooling-preflight.sh`
    (issue #165): it reads the declared list as versioned job data, verifies
    every binary with `command -v`, names ALL missing ones and fails (exit
    != 0) when the list is missing, unreadable or empty — an undeclared
    contract never passes silently.
    Verifying only on the historical runner proves nothing: the absence
    shows up on the first pool runner that lacks the binary (exit 127 on
    `envsubst` in `deploy`, run 37037386528 — the historical runner had it
    installed by hand, the rest of the pool did not). A preflight that
    does not declare the list MUST fail, not pass silently.
17. **The declaration is the only durable place for host state.** Every
    host-level fix (a package installed by hand, a label added ad hoc)
    MUST land in the service definition (compose/env/image) in the same
    session; manually installed state on a managed host MUST NOT be
    treated as durable — `--force-recreate` reverts it. A binary that
    exists only through a manual installation is absent for all purposes,
    and the preflight exposes it on the next run, not weeks later.

## Decision Gates

| Symptom | Read | Action |
|---|---|---|
| Offline after restart, or restart loop | FM-1 | Migrate to `ACCESS_TOKEN` |
| Still looping with a valid PAT | FM-5 | Drop reusage; recreate, not restart |
| More registrations than containers | FM-6 | Ghost — delete it |
| `online, busy=false` + queued | FM-2 | Labels first, then job count |
| Jobs queue instead of overlapping | FM-2 | `pool-sizing.md` |
| Mute hang on `docker run` | FM-3 | `timeout 30 docker info` |
| Overlap-only flakiness | FM-4 | Dynamic ports, per-runner `_work` |
| Exit 127, or a gate that never fails | R1 | `repo-migration.md` |
| Private repo burning hosted minutes (measured: 23,500 min/mo on APAP_WEB) | HR-15 | Migrate — seven-step recipe in `repo-migration.md` |

## Execution Steps

1. Read `references/vps-topology.md` and re-verify host facts before citing them.
2. Diagnose against `references/failure-modes.md` before changing anything.
3. **Provisioning or repairing**: use `assets/runner-compose.yml`, replacing
   every placeholder. Set `ACCESS_TOKEN` as a Coolify env var, never inline.
   Delete any leftover `RUNNER_TOKEN`. Verify per Rule 11.
4. **Migrating a repository**: follow `references/repo-migration.md` in order,
   applying the blocks it names from `assets/repo-ci-hardening.yml`. Blocks 1
   and 4 gate any future second runner.
5. Confirm one real run goes green, then the post-merge runs too (Rule 12).

## Output Contract

Report: runner name, labels, `_work`, auth, deregister-restart result,
hardening blocks still missing, and post-merge run conclusions.

## References

- `references/vps-topology.md` — host, runners, labels, billing facts
  (public repos ride free; the migration trigger is a visibility flip).
- `references/failure-modes.md` — FM-1..FM-6 and how to tell them apart.
- `references/repo-migration.md` — the visibility decision rule (public
  keeps PR jobs hosted, private migrates), the seven-step private-repo
  recipe, ordered migration, and writing gates that cannot quietly stop
  working.
- `references/pool-sizing.md` — how many runners, measured not guessed.
- `assets/runner-compose.yml` — verified Coolify template.
- `assets/repo-ci-hardening.yml` — the seven consumer requirements.
