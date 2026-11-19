---
name: a2w-local-dev
description: 'Trigger: probar una issue en local, levantar el stack, tests unitarios, integración, e2e, make verify, WSL. Per-issue local stack of access2web, the same on Linux and on Windows through WSL.'
license: Apache-2.0
metadata:
  author: ardelperal
  version: "1.0"
---

# access2web local dev

## Activation Contract

Use this skill to test an access2web issue locally before pushing: unit tests, the integration suite against Postgres, the app with its migrations, the Playwright smoke and `make verify`. It is the local half of `a2w-issue-flow` (#1776), the counterpart of cadete's `cadete-local-dev`. Owned by this repository, not propagated.

## Hard Rules

- **HR-1** Everything runs in containers: Python 3.12.11 with `app/requirements-dev.lock` and `jq`, Postgres 16, the image of `app/Dockerfile` and Playwright at the version locked in `e2e/package-lock.json`. The host needs only Docker, bash and curl.
- **HR-2** One stack per issue: compose project `a2w<n>`, app on `28000 + n % 1000`, Postgres on `25000 + n % 1000`. Two issues run at once.
- **HR-3** No secret file is read: the app gets the development values the CI uses (`ci.yml`, integration and e2e-smoke). Never point the stack at a real database.
- **HR-4** On Windows, run it from WSL: a worktree created by the Windows git has a `gitdir: C:/...` that WSL cannot open, so the script exports `GIT_DIR` and `GIT_WORK_TREE`.

## Execution Steps

From the worktree root (`bash skills/a2w-local-dev/assets/dev.sh <n> <cmd>`):

| Command | What it does | CI job it mirrors |
| --- | --- | --- |
| `up` | Postgres, `alembic upgrade head`, the app image; waits for `/health` | `e2e-smoke` (app start) |
| `test [expr]` | `pytest -c app/pyproject.toml --rootdir=app -m "not integration"` (optionally `-k expr`) | `quality` (tests) |
| `integration` | Postgres, migrations and `pytest -m integration` with `APAP_INTEGRATION_ENABLED=1` | `integration` |
| `e2e [specs]` | Playwright against the running app (`BASE_URL=http://app:8000`); after `up` | `e2e-smoke` |
| `verify` | `bash scripts/local-preflight.sh`, the same as `make verify` | the local preflight |
| `down` | Removes the stack and its volumes | — |
| `env` | Prints the project, the ports and the git translation | — |

## Decision Gates

| Situation | Action |
| --- | --- |
| A port is taken | `down` the other stack of the same issue number modulo 1000, or free the port |
| `requirements-dev.lock` changed | Nothing: the tools image is rebuilt for the new lock content |
| The app does not answer `/health` | Read the logs the script prints; a migration failure shows in `up` |

## Output Contract

Report the commands run, their exit codes and the test totals.

## References

- `skills/a2w-local-dev/assets/dev.sh` — the script; `tests/test_a2w_local_dev.py` — its tests.
- `.github/workflows/ci.yml` — the jobs it mirrors.
- `skills/a2w-issue-flow/SKILL.md` — the end-to-end flow.
