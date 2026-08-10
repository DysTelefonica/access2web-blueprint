# app-lanzadera-mvp

> Hexagonal Python platform — first slice: the Lanzadera admin module.

This package is the Phase 0 foundation of the platform that replaces the legacy
Access/VBA `lanzadera.accdb` admin slice. The architecture is hexagonal
(`domain`/`ports`/`application`/`adapters`/`di`/`delivery`) and every layer
boundary is enforced by the gate at `scripts/check_layers.py`.

> **Naming deviation (2026-08-09)**: the on-disk directory is `app/` instead of
> the openspec `platform/`. Python's stdlib ships a `platform` module that
> coverage.py and other tools import; shadowing it from a same-named package
> broke `coverage` and `pip` at runtime. `app/` keeps the layout intent
> (hexagonal Python package for the platform) without colliding with stdlib.
> All gate constants, plugin paths and pytest configuration were updated to
> `app.*` accordingly.

## Status

| Phase | What it ships |
|---|---|
| 0 (this PR) | Platform scaffold, CI gates, `pytest_plugin/coverage_gate.py`, hexagonal layer gate |
| 1 | Pure domain entities (user, app, profile, assignment, reset_token, global_admin, audit) |
| 2 | Alembic 0001..0006 — schema + seeds from fixtures |
| 3 | Application use cases, including `CRITICAL_HELPERS` (hash/verify_password, issue/consume_reset_token) |
| 4 | Driven adapters — Postgres, Argon2id crypto, NationalIdCipher, mail queue, TTL cache, assume-in-office, BootstrapAdapter |
| 5 | Delivery — FastAPI + HTMX routes + Mistica templates + `gentle-ai platform user ...` CLI |
| 6 | Full test suite + quality gates (DA-13 AST sweep, lockout policy, same-transaction audit) |
| 7 | E2E on docker-compose |

See `openspec/changes/lanzadera-mvp/{proposal,specs,design,tasks}.md` for the change-level contract.

## Quick map

| Need | Open |
|---|---|
| Local setup | `docker-compose.yml` at the worktree root |
| Quality gates contract | `docs/calidad-de-codigo-y-ci.md` |
| Architectural decisions | `openspec/changes/lanzadera-mvp/design.md` |
| Gate scripts | `scripts/check_*.py` at the worktree root |
| Coverage gate plugin | `platform/pytest_plugin/coverage_gate.py` |

## Running

```bash
make lint typecheck test check-layers check-complexity check-crap check-dry \
     check-branch-name quality-report
```

Every target invokes a real, pinned tool or script. Wrappers (`|| true`,
`continue-on-error`) are forbidden — see `tests/test_ci_workflow.py`.

## Licence

Proprietary. See repository root for the licence file.