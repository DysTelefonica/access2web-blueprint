# cache-idempotent-warmup — Cache_Warmup_Opciones wrapper, granular centralizado, TDD

> Single entry point for warming all cache tiers. Replaces ad-hoc invocations scattered across `InicializadorCache.PrepararStagingConCaches`, `Cache_Indicadores_ReconstruirTodo`, `CacheNCProyecto.ReconstruirListadoEstados`, `NCAuditoriaListadoCache.RebuildNCAuditoriaListadoCache`, `EstadoCatalogoBootstrap.BootstrapEstadoCatalogo`. Addresses issues #88, #89, #90, #92, #93, supersedes the ad-hoc workaround tracked in #91.

## Status

| Field | Value |
|-------|-------|
| **Current** | `active` (W1 written 2026-06-17, tests pending user recompile + `dysflow_test_vba`) |
| **Last verified** | — |
| **Manifest drift** | `clean` (extending existing `tests.vba.cache-warmup.json`, not creating a new manifest) |
| **Staging reachability** | `not-reachable` (work in progress on `feature/cache-idempotent-warmup-2026-06-17`) |
| **TDD evidence** | `thin` — code written 2026-06-17, but `dysflow_test_vba` not yet run against the new 8 procedures |
| **Last verified commit** | — |
| **Last verified at** | — |
| **Test evidence** | — |
| **Staging integration commit** | — |
| **Evidence updated at** | 2026-06-17 (code written; awaiting test_vba) |

## Release Tracking

| Field | Value |
|-------|-------|
| **UAT status** | `pending` |
| **UAT tag** | `PRUEBAS-004` (next after `PRUEBAS-003`) |
| **UAT date** | — |
| **UAT evidence** | — |
| **UAT tag history** | `PRUEBAS-001` → `PRUEBAS-002` → `PRUEBAS-003` (this change → PRUEBAS-004) |
| **Approved UAT tag** | — |
| **Production release tag** | — |
| **Production release commit** | — |
| **Production date** | — |
| **Rollback release tag** | revert merge commit on staging |

## Business Behavior

Operators and UAT scripts need a single, predictable entry point to warm the cache system before testing or UAT rounds. Today the operator must remember which function warms which table and which ones to skip (e.g. "don't warm audit cache if I'm only testing proyecto"). This feature unifies all warming under one function with a CSV selector and dry-run preview, so a UAT round can be prepared with one line: `? Cache_Warmup_Opciones("")`. The wrapper is idempotent so it can be re-run on the same backend without side effects, and dry-run capable so a typo can't corrupt the cache.

## Acceptance Criteria

- [ ] `Cache_Warmup_Opciones("")` on a fresh sandbox populates all 6 tiers and returns JSON with `ok=true` for each, `errores=[]`.
- [ ] `Cache_Warmup_Opciones("LISTADO")` populates only LISTADO; the other 5 tiers appear in `vias_omitidas`.
- [ ] `Cache_Warmup_Opciones("-INDICADORES_AUDITORIA")` populates 5 tiers, omits INDICADORES_AUDITORIA.
- [ ] `Cache_Warmup_Opciones("", True, True)` (dry-run) does not write; row counts before/after identical.
- [ ] `Cache_Warmup_Opciones("", False, ...)` skips EnsureSchema (faster re-sync; assumes schema ready).
- [ ] Called twice in a row on the same backend does not duplicate rows.
- [ ] JSON return is parseable by `JsonConverter.ParseJson` (round-trip test).
- [ ] 8 TDD tests in `tests/tests.vba.cache-warmup.json` pass green.
- [ ] `EnsureCacheIndicadoresSchema` creates `TbCacheIndicadoresAuditoriaHeader` + `TbCacheIndicadoresAuditoriaDetalle` if missing (#93).
- [ ] Feature page and UAT docs produced.

## Required Tests

| Procedure | Manifest | Status |
|-----------|----------|--------|
| `Test_Cache_Warmup_Opciones_VacioPueblaTodo` | `tests/tests.vba.cache-warmup.json` | pending (W1) |
| `Test_Cache_Warmup_Opciones_FiltrarPorInclusion` | `tests/tests.vba.cache-warmup.json` | pending (W1) |
| `Test_Cache_Warmup_Opciones_FiltrarPorExclusion` | `tests/tests.vba.cache-warmup.json` | pending (W1) |
| `Test_Cache_Warmup_Opciones_DryRunNoEscribe` | `tests/tests.vba.cache-warmup.json` | pending (W1) |
| `Test_Cache_Warmup_Opciones_EnsureSchemaTrueCreaTablasFaltantes` | `tests/tests.vba.cache-warmup.json` | pending (W1) |
| `Test_Cache_Warmup_Opciones_EnsureSchemaFalseAsumeSchemaListo` | `tests/tests.vba.cache-warmup.json` | pending (W1) |
| `Test_Cache_Warmup_Opciones_Idempotente` | `tests/tests.vba.cache-warmup.json` | pending (W1) |
| `Test_Cache_Warmup_Opciones_JsonEsValido` | `tests/tests.vba.cache-warmup.json` | pending (W1) |
| (existing) `Test_Cache_PrepararStaging_EmptySandbox_NoCrash_Atomic` | `tests/tests.vba.cache-warmup.json` | PASS (PRUEBAS-003) |
| (existing) `Test_Cache_PrepararStaging_Idempotente_SegundaLlamadaOk_Atomic` | `tests/tests.vba.cache-warmup.json` | PASS (PRUEBAS-003) |
| (existing) `Test_Cache_PrepararStaging_Resiliencia_CacheFallidaNoAborta_Atomic` | `tests/tests.vba.cache-warmup.json` | PASS (PRUEBAS-003) |

## Last Known Passing

| Field | Value |
|-------|-------|
| **Date** | — (after W1.7) |
| **Commit** | — |
| **Manifest** | `tests/tests.vba.cache-warmup.json` |
| **Result** | — |

## Integration Commits

| SHA | Message | Ancestor of staging |
|-----|---------|---------------------|
| `adf56ff` | `feat(cache): add Cache_Warmup_Opciones wrapper + 8 TDD tests` | No (work in progress on `feature/cache-idempotent-warmup-2026-06-17`) |
| `f34b827` | `chore(cache): add stub BootstrapEstadoCacheWarmup for backward-compat` | No |
| `08a5a4a` | `docs(cache): add spec change cache-idempotent-warmup-2026-06-17` | No |
| — | (W2 commit pending) | No |
| — | (W3 commit pending) | No |
| — | (W4 commit pending) | No |

## Access Sync Status

- **Import method**: Dysflow `import_modules` for the new `ModuloCacheWarmupWrapper.bas`; ensure `EnsureCacheIndicadoresSchema` extension is added to `ModuloCacheIndicadores.bas`.
- **Manual compile**: user compiles manually in Access VBE before any `test_vba` run. **No `compile_vba`** invoked by the orchestrator.
- **verify_binary**: to be run after first import; expected clean.

## Rollback Anchor

Revert the merge commit on `staging`. No destructive changes (Ensure* is opt-in; migration extension is additive with `On Error Resume Next`).

## Business Rules

The wrapper MUST preserve all existing cache semantics defined in `openspec/specs/cache-trust/spec.md` (issue #39). The wrapper MUST NOT change the read-path contract: `Cache_IndicadoresProyectoMaterializado_CargarConteos` and friends continue to be the read API. The wrapper only orchestrates the warmup/rebuild surface.

## Legacy Not to Copy

- Do not introduce `Debug.Print` as UI feedback. Use `Immediate` window output only.
- Do not couple the wrapper to specific forms or `Screen.ActiveForm`.
- Do not introduce tempvar-based state for the wrapper's return; it MUST be a pure function of inputs.

## Migration Notes

Web migration consideration: the wrapper is `Access.CurrentDb`-bound. For a web migration, the equivalent is a `POST /cache/warmup` admin endpoint with the same JSON contract. The tier names map cleanly to backend modules.

## Open Decisions

- Whether to expose `WARMUP_DRY_RUN_TIMEOUT` as a constant in the wrapper config (deferred — dry-run is fast enough on a sandbox that we don't need a timeout).
- Whether the wrapper should accept `p_MaxDuration` to abort long warmups (deferred — current UAT scale doesn't need it).

## Evidence Sources

- `openspec/changes/cache-idempotent-warmup-2026-06-17/proposal.md` — proposal.
- `openspec/changes/cache-idempotent-warmup-2026-06-17/tasks.md` — work units.
- `openspec/changes/cache-idempotent-warmup-2026-06-17/specs/cache-idempotent-warmup/spec.md` — delta spec.
- `openspec/specs/cache-trust/spec.md` — base spec (issue #39).
- `docs/features/cache-management/indicator-issues-cleanup.md` — predecessor feature (Issue #18).
- GitHub issues #88, #89, #90, #91, #92, #93.
- `docs/uat/PRUEBAS-003/` — last UAT round (this change → PRUEBAS-004).

## Post-Test Documentation Gate

> **Rule**: Integration is not done until this section is updated. After staging integration and passing tests, update the Status section fields (`last_verified_commit`, `last_verified_at`, `test_evidence`, `staging_integration_commit`, `evidence_updated_at`) before declaring the work complete.

| Step | Action | Done |
|------|--------|------|
| 1 | Tests pass against staging HEAD | [ ] |
| 2 | `last_verified_commit` updated with SHA | [ ] |
| 3 | `last_verified_at` updated with ISO datetime | [ ] |
| 4 | `test_evidence` updated with manifest + pass/total | [ ] |
| 5 | `staging_integration_commit` updated with merge SHA | [ ] |
| 6 | `evidence_updated_at` updated with current datetime | [ ] |
| 7 | Feature status reflects current state | [ ] |
