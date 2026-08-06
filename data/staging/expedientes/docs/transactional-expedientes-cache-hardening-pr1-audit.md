# PR-1 Cache Mutation Audit

Change: `transactional-expedientes-cache-hardening`

Scope: PR-1 mapping/audit only. Later PRs will harden the unsafe groups; this PR adds the controlled cache failure seam and documents the transactionality baseline.

## Cache Refresh Entry Points

| Method | Cache scope | Transactionality classification | PR owner |
|---|---|---|---|
| `ExpedienteOperaciones.Registrar` | `Cabecera` | Safe top-level owner: starts DAO transaction, writes source rows, refreshes cache, commits/rolls back. | Existing |
| `ExpedienteOperaciones.RegistrarAlta` | `Cabecera` | Safe top-level owner, but `ActualizarNCs` still uses its legacy DB path. | PR-2 |
| `ExpedienteOperaciones.RegistrarExpEntidades` | configurable | Cache facade. Safe only when caller passes the transaction `p_db`; standalone use can fall back to `getdb()`. | PR-1 seam, later governance |
| `ExpedienteEntidadOperaciones.Registrar` | configurable | Cache writer. Safe with caller `p_db`; standalone fallback remains for compatibility. `ActualizarNCsContratistas` is outside caller `p_db` today. | PR-2 |
| `ExpedienteOperaciones.RegistrarComercial` / `EliminarComercial` | `Comerciales` | Uses optional `p_db`; standalone call is not transaction-owned. | PR-4 |
| `ExpedienteOperaciones.RegistrarHito` / `EliminarHito` | `Hitos` | Uses optional `p_db`; standalone call is not transaction-owned. | PR-4 |
| `ExpedienteOperaciones.EditarEnvioCorreoResponsable` / `EditarJPResponsable` / `EliminarResponsable` | `Responsables` | Uses optional `p_db`; standalone call is not transaction-owned. | PR-3 |
| `ExpedienteOperaciones.EliminarLugarEjecucion` | `Lugares` | Uses optional `p_db`; standalone call is not transaction-owned. | PR-3 |
| `ExpedienteOperaciones.EliminarPECAL` | `PECAL` | Uses optional `p_db`; standalone call is not transaction-owned. | PR-3 |
| `ExpedienteOperaciones.EliminarRAC` | `RACs` | Uses optional `p_db`; standalone call is not transaction-owned. | PR-3 |
| `ExpedienteOperaciones.EliminarSuministrador` | `Suministradores` | Uses optional `p_db`; standalone call is not transaction-owned. | PR-3 |
| `ExpedienteSuministradorServicio` cache refresh calls | `Suministradores` | Service owns DAO transaction and passes shared DB to cache facade. | Existing/PR-3 verification |
| `GuardadoAutomaticoHelper.ActualizarCache` | `Cabecera` | Explicit cache-only refresh through `getdb()`, not a source mutation. | Out of transactional mutation scope |

## PR-1 Seam

`ExpedienteEntidadOperaciones.Registrar` now accepts `p_ForceCacheFailureForTest`, default `False`, and also exposes explicit `TestOnly*` instance methods. When inactive, production behavior is unchanged. When tests activate the seam, the method returns the canonical `TEST_ONLY_CACHE_REFRESH_FAILURE` before any cache write.

Future rollback tests can pass the seam through `ExpedienteOperaciones.RegistrarExpEntidades` after source writes to prove source/cache atomicity per slice.
