# apply-progress — Lanzadera MVP fase apply (issue #623)

Registro de evidencia TDD de la fase apply de `lanzadera-mvp`. Lista
los PRs mergeados por work-unit con su SHA, su commit y el sub-issue
del que depende, para que el revisor pueda reconstruir el orden de
integración desde el ledger de Git.

## Tabla de progreso

| WU | Issue | PR | SHA | Título |
|---|---|---|---|---|
| E2 | #52 | #744 | `d163a55` | feat(lanzadera): migrate_from_access.py smoke script |
| F0 | #23 (cerrado duplicado) | — | `68db5ae` | chore: Dockerfile + docker-compose.yml (trabajo previo) |
| F1 | #222 | #646 | `6a987b5` | feat(exp): ports layer F01 |
| F2 | #224 (cerrado duplicado) | #647 | `e7ec475` | feat(exp): C01 create_expediente use case |
| W62 | #595 | #595 | `ec4c694` | test: W-TEST cleanup |
| W60 | #600 (cerrado por ya implementado) | — | — | feat: real-time session presence via SSE |
| W61 | #601 (cerrado por ya implementado) | — | — | feat: app CRUD admin endpoints G2 |
| migrations | #35-39 (cerrados como duplicados) | — | — | migrations 0001..0006 |
| contracts | #622 (cerrado por ya implementado) | — | — | repository contract tests |
| quality gates | #53 (cerrado por ya implementado) | — | — | final quality gates verification |

## Work units pendientes

| WU | Issue | Estado |
|---|---|---|
| E2 (sub-issue específico) | #625 | cerrado (post-MVP, .accdb reader diferido) |
| docs | #656 | aprobado, pendiente |

## Cómo leer esta tabla

Cada fila es un PR mergeado a `main` con su SHA corto (`git log --oneline | grep <sha>`). Los
PRs marcados como "cerrado por ya implementado" no requirieron nuevo PR — el trabajo ya
estaba en `main` cuando el ciclo del issue los alcanzó; el cierre sólo formaliza el ledger.

## Validación

Para confirmar la integridad del registro:

```
git log --oneline | head -20
```

debe mostrar todos los SHAs listados arriba en orden cronológico inverso.
