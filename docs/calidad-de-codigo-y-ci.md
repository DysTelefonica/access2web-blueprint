[← Back to architecture.md](architecture.md)

# access2web-blueprint — Calidad de código y CI

**Esta guía documenta los gates de calidad y el contrato CI del MVP de plataforma. Las decisiones arquitectónicas que motivan cada gate viven en [`docs/architecture.md`](architecture.md) — aquí están los comandos, los QC-<n> que atienden, y los workflows que los orquestan.**

## Sentence que organiza

> «La calidad se enforza en CI, no en revisión. Cada gate tiene un comando reproducible, un QC que documenta su contrato, y un test que pinea el wiring contra drift.» — `openspec/changes/lanzadera-mvp/design.md` §Pipeline de calidad

## What this is / is not

| Es | No es |
|---|---|
| Catálogo de los 12 `check_*.py` + sus QC.<br>Contrato de los 4 workflows de `.github/`.<br>Cómo se invoca cada gate localmente. | Réplica de [`docs/architecture.md`](architecture.md) §CI gates.<br>Manual de uso de dysflow ni de las migraciones Alembic.<br>Política del revisor humano (eso vive en [`AGENTS.md`](../../AGENTS.md)). |

## Contrato global

```text
PR abierto contra main
       │
       ▼
┌──────────────────────────────┐
│  ci.yml (orquesta todo)      │
│   ├─ pip-audit               │
│   ├─ gitleaks                │
│   ├─ trivy-config            │
│   ├─ ruff                    │
│   ├─ mypy                    │
│   ├─ pytest --cov            │
│   └─ scripts/check_*.py * 12 │
└──────────────────────────────┘
       │
       ▼ (revisión humana)
PR mergeado con --squash; rama remota conservada
```

Tres grupos de gates:
- **Estáticos**: `ruff`, `mypy` (de pyproject), `gitleaks`, `pip-audit`, `trivy-config`.
- **Tests**: `pytest --cov` con el plugin `coverage_gate.py` (QC-5).
- **De contrato**: los 12 `check_*.py` que pinean invariantes arquitectónicas.

## Los 12 `scripts/check_*.py`

Estos viven en `scripts/check_*.py` y se invocan desde `ci.yml` por PR, y semanalmente donde aplica. La tabla mapea el gate al QC que documenta su contrato.

| Gate | QC | Decisión arquitectónica | Comando local |
|---|---|---|---|
| `check_branch_name.py` | QC-6 | Branch naming `<tipo>/<nº>-<kebab-slug>` (AGENTS.md) | `python scripts/check_branch_name.py` |
| `check_pr_size.py` | QC-6 | 400 líneas `additions + deletions` (CONTRIBUTING.md) | `python scripts/check_pr_size.py` |
| `check_workflows.py` | QC-9 | Actions fijadas por SHA de 40 hex; `concurrency.group` por job (AGENTS.md) | `python scripts/check_workflows.py` |
| `check_layers.py` | QC-2, QC-9 | `ROOT_PACKAGE = "app.src.modules"`; slicing vertical prohibido entre módulos (DA-1) | `python scripts/check_layers.py` |
| `check_complexity.py` | QC-1, QC-10 | Techo de complejidad ciclomática 15 (DA-1) | `python scripts/check_complexity.py` |
| `check_crap.py` | QC-11 | CRAP score ≤ 6; pendiente decisión v0.2 sobre el techo efectivo (issue #266) | `python scripts/check_crap.py` |
| `check_dry.py` | QC-11 | DRY: no duplicación de conocimiento por módulo | `python scripts/check_dry.py` |
| `check_legacy_hashes.py` | QC-5, DA-13 | Pin AST rechaza `legacy_hash`, `sha256`, `migrate_password` (D88+D89) | `python scripts/check_legacy_hashes.py` |
| `check_mutation_sites.py` | — | Lista de sitios donde se ejecuta mutación semanal | `python scripts/check_mutation_sites.py` |
| `check_mutation.py` | — | Corre mutación semanal; falla si la mutation score cae | `python scripts/check_mutation.py` |
| `quality_report.py` | QC-11 | Agrega envelopes de todos los gates en `quality.report.json` | `python scripts/quality_report.py` |
| `app/pytest_plugin/coverage_gate.py` | QC-5 | `--cov-fail-under=69` por paquete + CRITICAL_HELPERS a 100 % | activado por `pytest --cov` |

> **QC mapping incompleto**: la tabla arriba es best-effort. El catálogo QC-1..QC-18 vive en `openspec/changes/lanzadera-mvp/design.md`. Este doc no es la fuente; se cruza contra el design para validar la asignación.

## Los 4 workflows de `.github/workflows/`

| Workflow | Cuándo corre | Qué hace |
|---|---|---|
| `ci.yml` | cada PR + push a main | Orquesta: pip-audit, gitleaks, trivy-config, ruff/mypy/pytest, los 12 check_*.py. |
| `security.yml` | cada PR + push a main | Fast subset de seguridad: pip-audit, gitleaks, trivy config. |
| `security-deep.yml` | semanal (cron) | Trivy filesystem + image, mutation semanal. |
| `release.yml` | tag `v*` pushed | Gate de identidad + verify checksum; ata al release pipeline. |

Los SHA de las actions se pinean vía `check_workflows.py`; actualizarlos requiere PR explícito.

## Cómo correrlo en local

```bash
# Workflow completo (lo que corre CI por PR)
make lint typecheck test check-layers check-complexity check-crap check-dry \
     check-branch-name quality-report

# Cada target invoca un tool pinned o un check_*.py.
# Wrappers (|| true, continue-on-error) están prohibidos
# — ver tests/test_ci_workflow.py.
```

(Nota: `make` no está disponible en Git Bash en Windows. Invocar los scripts con `python scripts/<nombre>.py` directamente.)

## Lo que el revisor humano hace

El CI no bloquea el merge automáticamente (decisión del 2026-08-10: GitHub Team cuesta $4/user/mes, no aprobado). El revisor verifica manualmente:

1. `ci / quality` verde.
2. `ci / review-budget` verde.
3. Si `ci / security` falla **solo** por paquetes externos sin relación con el cambio, abrir issue y mergear con un follow-up PR.
4. Si cualquier otro check falla, pedir arreglo antes de mergear.

Detalle en [`AGENTS.md` §Hard rule del CI](../../AGENTS.md).

## Hexagonal layer gate (QC-2, QC-9)

`check_layers.py` es el gate más estructural. Enforza:

| Capa | Pureza | Puede importar | Prohibido |
|---|---|---|---|
| `domain/` | PURE | sólo el propio módulo | frameworks, ports, adapters |
| `ports/` | PURE | `typing.Protocol` | implementaciones concretas, frameworks |
| `application/` | PURE | domain + ports | adapters, delivery |
| `adapters/` | driven | domain, ports, frameworks | delivery |
| `di/` | composition root | todos | — |
| `delivery/` | driving | application, ports, adapters | domain |
| `shared/` (cross-cutting) | mixto | cache, audit, lockout | slicing vertical entre módulos |

El test `tests/test_ci_workflow.py` pinea el wiring (`ROOT_PACKAGE`, `ALLOWED_IMPORTS`, `PURE_LAYERS`) contra drift.

## Mutation semanal

`check_mutation.py` corre semanalmente (no por PR). Detecta:

- Lógica auth (hash, verify, reset, bootstrap) — debe mantener score alto por DA-2 + DA-13.
- Reglas de mapping legacy (DA-12) — la cardinalidad exhaustiva de `SinAcceso` exclusivo.
- Servicios cross-cutting (CachePort invalidate, AuditLogPort append atómico, DA-11).

Cuando la mutation score cae por debajo del umbral, `security-deep.yml` falla y crea issue automático.

## Cómo añadir un nuevo gate

1. Crear `scripts/check_<nombre>.py` con docstring que cite el QC + decisión.
2. Añadirlo a la tabla §Los 12 check_*.py de este doc.
3. Añadirlo al matrix de `tests/test_ci_workflow.py` (pin AST).
4. Cablearlo en `ci.yml` con timeout explícito.
5. Subir un PR; el revisor valida que el gate tenga un QC documentado y un test de smoke en `tests/test_gate_smoke.py` (QC-18).

---

[Next: CHANGELOG →](../../CHANGELOG.md)
