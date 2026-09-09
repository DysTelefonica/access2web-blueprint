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
| `check_dry.py` | QC-11 | DRY: no duplicación de conocimiento por módulo | `python scripts/check_dry.py` |
| `check_legacy_hashes.py` | QC-5, DA-13 | Pin AST rechaza `legacy_hash`, `sha256`, `migrate_password` (D88+D89) | `python scripts/check_legacy_hashes.py` |
| `check_mutation_sites.py` | — | Lista de sitios donde se ejecuta mutación semanal | `python scripts/check_mutation_sites.py` |
| `check_mutation.py` | — | Corre mutación semanal; falla si la mutation score cae | `python scripts/check_mutation.py` |
| `quality_report.py` | QC-11 | Agrega envelopes de todos los gates en `quality.report.json` | `python scripts/quality_report.py` |
| `check_walkthrough_schema.py` | — | MUST fields del template `walkthrough.json` presentes en `docs/03-aplicaciones/*/walkthrough-*.json` (53 archivos) | `python scripts/check_walkthrough_schema.py` |
| `app/pytest_plugin/coverage_gate.py` | QC-5 | `--cov-fail-under=69` global + cuatro targets auth exactos a 100 % | activado por `pytest --cov` |

> **QC mapping incompleto**: la tabla arriba es best-effort. El catálogo QC-1..QC-18 vive en `openspec/changes/lanzadera-mvp/design.md`. Este doc no es la fuente; se cruza contra el design para validar la asignación.

## Los 4 workflows de `.github/workflows/`

| Workflow | Cuándo corre | Qué hace |
|---|---|---|
| `ci.yml` | cada PR + push a main | Orquesta: pip-audit, gitleaks, trivy-config, ruff/mypy/pytest, los 13 check_*.py. |
| `security.yml` | cada PR + push a main | Fast subset de seguridad: pip-audit, gitleaks, trivy config. |
| `security-deep.yml` | semanal (cron) | Trivy filesystem + image, mutation semanal. |
| `release.yml` | tag `v*` pushed | Gate de identidad + verify checksum; ata al release pipeline. |

Los SHA de las actions se pinean vía `check_workflows.py`; actualizarlos requiere PR explícito.

## Cómo correrlo en local

```bash
# Workflow completo (lo que corre CI por PR)
make lint typecheck test check-layers check-complexity check-dry \
     check-branch-name quality-report

# Cada target invoca un tool pinned o un check_*.py.
# Wrappers (|| true, continue-on-error) están prohibidos
# — ver tests/test_ci_workflow.py.
```

(Nota: `make` no está disponible en Git Bash en Windows. Invocar los scripts con `python scripts/<nombre>.py` directamente.)

## Protección de `main`

`main` tiene branch protection desde el 2026-09-09. GitHub exige un PR
actualizado, conversaciones resueltas y estos checks en verde:

| Workflow | Check requerido |
|---|---|
| `ci.yml` | `quality` |
| `ci.yml` | `review-budget` |
| `security.yml` | `pip-audit` |
| `security.yml` | `gitleaks` |
| `security.yml` | `trivy-config` |

La protección se aplica a administradores y bloquea force-push y borrado de
`main`. `merge-ready` es informativo: no agrega los otros jobs y no sustituye a
los checks protegidos.

Detalle operativo en [`AGENTS.md` §Hard rule del CI](../../AGENTS.md).

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

## Core invariants

- **SHA de Actions pineados**: las versiones de actions de terceros en `.github/workflows/*.yml` van fijadas por SHA de 40 hex. `scripts/check_workflows.py` enforza esto y exige un `concurrency.group` por job. Actualizar requiere PR explícito.
- **Wrappers prohibidos**: `|| true`, `continue-on-error`, y cualquier otro wrapper que silencie un fallo están prohibidos en `ci.yml` y en los `check_*.py`. `tests/test_ci_workflow.py` pinea el contrato.
- **Cuatro targets auth a 100 % (DA-2, DA-4 y QC-5)**: `CredentialHasherArgon2id.hash`, `CredentialHasherArgon2id.verify`, `issue_reset_token` y `consume_reset_token`. El plugin falla el build con `session.exitstatus = 1` si cualquiera queda infracubierto.
- **Migraciones aditivas (D82)**: cada release Alembic es aditiva. El rollback es `DROP SCHEMA <módulo> CASCADE;` con el legacy intacto. No se permiten `DROP COLUMN`, `ALTER` destructivos ni `RENAME` en la misma release.
- **Mutation semanal fuera de PR**: `check_mutation.py` corre desde `security-deep.yml` por cron semanal, no por commit. Cuando la mutation score cae del umbral, el workflow falla y crea issue automático; no bloquea PRs individuales.

## Contributor checklist

- [ ] Si el PR añade un `check_*.py`, el script cita el QC y la decisión D-/DA- en el docstring, y existe un test de smoke (`tests/test_gate_smoke.py`).
- [ ] Si el PR modifica `.github/workflows/`, las Actions nuevas van pineadas por SHA de 40 hex y cada job declara `concurrency.group`.
- [ ] Si el PR toca `pyproject.toml` (ruff/mypy/argon2), las versiones quedan pinned en `>=X,<Y+1` y `scripts/install-skills.sh` sigue corriendo.
- [ ] Si el PR añade cobertura a `pytest --cov`, el threshold `--cov-fail-under` se mantiene o sube; nunca baja.
- [ ] El `make lint typecheck test check-layers check-complexity check-dry check-branch-name quality-report` corre verde en local antes de push.
- [ ] Si se cambia un comando de la sección §Cómo correrlo en local, también se actualiza `Makefile` y el `Makefile` no introduce wrappers.
- [ ] Si se introduce una dependencia nueva en `app/pyproject.toml`, `pip-audit` corre verde y se documenta en la sección §Stack de [`docs/architecture.md`](architecture.md).

## Navigation

Previous: [architecture.md](architecture.md) | Next: [AGENT-SETUP.md](AGENT-SETUP.md)

---

[Next: CHANGELOG →](../../CHANGELOG.md)
