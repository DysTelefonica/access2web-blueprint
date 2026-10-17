[← Volver al Codebase Guide](../CODEBASE-GUIDE.md)

# Testing y CI

> Esta página rastrea `main`, que puede contener cambios sin liberar. Para la última versión liberada, véase la etiqueta correspondiente.

Este repo separa el veredicto en **cinco niveles de disparo** sobre dos workflows — PR, post-merge, nightly, manual y release — y cada nivel tiene un contrato propio. La superficie común está en `.github/workflows/ci.yml` (PR + post-merge + nightly + manual) y `.github/workflows/release.yml` (release); el contrato manual de los quality gates vive en `docs/gates.md`.

## Mapa de niveles

| Nivel | Workflow | Disparador | Jobs | Veredicto |
|---|---|---|---|---|
| **PR** | `.github/workflows/ci.yml` | `pull_request` (cualquier rama) | `pr-gate`, `quality`, `integration`, `security`, `codeql`, `required` | Branch protection exige `required` verde. |
| **Post-merge** | `.github/workflows/ci.yml` | `push` a `branches: [main]` (red de seguridad tras el merge) | `quality`, `integration`, `security`, `codeql`, `required` | Misma matriz que la PR, sin `pr-gate` (su `if: github.event_name == 'pull_request'` lo excluye; el `if:` de los heavy jobs acepta `skipped` de `pr-gate`). Issue #857, PR-D. |
| **Nightly** | `.github/workflows/ci.yml` | `schedule` (`cron: "0 3 * * *"`) | `quality`, `integration`, `security`, `codeql`, `required` | Sin rama; alimenta el snapshot del quality report. |
| **Manual** | `.github/workflows/ci.yml` | `workflow_dispatch` | `quality`, `integration`, `security`, `codeql`, `required` | El mismo set de jobs que el push a `main`; útil para reproducir el veredicto sin abrir una PR. |
| **Release** | `.github/workflows/release.yml` | `push tags: "v*"` | `preflight`, `mutation`, `e2e`, `publish`, `verify` | `verify` cierra la ratchet de provenance. |

`schedule`, `push` a `main` y `workflow_dispatch` reusan los mismos jobs del PR (excepto `pr-gate`, que se salta por su `if: github.event_name == 'pull_request'`). El `if:` de cada heavy job lista explícitamente los cuatro eventos y acepta `needs.pr-gate.result == 'skipped'` para no abortar la red de seguridad cuando `pr-gate` no corre.

## Nivel PR (`.github/workflows/ci.yml`)

Todos los jobs comparten `concurrency.group: ${{ github.workflow }}-${{ github.ref }}` con `cancel-in-progress: true` para no apilar corridas obsoletas. Los jobs no son independientes: `quality`, `integration`, `security` y `codeql` declaran `needs: pr-gate`, y el agregador `required` necesita los cuatro.

| Job | Runs-on | Steps principales | Comando / coste |
|---|---|---|---|
| `pr-gate` | `ubuntu-24.04` | `pr size`, `issue spec`, `walkthrough schema` | Llama a `scripts/check_pr_size.py`, `scripts/check_issue_spec.py` y `scripts/check_walkthrough_schema.py`. Solo corre si `event == pull_request`. |
| `quality` | `ubuntu-24.04` | `install`, `format scripts`, `format`, `lint`, `typecheck`, `workflow gate`, `suites populated`, `tests`, `quality gates and indicators` | <code>pytest -c app/pyproject.toml --rootdir=app -m "not integration"</code>; al final, `scripts/quality_report.py`. |
| `integration` | `ubuntu-24.04` (Postgres como service) | `install`, `Resolve the DSN from the assigned host port`, `Apply migrations`, `Run integration tests` | <code>pytest -c app/pyproject.toml --rootdir=app -m integration</code>. El DSN no viene de un plugin: lo construye el step `Resolve the DSN from the assigned host port`, que lee `POSTGRES_HOST_PORT=${{ job.services.postgres.ports['5432'] }}` y exporta `INTEGRATION_DSN` y `APAP_INTEGRATION_URL` (<code>postgresql+asyncpg://lanzadera:lanzadera-dev-only@127.0.0.1:${POSTGRES_HOST_PORT}/lanzadera</code>) a `$GITHUB_ENV`. `app/migrations/env.py` lee `DATABASE_URL`, que el step `Apply migrations` iguala a `APAP_INTEGRATION_URL` antes de `alembic upgrade head`. |
| `security` | — (workflow reutilizado, jobs `gitleaks`, `trivy-config`, `pip-audit`) | Escaneo estático y de dependencias | `gitleaks`, `trivy config`, `pip-audit`. |
| `codeql` | `ubuntu-24.04` | Análisis estático de GitHub | Lenguajes: python, javascript. |
| `required` | `ubuntu-24.04` | `Enforce the required-job policy` | Ejecuta `scripts/check_required_jobs.py` — es el agregador que la branch protection exige. |

### `pr-gate` (rápido, primer filtro)

Sólo corre en `pull_request`. Los tres gates son baratos y rechazan PRs que no deberían gastar minutos de runner:

- `pr size`: rechaza PRs de más de 400 líneas (`additions + deletions` = `MAX_CHANGED_LINES` en `scripts/check_pr_size.py`). La única exención es la etiqueta `size:exception` aplicada por el mantenedor (`SIZE_EXCEPTION_LABEL`, función `has_size_exception_label`); el cuerpo del PR no se lee. El modo local del mismo script no tiene exención: `El modo local no tiene excepción: reduzca el diff o parta el PR.`
- `issue spec`: verifica que el cuerpo del PR cita una issue (Refs/Closes) y, si es una epic de walkthrough, que el bloque de spec está bien formado.
- `walkthrough schema`: valida que cada <code>docs/03-aplicaciones/&lt;app&gt;/walkthrough-G*.json</code> cumple el esquema vigente.

### `quality` (suite principal)

Orden interno de los steps (ver `ci.yml` línea 188 a línea 360):

1. `install` — <code>pip install --require-hashes --requirement app/requirements-dev.lock</code>. Hard Rule 2: pins exactos en `app/pyproject.toml` y `app/requirements-dev.lock`.
2. `format scripts` — `ruff format` sobre `scripts/` (no sobre `app/`).
3. `format` — `ruff format --check` sobre `app/src`, `tests/`, `openspec/`. Si falla, el runner reporta el diff.
5. `lint` — `ruff check` con la `select` de `app/pyproject.toml`.
6. `typecheck` — `mypy --explicit-package-bases` con la config en `app/pyproject.toml` (Hard Rule 15: pins de mypy==1.13.0).
7. `workflow gate` — ejecuta `scripts/check_workflows.py` y `scripts/check_required_jobs.py` (parte de la matriz `required`).
8. `suites populated` — ejecuta `scripts/check_suites_populated.py`. Issue #882: una suite obligatoria vacía es una garantía falsa; cada suite (`unit`, `integration`, `e2e_smoke`) debe recoger al menos un test bajo `--collect-only -m <expr>`.
9. `tests` — <code>pytest -c app/pyproject.toml --rootdir=app -m "not integration"</code>. El flag `--rootdir=app` es load-bearing: sin él, `pytest` no encuentra la configuración y cada test falla con `import error`.
10. `quality gates and indicators` — ejecuta `scripts/quality_report.py`. Renderiza el quality report en el step summary del job.

Cobertura: el plugin `coverage` se retiró en #877; el `pytest-cov` ya no está en `app/pyproject.toml` y los flags `--cov` no existen. La ratchet shrink-only actual es la de mutación (release), no de cobertura.

### `integration` (Postgres + Alembic)

El job arranca Postgres como service y aplica el esquema con `alembic upgrade head` antes de correr los tests. La fixture de tests/migrations (`tests/lanzadera/migrations/test_seed_migrations.py`) verifica el seed. El comando es <code>pytest -c app/pyproject.toml --rootdir=app -m integration</code>. El driver es `asyncpg` (no `psycopg2`) y no se usa `pytest-postgresql`: el DSN lo calcula y exporta el step `Resolve the DSN from the assigned host port` como se describe arriba.

## Level nightly

El trigger `schedule` en `ci.yml` corre los mismos jobs `quality`, `integration`, `security`, `codeql` y `required`. La diferencia operacional es:

- Sin `pr-gate` (su `if: github.event_name == 'pull_request'` lo excluye): las verificaciones baratas se duplican en `quality`/`workflow gate` así que no hay regresión de cobertura.
- Snapshots de `scripts/quality_report.py` se acumulan para auditoría de tendencia (no se exponen en este repo; ver `docs/gates.md` §3).

No hay un workflow nightly separado; el schedule vive en `ci.yml` a propósito para mantener una sola fuente de verdad.

## Level release (`.github/workflows/release.yml`)

Sólo se dispara con push de tags (`v*`). El veredicto final es `cosign verify` sobre el digest publicado. Jobs:

| Job | Runs-on | Coste | Verifica |
|---|---|---|---|
| `preflight` | `ubuntu-24.04` | 15 min | Identidad del tag (`scripts/release-preflight.sh`) y CI verde en el SHA (`scripts/require-ci-success.sh`). Clasifica el canal como `rc` o `stable`. |
| `mutation` | `ubuntu-24.04-arm` | 150 min | Cosmic-ray sobre el dominio (`scripts/mutation_shards.py`, 10 tramos deterministas) y ratchet `scripts/check_mutation.py`. Issue #857, PR-C1d: la mutación salió de `ci.yml` (#881) y vive aquí porque un release es raro y la latencia es aceptable. |
| `e2e` | `ubuntu-24.04-arm` | 20 min | Pull de la imagen publicada en `ghcr.io`, arranque con `docker run`, suite Playwright (<code>e2e/*.spec.ts</code>) contra `http://localhost:8000`. |
| `publish` | `ubuntu-24.04-arm` | 45 min | Buildx multi-arch + push a <code>ghcr.io/&lt;owner&gt;/access2web-blueprint:&lt;tag&gt;</code> + firma keyless con `cosign sign`. Necesita `e2e` y `mutation` verdes. |
| `verify` | `ubuntu-24.04` | 15 min | `cosign verify` con `--certificate-identity` y `--certificate-oidc-issuer` anclados al workflow y a la identidad OIDC de GitHub Actions. Pull del digest verificado. |

ARM64 conservado: la imagen que `publish` produce se consume en `e2e`; si cambia la arquitectura, el smoke falla. Documentado en `docs/10-runners.md`.

## Suites y markers

`app/pyproject.toml` declara los markers en `[tool.pytest.ini_options].markers`. Con `--strict-markers` activo, una errata en un marker falla la recolección antes de las asserts.

| Marker | Significado | Quién lo recoge |
|---|---|---|
| `integration` | Toca Postgres u otra dependencia out-of-process. | `ci.yml::integration` (`-m integration`) |
| `e2e_smoke` | Gate-smoke tests que ejecutan scripts reales como subprocessos. | `ci.yml::quality` (parte de la suite unit; ver `scripts/check_suites_populated.py`) |
| `e2e` | Test end-to-end contra la imagen publicada (ASGI + DB + fixtures). | `release.yml::e2e` (Playwright) |
| `allow_uncaught_hosts` | Opt-out de `--disable-socket` para los hosts listados en `--allow-hosts` (plugin `pytest-socket`). | Unit suite (`tests/test_isolation.py` W88 #883) |
| (sin marker) | Unit suite por exclusión (`-m "not integration and not e2e_smoke"`). | `ci.yml::quality::tests` |

Las suites obligatorias — `unit`, `integration`, `e2e_smoke` — son las que `scripts/check_suites_populated.py` exige no vacías. Renombrar una suite rompe la guard antes de gastar el runner.

## Quality gates

Los quality gates son scripts <code>scripts/check_*.py</code> que el runner invoca como steps. El contrato manual está en `docs/gates.md`; el contrato de wiring está en `tests/test_ci_workflow.py`. Aquí solo la tabla de responsabilidad:

| Gate | Job que lo invoca | Protege |
|---|---|---|
| `scripts/check_pr_size.py` | `pr-gate` | PR ≤ 400 líneas. |
| `scripts/check_issue_spec.py` | `pr-gate` | Cuerpo de issue conforme. |
| `scripts/check_walkthrough_schema.py` | `pr-gate` | `walkthrough-*.json` válidos. |
| `scripts/check_workflows.py` | `quality` (`workflow gate`) | `.github/workflows/*.yml` bien formado. |
| `scripts/check_required_jobs.py` | `quality` (`workflow gate`) + `required` (enforce) | Matriz `required` completa. |
| `scripts/check_layers.py` | `quality` (paso de quality gates) | DA-1 hexagonalidad. |
| `scripts/check_legacy_hashes.py` | `quality` | DA-13 prohibición legacy. |
| `scripts/check_mutation_sites.py` | `quality` | DA-13 pin AST. |
| `scripts/check_complexity.py` | `quality` | Complejidad. |
| `scripts/check_dry.py` | `quality` | DRY. |
| `scripts/check_deadcode.py` | `quality` | Código muerto. |
| `scripts/check_decision_guards.py` | `quality` | Decisiones D-<n> con guardia. |
| `scripts/check_suites_populated.py` | `quality` | Suites obligatorias no vacías (#882). |
| `scripts/check_test_classification.py` | `quality` | Tests bien clasificados. |
| `scripts/check_legacy_retirement.py` | `quality` | Retirada de paths legacy. |
| `scripts/check_mutation.py` | `release.yml::mutation` | Ratchet shrink-only de mutación. |

## Cómo escribir un test

| Tipo de test | Carpeta | Marker | Comando local |
|---|---|---|---|
| Unit (caso de uso) | <code>tests/lanzadera/application/&lt;case&gt;/</code> | (sin marker) | <code>pytest -c app/pyproject.toml --rootdir=app tests/lanzadera/application/&lt;case&gt;/</code> |
| Unit (adapter) | `tests/lanzadera/adapters/` | (sin marker) | Igual, con la subcarpeta del adapter. |
| Unit (gate smoke) | `tests/lanzadera/scripts/` | `@pytest.mark.e2e_smoke` | <code>pytest -c app/pyproject.toml --rootdir=app -m e2e_smoke</code> |
| Integration | `tests/lanzadera/` (subcarpeta específica) | `@pytest.mark.integration` | Requiere Postgres en `localhost:5432` o DSN exportado. |
| E2E (Playwright) | <code>e2e/&lt;spec&gt;.spec.ts</code> | (no aplica, es TS) | <code>npx playwright test e2e/&lt;spec&gt;.spec.ts</code> (no corre en PR). |

La regla práctica: si el test toca Postgres, márquelo `integration`; si toca una dependencia de runtime externa (HTTP live, SMTP, S3), cree un fake o un adapter y déjelo en unit. Las pruebas que necesitan un servidor en vivo viven sólo en `e2e/` y se reservan al release.

## Pre-flight local

Antes de empujar una PR:


    ./scripts/local-preflight.sh


El script es la versión resumida de los steps baratos de `ci.yml::pr-gate` + `ci.yml::quality` (lint, typecheck, suites populated, unit). Si pasa, la PR no debería gastar minutos de runner en el primer filtro.

## Fronteras explícitas

- **No añada un job nuevo a `ci.yml` sin pasar por `scripts/check_workflows.py` y `scripts/check_required_jobs.py`.** Ambos gates validan la matriz y rompen la ratchet si se omite una actualización.
- **No mueva un step entre jobs sin revisar `tests/test_ci_workflow.py`.** El wiring está pinado por tests (`test_pull_request_types_cover_edit_and_label_events` y similares).
- **No cambie los markers sin actualizar `docs/gates.md` §2 y `scripts/check_suites_populated.py`.** Con `--strict-markers` activo, un marker nuevo no declarado falla la recolección.
- **No cubra un test que requiere servidor con la suite `unit`.** Use `e2e_smoke` (subproceso real) o `e2e` (release).

## Si necesitas X, lee/toca Y

| Si necesitas... | Empieza por... | Luego revisa... |
|---|---|---|
| Ejecutar los tests de un PR | La sección «Nivel PR» de esta página | `scripts/local-preflight.sh` |
| Añadir un gate | `scripts/check_required_jobs.py` como modelo | `.github/workflows/ci.yml` y `tests/test_ci_workflow.py` |
| Cambiar la suite E2E | `e2e/` | `.github/workflows/release.yml` |

## Regla de colocación

Un test vive junto a la capa que prueba, bajo `tests`, en la carpeta de su módulo y de su capa, y el job que lo ejecuta se declara en `.github/workflows/ci.yml`.

## Lista de comprobación del contribuidor

- [ ] Si el test toca Postgres, lleva `@pytest.mark.integration`.
- [ ] Si el test ejecuta un script real como subproceso, lleva `@pytest.mark.e2e_smoke`.
- [ ] Si el test abre un socket no bloqueado por `--disable-socket`, lleva `@pytest.mark.allow_uncaught_hosts` y el host está en `--allow-hosts`.
- [ ] `./scripts/local-preflight.sh` corre limpio antes del push.
- [ ] Si añadió un quality gate, actualizó `docs/gates.md` y `tests/test_ci_workflow.py` y esta tabla.

## Siguiente paso

Continúe con [Lo que no está aquí](what-is-not-here.md) para saber qué NO está en este repo y cómo no inventarlo.
