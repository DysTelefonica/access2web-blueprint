[← Volver al Codebase Guide](../CODEBASE-GUIDE.md)

# Testing y CI

> Esta página rastrea `main`, que puede contener cambios sin liberar. Para la última versión liberada, véase la etiqueta correspondiente.

Este repo separa el veredicto en **cinco niveles de disparo** sobre tres workflows — PR, post-merge, nightly, manual y release — y cada nivel tiene un contrato propio. La superficie común está en `.github/workflows/ci.yml` (PR pesada + post-merge + nightly + manual), `.github/workflows/pr-gate.yml` (el gate ligero de la PR) y `.github/workflows/release.yml` (release); el contrato manual de los quality gates vive en `docs/gates.md`.

## Mapa de niveles

| Nivel | Workflow | Disparador | Jobs | Veredicto |
|---|---|---|---|---|
| **PR (pesada)** | `.github/workflows/ci.yml` | `pull_request` con `types: [opened, synchronize, reopened]` (cualquier rama) | `pr-gate`, `quality`, `integration`, `security`, `codeql`, `e2e-smoke`, `required` | Branch protection exige `required` verde. `e2e-smoke` solo corre aquí (issue #876). Los eventos que solo cambian el cuerpo o las etiquetas (`edited`, `labeled`, `unlabeled`) ya no llegan aquí (issue #1368). |
| **PR (ligera)** | `.github/workflows/pr-gate.yml` | `pull_request` con los seis tipos | `pr-gate` | El gate de cuerpo, etiquetas y walkthrough. `ci.yml` lo invoca (`workflow_call`), así que `required` sigue leyendo su veredicto. |
| **Post-merge** | `.github/workflows/ci.yml` | `push` a `branches: [main]` (red de seguridad tras el merge) | `quality`, `integration`, `security`, `codeql`, `required` | Misma matriz que la PR, sin `pr-gate` (su `if: github.event_name == 'pull_request'` lo excluye; el `if:` de los heavy jobs acepta `skipped` de `pr-gate`). Issue #857, PR-D. |
| **Nightly** | `.github/workflows/ci.yml` | `schedule` (`cron: "0 3 * * *"`) | `quality`, `integration`, `security`, `codeql`, `required` | Sin rama; alimenta el snapshot del quality report. |
| **Manual** | `.github/workflows/ci.yml` | `workflow_dispatch` | `quality`, `integration`, `security`, `codeql`, `required` | El mismo set de jobs que el push a `main`; útil para reproducir el veredicto sin abrir una PR. |
| **Release** | `.github/workflows/release.yml` | `push tags: "v*"` | `preflight`, `mutation`, `e2e`, `publish`, `verify` | `verify` cierra la ratchet de provenance. |

`schedule`, `push` a `main` y `workflow_dispatch` reusan los mismos jobs del PR (excepto `pr-gate`, que se salta por su `if: github.event_name == 'pull_request'`). El `if:` de cada heavy job lista explícitamente los cuatro eventos y acepta `needs.pr-gate.result == 'skipped'` para no abortar la red de seguridad cuando `pr-gate` no corre.

### El carril docs-only (issue #1368)

Un pull request cuyo diff cae entero en `docs/` no necesita la suite pesada. El `pr-gate` publica el marcador `docs_changed` (`scripts/check_docs_only_lane.py`, fail-closed: un diff vacío o con un solo path fuera de `docs/` responde `false`), y los jobs pesados (`quality`, `integration`, `security`, `codeql`, `e2e-smoke`) se saltan cuando el marcador vale exactamente `'true'`. El agregador `required` acepta esos saltos por diseño: los declara el `docs_only_lane` de `.github/required-jobs-policy.json`. `docs-links` sigue corriendo en ese carril, porque es el gate que caza los enlaces rotos del PRD.

## Nivel PR (`.github/workflows/ci.yml`)

Todos los jobs comparten `concurrency.group: ${{ github.workflow }}-${{ github.ref }}`, que solo cancela en `synchronize` (un commit nuevo deja obsoleto el run anterior; los eventos que mantienen el SHA no pueden cancelarlo, o quedarían un check cancelado que bloquea el merge). El `pr-gate` tiene además su propio grupo FIFO por PR (`gate-pr-gate-${{ github.workflow }}-…`, `cancel-in-progress: false`), porque su gate no puede saltarse a sí mismo: el job corre en dos instancias —su propio disparador `pull_request` y la llamada de `ci.yml`— y el grupo nombra la instancia (`github.workflow` vale el workflow *caller* dentro de un workflow reutilizable); si compartieran cola, la instancia pendiente se cancelaría y `required` leería la cancelación como fallo (#1372). Los jobs no son independientes: `quality`, `integration`, `security` y `codeql` declaran `needs: pr-gate`, y el agregador `required` necesita los cuatro.

| Job | Runs-on | Steps principales | Comando / coste |
|---|---|---|---|
| `pr-gate` | `ubuntu-24.04` | `pr size`, `issue spec`, `walkthrough schema`, `classify the diff (docs-only lane)` | Job de `.github/workflows/pr-gate.yml`, invocado desde aquí con el `uses:` de `./.github/workflows/pr-gate.yml` (issue #1368). Llama al `check_pr_size.py` del canon (`.agents/skills/ci-pattern/assets/pr-size/`), el `check_issue_spec.py` del canon (`.agents/skills/ci-pattern/assets/issue-spec/`) y `scripts/check_walkthrough_schema.py` y `scripts/check_docs_only_lane.py` (marcador `docs_changed`). Solo corre si `event == pull_request`. |
| `quality` | `ubuntu-24.04` | `install`, `format scripts`, `format`, `lint`, `typecheck`, `workflow gate`, `suites populated`, `tests`, `quality gates and indicators` | la suite sin el marker `integration` (comando 1 de «Comandos de referencia»); al final, `scripts/quality_report.py`. |
| `integration` | `ubuntu-24.04` (Postgres como service) | `install`, `Resolve the DSN from the assigned host port`, `Apply migrations`, `Run integration tests` | la suite con el marker `integration` (comando 2 de «Comandos de referencia»). El DSN no viene de un plugin: lo construye el step `Resolve the DSN from the assigned host port`, que lee `POSTGRES_HOST_PORT=${{ job.services.postgres.ports['5432'] }}` y exporta `INTEGRATION_DSN` y `APAP_INTEGRATION_URL` (`postgresql+asyncpg` sobre `127.0.0.1` y el puerto asignado) a `$GITHUB_ENV`. `app/migrations/env.py` lee `DATABASE_URL`, que el step `Apply migrations` iguala a `APAP_INTEGRATION_URL` antes de `alembic upgrade head`. |
| `e2e-smoke` | `ubuntu-24.04` (Postgres como service) | `Build the PR image`, `Resolve the DSN from the assigned host port`, `Install the lockfile and apply the real migrations`, `Start the app and wait for the health check`, `Run the E2E smoke specs` | Solo `pull_request` (issue #876). Construye la imagen de la PR, espera a `/health` hasta 60 s y ejecuta `smoke.spec.ts` y `auth.spec.ts` de Playwright; tope de 5 min. En push, schedule y workflow_dispatch se salta y `.github/required-jobs-policy.json` acepta ese salto; en `pull_request` un salto deja `required` en rojo. La suite completa sigue en `release.yml` (job `e2e`). |
| `security` | — (workflow reutilizado, jobs `gitleaks`, `trivy-config`, `pip-audit`) | Escaneo estático y de dependencias | `gitleaks`, `trivy config`, `pip-audit`. |
| `codeql` | `ubuntu-24.04` | Análisis estático de GitHub | Lenguajes: python, javascript. |
| `required` | `ubuntu-24.04` | `Enforce the required-job policy` | Ejecuta el `check_required_jobs.py` del canon (`.agents/skills/ci-pattern/assets/required-jobs/`) contra `.github/required-jobs-policy.json` — es el agregador que la branch protection exige. |

### `pr-gate` (rápido, primer filtro)

Vive en `.github/workflows/pr-gate.yml` (issue #1368) y `ci.yml` lo invoca con el `uses:` de `./.github/workflows/pr-gate.yml`, de modo que el agregador `required` sigue leyendo su veredicto y su marcador `docs_changed`: `needs:` no cruza ficheros de workflow. Sólo corre en `pull_request`, en los seis tipos de evento: los tres que pueden cambiar el código y los tres que cambian el cuerpo o las etiquetas, que es lo único que hay que reevaluar cuando el cuerpo o las etiquetas cambian. Los tres gates son baratos y rechazan PRs que no deberían gastar minutos de runner:

- `pr size`: rechaza PRs de más de 400 líneas (`additions + deletions` = `MAX_CHANGED_LINES` en `.agents/skills/ci-pattern/assets/pr-size/check_pr_size.py`, el gate del canon). La única exención es la etiqueta `size:exception` aplicada por el mantenedor (`OVERRIDE_LABEL`); el cuerpo del PR no se lee. En local, sin evento, el gate mide con `git diff` contra la base y la etiqueta se pasa con `--labels` o `$PR_LABELS`.
- `issue spec`: verifica que el cuerpo del PR cita una issue (Refs/Closes) y, si es una epic de walkthrough, que el bloque de spec está bien formado.
- `walkthrough schema`: valida que cada `docs/03-aplicaciones/<app>/walkthrough-G*.json` cumple el esquema vigente.
- `classify the diff (docs-only lane)`: calcula el diff contra el merge-base con la rama base y publica el marcador `docs_changed` (`true` solo si todo el diff cae bajo `docs/`). Es un clasificador, no un gate: no falla nunca; su salida decide qué jobs pesados corren.

### `quality` (suite principal)

Orden interno de los steps (ver `ci.yml` línea 188 a línea 360):

1. `install` — instalación con hashes exigidos desde `app/requirements-dev.lock`. Hard Rule 2: pins exactos en `app/pyproject.toml` y `app/requirements-dev.lock`.
2. `format scripts` — `ruff format` sobre `scripts/` (no sobre `app/`).
3. `format` — `ruff format --check` sobre `app/src`, `tests/`, `openspec/`. Si falla, el runner reporta el diff.
5. `lint` — `ruff check` con la `select` de `app/pyproject.toml`.
6. `typecheck` — `mypy --explicit-package-bases` con la config en `app/pyproject.toml` (Hard Rule 15: pins de mypy==1.13.0).
7. `workflow gate` — ejecuta `scripts/check_workflows.py`; la matriz `required` la comprueba el job `required` con el gate del canon.
8. `suites populated` — ejecuta `scripts/check_suites_populated.py`. Issue #882: una suite obligatoria vacía es una garantía falsa; cada suite (`unit`, `integration`, `e2e_smoke`) debe recoger al menos un test bajo `--collect-only -m <expr>`.
9. `tests` — la suite sin el marker `integration` (comando 1 de «Comandos de referencia»). El flag `--rootdir=app` es load-bearing: sin él, `pytest` no encuentra la configuración y cada test falla con `import error`.
10. `quality gates and indicators` — ejecuta `scripts/quality_report.py`. Renderiza el quality report en el step summary del job.

Cobertura: el plugin `coverage` se retiró en #877; el `pytest-cov` ya no está en `app/pyproject.toml` y los flags `--cov` no existen. La ratchet shrink-only actual es la de mutación (release), no de cobertura.

### `integration` (Postgres + Alembic)

El job arranca Postgres como service y aplica el esquema con `alembic upgrade head` antes de correr los tests. La fixture de tests/migrations (`tests/lanzadera/migrations/test_seed_migrations.py`) verifica el seed. El comando es la suite con el marker `integration` (comando 2 de «Comandos de referencia»). El driver es `asyncpg` (no `psycopg2`) y no se usa `pytest-postgresql`: el DSN lo calcula y exporta el step `Resolve the DSN from the assigned host port` como se describe arriba.

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
| `e2e` | `ubuntu-24.04-arm` | 20 min | Pull de la imagen publicada en `ghcr.io`, arranque con `docker run`, suite Playwright (`e2e/*.spec.ts`) contra `http://localhost:8000`. |
| `publish` | `ubuntu-24.04-arm` | 45 min | Buildx multi-arch + push a `ghcr.io` (imagen del repositorio, con la etiqueta de la release) + firma keyless con `cosign sign`. Necesita `e2e` y `mutation` verdes. |
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

Los quality gates son scripts `scripts/check_*.py` que el runner invoca como steps. El contrato manual está en `docs/gates.md`; el contrato de wiring está en `tests/test_ci_workflow.py`. Aquí solo la tabla de responsabilidad:

| Gate | Job que lo invoca | Protege |
|---|---|---|
| `.agents/skills/ci-pattern/assets/pr-size/check_pr_size.py` | `pr-gate` | PR ≤ 400 líneas. |
| `.agents/skills/ci-pattern/assets/issue-spec/check_issue_spec.py` (canon) | `pr-gate` | Cuerpo de issue conforme. |
| `scripts/check_walkthrough_schema.py` | `pr-gate` | `walkthrough-*.json` válidos. |
| `scripts/check_workflows.py` | `quality` (`workflow gate`) | `.github/workflows/*.yml` bien formado. |
| `.agents/skills/ci-pattern/assets/required-jobs/check_required_jobs.py` (canon) | `required` (enforce) | Matriz `required` completa. |
| `.importlinter` | `docs-architecture` (job requerido) | DA-1 hexagonalidad: derivado del bloque `architecture` de `docs/architecture.md`. |
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
| Unit (caso de uso) | `tests/lanzadera/application/<case>/` | (sin marker) | la suite de un caso de uso (comando 3 de «Comandos de referencia») |
| Unit (adapter) | `tests/lanzadera/adapters/` | (sin marker) | Igual, con la subcarpeta del adapter. |
| Unit (gate smoke) | `tests/lanzadera/scripts/` | `@pytest.mark.e2e_smoke` | la suite con el marker `e2e_smoke` (comando 4 de «Comandos de referencia») |
| Integration | `tests/lanzadera/` (subcarpeta específica) | `@pytest.mark.integration` | Requiere Postgres en `localhost:5432` o DSN exportado. |
| E2E (Playwright) | `e2e/<spec>.spec.ts` | (no aplica, es TS) | Playwright sobre una spec (comando 5 de «Comandos de referencia») (no corre en PR). |

La regla práctica: si el test toca Postgres, márquelo `integration`; si toca una dependencia de runtime externa (HTTP live, SMTP, S3), cree un fake o un adapter y déjelo en unit. Las pruebas que necesitan un servidor en vivo viven sólo en `e2e/` y se reservan al release.

## Comandos de referencia

Se ejecutan desde la raíz del repositorio. Los números son los que citan las secciones de arriba.

```
# 1. Suite sin el marker integration (job quality)
pytest -c app/pyproject.toml --rootdir=app -m "not integration"
# 2. Suite con el marker integration (job integration, con Postgres)
pytest -c app/pyproject.toml --rootdir=app -m integration
# 3. Suite de un caso de uso
pytest -c app/pyproject.toml --rootdir=app tests/lanzadera/application/<case>/
# 4. Suite con el marker e2e_smoke
pytest -c app/pyproject.toml --rootdir=app -m e2e_smoke
# 5. Playwright sobre una spec (no corre en PR)
npx playwright test e2e/<spec>.spec.ts
# Instalación con hashes exigidos
pip install --require-hashes --requirement app/requirements-dev.lock
```

## Pre-flight local

Antes de empujar una PR:

```bash
./scripts/local-preflight.sh
```

El script es la versión resumida de los steps baratos de `ci.yml::pr-gate` + `ci.yml::quality` (lint, typecheck, suites populated, unit). Si pasa, la PR no debería gastar minutos de runner en el primer filtro.

## Fronteras explícitas

- **No añada un job nuevo a `ci.yml` sin pasar por `scripts/check_workflows.py` y por la política `.github/required-jobs-policy.json`.** Ambos gates validan la matriz y rompen la ratchet si se omite una actualización.
- **No mueva un step entre jobs sin revisar `tests/test_ci_workflow.py`.** El wiring está pinado por tests (`test_pr_gate_listens_to_edit_and_label_events` y similares). Desde #1375 los pins de la FORMA de `ci.yml` —su lista de eventos, los `if:` exactos de los jobs pesados, sus `needs:`— no viven en ese fichero: los verifica en cada run el gate estructural del canon (`check_workflow_policy.py`, `check_parity.py`) junto con la política `.github/required-jobs-policy.json`. Lo que queda en los tests es lo que ningún otro gate cubre: qué ejecuta cada gate, cómo se blinda y que existe y corre.
- **No cambie los markers sin actualizar `docs/gates.md` §2 y `scripts/check_suites_populated.py`.** Con `--strict-markers` activo, un marker nuevo no declarado falla la recolección.
- **No cubra un test que requiere servidor con la suite `unit`.** Use `e2e_smoke` (subproceso real) o `e2e` (release).

## Si necesitas X, lee/toca Y

| Si necesitas... | Empieza por... | Luego revisa... |
|---|---|---|
| Ejecutar los tests de un PR | La sección «Nivel PR» de esta página | `scripts/local-preflight.sh` |
| Añadir un gate | `.github/required-jobs-policy.json` y su `needs` en el job `required` | `.github/workflows/ci.yml` y `tests/test_ci_workflow.py` |
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