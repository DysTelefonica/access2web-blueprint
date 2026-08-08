# Calidad de código y CI en access2web-blueprint

> **Sentence que organiza este documento**: la calidad de la plataforma no se discute, se verifica. Cada decisión arquitectónica aprobada se traduce en un gate que falla cuando el código la viola. Ver `docs/09-arquitectura-objetivo-y-principios.md` para el contrato de origen.

## Qué es y qué no es

| Es | No es |
|---|---|
| El contrato operativo de los quality gates del MVP de Lanzadera y de los módulos que le sigan. | Una guía de instalación ni de estilo de código. |
| La traducción de la decisión D8 (hexagonal global) a un gate ejecutable en CI. | Una transcripción del pipeline de APAP_WEB; cada proyecto adapta a su realidad. |
| El plan de implementación por día, con archivos concretos y coste estimado. | Un backlog de tickets; los tickets derivados viven en `docs/backlog-deuda.md` y en GitHub. |
| Una bitácora de las trampas conocidas, con su contramedida. | Una auditoría del estado actual; aquí solo se documenta el objetivo. |
| La referencia que cualquier IA o mantenedor consulta antes de proponer un PR de plataforma. | Un sustituto del CODEBASE-GUIDE; este doc entra como capítulo del mismo. |

## Decisiones adoptadas y su origen

| # | Decisión | Origen | Estado |
|---|---|---|---|
| QC-1 | Adoptar el set de quality gates del MVP de Lanzadera derivado de APAP_WEB, con exclusiones explícitas y justificadas. | Sesión de usuario del 2026-08-08; reporte de exploración sobre APAP_WEB issues #380, #381, #393, #424, #428, #436, #437, #441, #442, #443. | APROBADO |
| QC-2 | Hexagonal layer gate obligatorio desde el primer slice que se escriba. | D8 (`architecture/global-hexagonal-principle`). El gate verifica, no solo declara. | APROBADO |
| QC-3 | `ruff` con `select` base `E,F,W,I,UP,B` desde día uno, sin ratchet extendido. | APAP_WEB #380. El ratchet solo aporta valor en codebases con deuda heredada. | APROBADO |
| QC-4 | `mypy` strict en `app/` con `enable_error_code = ["ignore-without-code"]`. | APAP_WEB #201, #386. El estilo `Annotated[T, Depends(get_x)]` se fija antes de acumular 291 violations. | APROBADO |
| QC-5 | Cobertura mínima `85%` global y `100%` en `CRITICAL_HELPERS`. | APAP_WEB #199, #257, #331. Los `CRITICAL_HELPERS` del MVP son `hash_password` y `verify_password`. | APROBADO |
| QC-6 | PR size gate `400` líneas y branch-name gate `^(feat\|fix\|refactor\|docs\|ci\|test)/<n>-<slug>$`. | APAP_WEB #441, #442. Las PRs grandes agruparon fixes que luego se revertían juntas. | APROBADO |
| QC-7 | Security scanning con `pip-audit`, `gitleaks` y `trivy`, en venv throwaway y con imágenes pinned por digest. | APAP_WEB #381, #338, #393. El job «security-deep» corre semanal. | APROBADO |
| QC-8 | Mutación con cosmic-ray queda diferida a Fase 2; no se porta al MVP. | APAP_WEB #431, #434. Sin código que mutar ni suite madura, el gate sería ruido. | APROBADO (diferido) |
| QC-9 | `check_layers.py` se adapta como copia literal del script de APAP_WEB, ajustando `ALLOWED_IMPORTS`, `PURE_LAYERS` y `classify_layer`. | APAP_WEB #436, #437. El script es stdlib + ast; portable sin dependencias. | APROBADO |
| QC-11 | Se adoptan CRAP (`≤ 6`) y DRY (0 bloques duplicados) desde el MVP, más una capa de indicadores: cada gate emite un envelope machine-readable y `quality_report.py` publica `quality-report.json`. El orden de ejecución vive en código (`quality_report.GATES`), no en YAML. | Auditoría contra `unclebob/swarm-forge` (2026-08-08); `cleaner.prompt` fija `CRAP ≤ 6` y ordena CRAP antes que DRY. La cobertura sola no prueba que los tests aserten; CRAP pone precio a la complejidad no cubierta. Ojo: `CRAP ≤ 6` domina a `CC ≤ 15` — con cobertura total, ninguna función por encima de complejidad 6 pasa. | APROBADO |
| QC-10 | El gate de complejidad usa techo absoluto y global, nunca `top-N`. Todo `BASELINE` declara valor objetivo y fecha de caducidad, y el gate falla cuando esa fecha pasa sin haber alcanzado el objetivo. | Auditoría de la skill `deterministic-quality-harness` contra `unclebob/swarm-forge` (2026-08-08). Bajo `top-N` el veredicto sobre una función depende de la complejidad de otras funciones no relacionadas, así que el mismo código cambia de resultado cuando cambian sus vecinas. Sin fecha de caducidad, un ratchet nunca termina. | APROBADO |

## Quality gates del MVP de Lanzadera

El MVP activa los siguientes gates. Cada uno aparece en `ci.yml` con su step y su exit code explícito.

| Gate | Mecanismo | Frecuencia | Coste de adopción |
|---|---|---|---|
| Lint base (`ruff check .`) | `pyproject.toml` `[tool.ruff]` con `select = ["E","F","W","I","UP","B"]`, pin exacto `ruff==0.15.21`. | Por PR. | 30 min. |
| Typecheck (`mypy app/`) | `pyproject.toml` `[tool.mypy]` con `python_version = "3.12"`, `enable_error_code = ["ignore-without-code"]`. | Por PR. | 1 h. |
| Tests + cobertura (`pytest --cov-fail-under=85`) | `pyproject.toml` `[tool.coverage]` + plugin `scripts/pytest_plugin/coverage_gate.py`. | Por PR. | 1 h. |
| Hexagonal layer gate (`python scripts/check_layers.py`) | AST walk sobre `app/`; `ALLOWED_IMPORTS` por layer; `PURE_LAYERS` puro. | Por PR. | 1 día (script completo + tests). |
| Complexity ceiling (`python scripts/check_complexity.py`) | AST + cálculo de CC; techo absoluto y global `CC ≤ 15`, aplicado a todas las funciones sin excepción. | Por PR. | 3 h. |
| CRAP ceiling (`python scripts/check_crap.py`) | `CC² · (1 − cobertura)³ + CC` por función, techo `6`. Consume `coverage.json`; sin datos de cobertura falla cerrado. | Por PR. | 4 h. |
| DRY detector (`python scripts/check_dry.py`) | Clones type-1/type-2 sobre AST normalizado (sha256, nunca `hash()`); falla ante cualquier bloque de 5+ sentencias duplicado. | Por PR. | 4 h. |
| Indicadores (`python scripts/quality_report.py`) | Ejecuta layers → complexity → CRAP → DRY en orden fijo, agrega los envelopes y publica `quality-report.json` + tabla en el job summary. | Por PR. | 3 h. |
| PR size gate (`python scripts/check_pr_size.py`) | `git diff --stat` contra base; falla si `> 400`. Override `size:exception` exige justificación. | Por PR. | 30 min. |
| Branch-name gate (`python scripts/check_branch_name.py`) | Regex `^(feat\|fix\|refactor\|docs\|ci\|test)/<n>-<slug>$`. Allowlist `main`. | Por PR. | 30 min. |
| Secret scan (`gitleaks dir`) | Contenedor pinned por digest, `redact`, `exit-code 1`. | Por PR. | 1 h. |
| Dependency scan (`pip-audit`) | Venv throwaway, `pip install -e .[dev]`, `pip-audit --skip-editable`. | Por PR. | 1 h. |
| Dockerfile scan (`trivy config`) | Contenedor pinned por digest, severidad `HIGH,CRITICAL`. | Por PR. | 30 min. |
| Secret scan profundo (`gitleaks detect --source=.`) | `fetch-depth: 0` sobre todo el historial. | Semanal + manual. | 1 h. |
| Image scan (`trivy image`) | Para cada digest declarado en `Dockerfile`. | Semanal + tag push. | 1 h. |

## Hexagonal layer gate: el contrato arquitectónico verificable

La decisión D8 declara hexagonal global. Sin gate, la declaración queda en prosa y se viola sin huella. El script `scripts/check_layers.py` traduce la declaración a tres ejes verificables.

### Tres ejes que el gate enforce

| Eje | Qué detecta | Granularidad |
|---|---|---|
| Dirección de dependencias | Cada layer declara, vía `ALLOWED_IMPORTS`, qué otros layers puede importar. | Por archivo. |
| Vertical slicing | Cada módulo (`app/lanzadera/`, `app/expedientes/`, ...) posee su columna. `_check_slice` permite core cross-cutting exento y prohíbe reach-across entre módulos. | Por archivo. |
| Pureza interna | `domain`, `ports`, `application` no pueden importar frameworks (`fastapi`, `sqlalchemy`, `alembic`, `jinja2`, `httpx`, `asyncpg`). | Por archivo. |

### Layers y permisos de import

| Layer | Puede importar | Prohibido |
|---|---|---|
| `domain` | `domain` | Todo lo demás. |
| `ports` | `domain`, `ports` | Frameworks, adapters. |
| `application` | `domain`, `ports`, `application` | `adapters`, frameworks. |
| `adapters` | `domain`, `ports`, `adapters`, `shared` | `application`, frameworks cruzados. |
| `shared` | `domain`, `ports`, `shared` | Frameworks en purity-sensitive scopes. |
| `delivery` | Cualquier layer. | — |
| `di` | Cualquier layer (composition root). | — |

### Mecánica del script

1. AST walk sobre `app/`. Resuelve relativos y absolutos a archivo real, no a símbolos.
2. Clasifica cada archivo en `(módulo, layer)` con `classify_layer(path)`.
3. Compara imports contra `ALLOWED_IMPORTS[(módulo, layer)]`.
4. Verifica purity en `PURE_LAYERS`.
5. Compara contra `BASELINE` (shrink-only). BASELINE vacío al inicio: el MVP arranca limpio.
6. Exit `0` cuando clean; exit `1` ante cualquier violation nueva.

### Pin del wiring

`tests/test_layers.py::test_ci_workflow_lint_job_runs_layers_gate` verifica que `ci.yml` contiene el step `python scripts/check_layers.py`. Si alguien borra el step, el test falla. El guard vive en el workflow y en el test que lo pinea; no solo en el script.

## Configuración reproducible de ruff

### Bloque `[tool.ruff]` mínimo

```toml
[tool.ruff]
line-length = 100
target-version = "py312"
extend-exclude = [".git", ".venv", "venv", "build", "dist", ".pytest_cache", ".ruff_cache", "__pycache__"]

[tool.ruff.lint]
select = ["E", "F", "W", "I", "UP", "B"]
ignore = [
    "E501",  # identifiers no ASCII (dominio en castellano)
]

[tool.ruff.lint.per-file-ignores]
"tests/**" = ["B011"]

[tool.ruff.lint.isort]
known-first-party = ["app", "tests"]
```

### Reglas del setup

1. Pin exacto de versión (`ruff==0.15.21`). Drift de versión puede graduar o retirar rules sin aviso.
2. Select base, no extendido. El extendido (`S`, `ERA`, `ARG`, `FAST`, `PLR`, ...) se evalúa cuando la deuda aparezca.
3. Ignore por archivo, no global. Cada ignore es un smell a revisar.
4. Tests excluidos solo donde el idiom lo justifique (`B011` en `tests/**`).

## Security scanning

| Scanner | Invocación | Frecuencia | Configuración |
|---|---|---|---|
| `pip-audit` | `python -m venv .venv-audit && pip install -e .[dev] && pip-audit --skip-editable`. | Por PR. | `--ignore-vuln` arranca vacío; cada entry corresponde a un CVE documentado. |
| `gitleaks dir` | `docker run zricethezav/gitleaks@sha256:<digest> dir . --redact --no-banner --exit-code 1`. | Por PR. | `.gitleaksignore` con fingerprint + dated reason. |
| `trivy config` | `docker run aquasec/trivy@sha256:<digest> config /Dockerfile --severity HIGH,CRITICAL`. | Por PR. | — |
| `gitleaks detect` | `detect --source=.` con `fetch-depth: 0`. | Semanal + manual. | Mismo allowlist. |
| `trivy image` | Para cada digest del `Dockerfile`. | Semanal + tag push. | — |

### Reglas críticas

1. Pin de imágenes de los scanners por digest, no por tag.
2. Venv throwaway para `pip-audit`. El ambient del runner arrastra dependencias no relacionadas.
3. Sin `continue-on-error: true` en ningún step. Un gate que no falla no es un gate.
4. `pip-audit --ignore-vuln` shrink-only por entry. Quitar un ID equivale a cerrar el CVE.

## Plan de implementación por día

| Día | Horas | Tarea | Archivos creados |
|---|---|---|---|
| 0 | 0,5 | Estructura inicial de carpetas (`app/`, `tests/`, `scripts/`, `migrations/`). | `app/__init__.py`, `tests/__init__.py`, `pyproject.toml`, `Dockerfile`, `docker-compose.yml`, `.python-version`, `.dockerignore`. |
| 1 | 4 | Lint base, typecheck, cobertura, plugin coverage_gate, PR size, branch name, build. | `pyproject.toml`, `scripts/pytest_plugin/coverage_gate.py`, `scripts/check_pr_size.py`, `scripts/check_branch_name.py`, `.github/workflows/ci.yml`, `.github/workflows/pr-size.yml`, `Makefile`, `tests/test_ci_workflow.py`. |
| 2 | 3 | Security gates en PR y en deep semanal. | `.github/workflows/security.yml`, `.github/workflows/security-deep.yml`, `.gitleaksignore`. |
| 3 | 8 | Hexagonal layer gate + tests del wiring + `docs/quality-gates.md` (este doc). | `scripts/check_layers.py`, `tests/test_layers.py`, sección en `ci.yml`, ajustes de `pyproject.toml` y `Makefile`. |
| 4 | 3 | Complexity ratchet + tests. | `scripts/check_complexity.py`, `tests/test_complexity.py`, step en `ci.yml`. |
| 5 | 1 | Convenciones operativas en `AGENTS.md` raíz. | `AGENTS.md` (§Conf-1 a §Conf-8). |
| 6 | TDD | Primer módulo: `app/lanzadera/auth/`. Tests antes de código. CRITICAL_HELPERS registrados. | `app/lanzadera/...`, `tests/lanzadera/...`. |

Total estimado: 6,5 días con un mantenedor familiar con FastAPI, Docker multi-stage y GitHub Actions. Sin esa familiaridad, añadir dos días.

## Anti-patrones documentados (de APAP_WEB)

Cada fila resume un fallo real y la contramedida que se aplica en este repo.

| Issue APAP_WEB | Patrón de fallo | Contramedida adoptada |
|---|---|---|
| #393 | `grep ... \|\| true` mata un scanner que se ve «running». | Nunca `\|\| true` en un step. Si un scanner no aporta, se quita. |
| #424 | Regex de `vulture` asumía formato incorrecto; el gate nunca fallaba. | Todo scraper se valida con golden output antes de promotion. |
| #436 | `check_layers.py` existía solo en local, no commiteado. | Regla con gate debe estar en CI, commiteada, wired y pinned por test. |
| #437 | Regla demasiado estricta rompe su propio test suite. | Core cross-cutting exento por design; vertical slices aisladas entre sí. |
| #281 | Deploy job con `if` sobre merge-commit nunca corre. | Cualquier `if:` en workflow lleva test que pinea el predicado. |
| #282 | Variable `APAP_E2E_BASE_URL` sobrecargada (HTTP y DSN). | Un nombre, una cosa. Variables separadas. |
| #257 | `config.exitstatus = 1` en `pytest_terminal_summary` no cambia exit code. | Coverage gate via `pytest_sessionfinish` con `session.exitstatus`. Test con `pytester`. |
| #380 | `ruff>=0.6` deja entrar drift de versión que gradúa rules. | Pin exacto de `ruff` en `pyproject.toml`. |
| #442 | PR size gate mide cambios semánticos, no raw line endings. | `git diff --stat` con normalización CRLF→LF antes de contar. |
| #329, #332, #337, #339 | Fixes de auditoría 2026-07-30 se revertieron en masa el 2026-07-31. | Cada gate se shippea con su ratchet en el mismo PR; sin guard, el revert pasa silencioso. |
| #443 | Lazy-import con marker `# lazy-import:` documenta el baseline, no autoriza uno nuevo. | Marker requiere justificación y contador enlazado. |
| #440 | Ramas sin nombre disciplinado son indistinguibles dos años después. | Branch-name gate desde día uno. |
| #381 | `pip-audit` en ambient runner arrastra CVEs no del proyecto. | Venv throwaway scoped a `pyproject.toml`. |
| #201 | `# type: ignore` sin código gradúa silencio. | `enable_error_code = ["ignore-without-code"]`. |
| #203 | `app/main.py` 699/700 líneas, una línea de headroom. | Techo absoluto global `CC ≤ 15` (QC-10). El `top-N` queda prohibido: mide el vecindario, no la función. |
| #333 | `tests/` en wheel ships test fixtures a producción. | `hatch.build.targets.wheel.only-include = ["app"]`. |
| #338 | `python:3.11-slim` mutable tag. | Pin digest en cada `FROM`, dev y prod. |
| #389 | Adoptar rules con ruido conocido (`TRY003` 175, `PLR2004` 39). | Policy decision antes de activar rule. |
| #386 | `B008` ignore protege un antipatrón (`Depends()` en signature). | No ignorar `B008`; promover `Annotated[T, Depends(get_x)]`. |
| #210 | `make check-rules app` silencia detectors por scope relativo. | Cualquier linter custom se invoca con `.` (raíz). |
| #356 | Doble Postgres service container agota budget del runner. | Un service container por job. Self-hosted cuando se pase de 2000 min/mes. |

## Lo que NO se porta al MVP y por qué

| Gate o práctica | Razón de exclusión |
|---|---|
| Mutación con cosmic-ray (#431, #434). | Sin código que mutar ni suite madura; el gate sería ruido. Se reactiva cuando haya ≥3 módulos con cobertura > 70% y ≥5 tests por path. |
| Module size budget (`check_module_size.py`, 700L). | El MVP no genera módulos de 700L. El ratchet aplica a proyectos con deuda. |
| Route size budget (`check_route_size.py`, 50L). | El primer route handler será < 30L. El budget se vuelve irrelevante sin tráfico. |
| Docstring coverage ratchet (`check_docstring_coverage.py`, 73%). | El baseline inicial es 100% o N/A. Forzar floor sin docstrings reales es pulso sin señal. |
| Complexity gate por `top-N` de funciones. | El veredicto sobre una función dependería de la complejidad de otras funciones no relacionadas: el mismo código pasa o falla según sus vecinas. Sustituido por el techo absoluto global de QC-10. |
| JSCPD detector (`check_jscpd.py`, 1.78%). | Sustituido por `check_dry.py` (QC-11), que detecta clones type-2 sobre AST en vez de sobre texto y no necesita dependencia externa. |
| Mutation-sites cap (`check_mutation_sites.py`, 250 sites/archivo). | Ningún archivo del MVP se acerca a 250 sitios mutables. |
| Extended ruff ratchet (`check_ruff_ratchet.py`). | El ratchet asume 802 violations pre-existentes. El MVP arranca limpio. |
| `check_html_labels.py` (a11y). | 0 templates al principio. Se reactiva con `djlint` cuando se introduzcan. |
| E2E Playwright. | Sin UI funcional, no hay nada que probar. Se monta cuando el módulo `delivery` tenga flujo navegable. |
| Self-hosted ARM64 runner. | GitHub-hosted `ubuntu-latest` cubre los primeros meses. Self-hosted se evalúa al cruzar 2000 min/mes. |
| Deploy webhook. | MVP con deploy manual o vía Coolify directo. Webhook complica sin volumen. |

## Convenciones operativas (referencia)

Las ocho convenciones se defienden en prosa y se verifican con los gates listados. Esta sección las nombra; el detalle vive en los archivos que cada gate protege.

| Conf | Convención | Gate que la enforce |
|---|---|---|
| Conf-1 | Layer boundaries respetan `ALLOWED_IMPORTS` y `PURE_LAYERS`. | `check_layers.py`. |
| Conf-2 | `Annotated[T, Depends(get_x)]` para dependencias FastAPI. | `mypy` + ruff select. |
| Conf-3 | `# type: ignore[<código>]` con código explícito, nunca bare. | `mypy enable_error_code`. |
| Conf-4 | `# pragma: no cover` con razón en la línea adyacente. | Revisión de PR. |
| Conf-5 | `from __future__ import annotations` en cada `.py`. | Convención, sin script. |
| Conf-6 | PR ≤ 400 líneas; nombre de rama `^(feat\|fix\|refactor\|docs\|ci\|test)/<n>-<slug>$`. | `check_pr_size.py`, `check_branch_name.py`. |
| Conf-7 | Migraciones siempre backward-compatibles con estrategia Expand and Contract. | D82; revisión de PR sobre `migrations/versions/*.py`. |
| Conf-8 | `coverage_gate.py` marca los `CRITICAL_HELPERS` del módulo. | `pytest --cov-fail-under=85` + plugin. |

## Quick map inverso

| Si necesita | Abra primero | Y luego consulte |
|---|---|---|
| Entender qué gates corren y por qué | §Quality gates del MVP de Lanzadera. | §Anti-patrones documentados. |
| Saber si el hexagonal está siendo enforced | `scripts/check_layers.py`. | §Hexagonal layer gate. |
| Añadir una nueva rule de lint | §Configuración reproducible de ruff. | §Anti-patrones (issues #380, #389). |
| Configurar un nuevo scanner de seguridad | §Security scanning. | `.gitleaksignore`, `.trivyignore` cuando apliquen. |
| Diagnosticar por qué un gate falla | §Anti-patrones documentados. | El issue de APAP_WEB referenciado. |
| Planificar el setup desde cero | §Plan de implementación por día. | §Quality gates del MVP. |
| Diferir un gate «para más tarde» | §Lo que NO se porta al MVP y por qué. | La columna «Razón de exclusión» correspondiente. |
| Auditar el estado de los gates hoy | `gh run list --workflow=ci.yml` y `--workflow=security.yml`. | `tests/test_ci_workflow.py` (wiring pins). |

## Lectura recomendada

1. `docs/09-arquitectura-objetivo-y-principios.md` — el contrato arquitectónico que estos gates verifican.
2. `docs/08-decisiones-y-preguntas-abiertas.md` — D8 (hexagonal global), D66-D68 (stack y monolito), D82 (Expand and Contract).
3. Este documento, en el orden de las secciones.
4. `docs/03-aplicaciones/<app>/epic.md` — la épica del módulo que esté migrando, para ver cómo el gate aplica en cada caso.
5. APAP_WEB issues listados en §Anti-patrones — el origen de cada decisión.

## Referencias

| Recurso | Ruta o enlace |
|---|---|
| Repositorio APAP_WEB (issues cerradas: #380, #381, #393, #424, #428, #436, #437, #441, #442, #443). | `https://github.com/ardelperal/APAP_WEB` |
| Decisiones de plataforma (D8, D66, D67, D68, D73, D82). | `docs/08-decisiones-y-preguntas-abiertas.md` |
| Skill de documentación aplicada. | `documentation-alan-style` |

## Lista de comprobación final

- [ ] Los doce gates del MVP están commiteados, wired en `ci.yml` y pinned por test.
- [ ] `scripts/check_layers.py` clasifica los layers y verifica `ALLOWED_IMPORTS`, `PURE_LAYERS` y slicing.
- [ ] `pyproject.toml` declara pin exacto de `ruff` y `mypy`.
- [ ] `coverage_gate.py` registra `CRITICAL_HELPERS` del módulo `lanzadera.auth`.
- [ ] `.github/workflows/security.yml` y `security-deep.yml` corren scanners con imágenes pinned por digest.
- [ ] `Dockerfile` usa digest en cada `FROM`.
- [ ] El `Makefile` expone `make lint`, `make typecheck`, `make test`, `make check-layers`, `make check-complexity`, `make check-pr-size`, `make check-branch-name`, `make security`.
- [ ] La sección §Convenciones operativas referencia cada gate, sin duplicar su definición.
- [ ] El `README.md` y `CODEBASE-GUIDE.md` enlazan a este documento desde su índice de calidad.
