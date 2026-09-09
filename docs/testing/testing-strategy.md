[← Back to CODEBASE-GUIDE.md](../../CODEBASE-GUIDE.md) · [← Back to DOCS.md](../../DOCS.md) · [← Back to docs/calidad-de-codigo-y-ci.md](../calidad-de-codigo-y-ci.md)

# access2web-blueprint — Estrategia de testing

**Esta guía documenta la taxonomía de tests del MVP Lanzadera, los criterios de selección por capa y los gaps conocidos. El contrato de los gates mecánicos vive en [`docs/calidad-de-codigo-y-ci.md`](../calidad-de-codigo-y-ci.md). Aquí está el «por qué testeamos así» y el «dónde se rompe la disciplina hoy». El doc no describe cómo escribir un test en pytest ni cómo usar FastAPI TestClient — eso son frameworks cuyo uso se asume.**

> **Estado del doc:** v1.0 — audita 50 archivos reales en `tests/lanzadera/`, 5 categorías aplicables, 3 gaps abiertos.

## Sentence que organiza

> «Testeamos lo que rompe en producción. Una pieza testeada por debajo de su capa natural pierde señal; una pieza sobre-testada gasta tiempo de CI sin pagar riesgo real.»

## Las 5 categorías de test que este repo necesita

La suite de Lanzadera no necesita las 6 categorías que tendría un repo con browser e2e. Aquí las que aplica, con la regla que las separa:

| # | Categoría | Qué prueba | Velocidad | Fakes permitidos | Ruta esperada |
|---|---|---|---|---|---|
| 1 | Unit (domain) | Entidades, value objects, invariantes puras. | < 1 ms/test | Ninguno. Sólo `dataclasses` + `datetime`. | `tests/lanzadera/domain/test_*.py` |
| 2 | Unit (use case) | Use cases con repositorios en memoria. | < 10 ms/test | `FakeFixtures` (`tests/lanzadera/_fakes.py`). | `tests/lanzadera/application/test_*.py` |
| 3 | Contract (adapter) | El adapter cumple el `Protocol` del puerto. | < 50 ms/test | Adapter real, session Postgres mockeada. | `tests/lanzadera/adapters/test_*_contract.py` |
| 4 | Integration (HTTP) | Cadena HTTP route → container → use case → repo en memoria, vía `TestClient`. | < 200 ms/test | Repositorios en memoria; `auth_bypass` explícito. | `tests/lanzadera/delivery/test_*_routes*.py` |
| 5 | Integration (Postgres) | Round-trip real contra Postgres efímero. | ~1–3 s/test | Postgres real + migración Alembic sobre DB temporal. | Gated por `APAP_INTEGRATION_ENABLED=1`; skip silencioso en CI por defecto. |

**Lo que NO necesita este repo** (y por qué):

- **End-to-e2e con browser real (Playwright/Selenium).** No aplica a Lanzadera MVP porque las rutas admin devuelven HTML servido por plantillas Jinja; el UI visual se valida manualmente en staging (`docs/05-capacidades/expedientes/runtime.md` §UAT). Cuando se incorporen flujos con interacción rica (W60: SSE, W62: login web), esta categoría entra y se documenta aquí.
- **Migration test contra schema vivo en CI.** Existe `tests/lanzadera/migrations/test_migration_0001.py` y `test_migration_0001_downgrade.py`, pero corren contra SQLite en memoria, no contra Postgres real. Esa cobertura es suficiente mientras no haya migraciones con SQL específico de Postgres (D112: revisar cuando se introduzca `JSONB`).
- **Performance / load tests.** El cap de carga es «admins de una intranet departamental» (DA-1, `openspec/changes/lanzadera-mvp/proposal.md` §Stakeholders). Por debajo de 100 RPS, los tests de mutation semanal son la mejor inversión.

## Cómo elegir el tipo correcto — árbol de decisión

```text
¿Es lógica pura sin I/O?
  ├─ Sí → ¿Es de dominio (entidad, value object, error)?
  │        ├─ Sí → Categoría 1 (unit / domain)
  │        └─ No → ¿Es mapping legacy (legacy_role_map.py)?
  │                  └─ Sí → Categoría 1 sin red (DA-12 cubre la cardinalidad)
  └─ No → ¿Usa el `Protocol` de un puerto?
            ├─ Sí → ¿Existe impl real para Postgres?
            │        ├─ Sí, y la integración es barata → Categoría 5 con `APAP_INTEGRATION_ENABLED`
            │        └─ Sí, pero la impl es costosa → Categoría 3 (contract) + Categoría 5 manual
            │           └─ Categoría 5 manual corre en el job `mutation` semanal
            └─ No → ¿Es la capa delivery (HTTP route, CLI)?
                     ├─ Sí → Categoría 4 (integration HTTP) con fakes de repositorio
                     └─ No → Categoría 2 (unit / use case) con `FakeFixtures`
```

## Evidencia — qué hay hoy (auditoría 2026-08-29)

Resultado de `find tests/lanzadera -name 'test_*.py' | sort` clasificado por capa. Excluye `tests/fixtures/` (esos son fixtures sintéticos para los `check_*.py`, no tests del producto).

| Capa | Archivos | Categoría dominante | Test que la fija |
|---|---|---|---|
| `domain/` | 9 | 1 (unit) | `tests/lanzadera/domain/test_user.py` |
| `application/` | 3 | 2 (use case) | `tests/lanzadera/application/test_admin_use_cases.py` |
| `adapters/` | 8 | 3 (contract) + 5 (Postgres) | `tests/lanzadera/adapters/test_user_repository_pg.py` |
| `auth/` | 4 | 1 + 2 (hashing es puro, reset token no) | `tests/lanzadera/auth/test_credential_hasher_contract.py` |
| `delivery/` | 4 | 4 (HTTP integration) | `tests/lanzadera/delivery/test_admin_routes_integration.py` |
| `di/` | 1 | 4 (container wiring) | `tests/lanzadera/di/test_lanzadera_container.py` |
| `exp/` | 3 | 1 + 2 (extractor + UoW) | `tests/lanzadera/exp/test_extractor.py` |
| `migrations/` | 2 | 3 (schema round-trip) | `tests/lanzadera/migrations/test_migration_0001.py` |
| `notifications/` | 1 | 3 (mail queue contract) | `tests/lanzadera/notifications/test_mail_queue_contract.py` |
| `lanzadera/test_*.py` (root) | 15 | meta / gate-wiring | `tests/lanzadera/test_quality_report_smoke.py` |
| **TOTAL** | **50** | — | — |

Más `tests/test_*.py` raíz (gate-wiring del CI: `test_ci_workflow.py`, `test_gate_liveness.py`, `test_gate_smoke.py`, `test_check_workflows.py`, `test_app_stub.py`) — son tests del harness, no del producto.

## Gaps abiertos

### Gap G1 — Cobertura real de Postgres e2e en CI (severidad: media)

`tests/lanzadera/adapters/test_user_repository_pg.py::test_user_repository_pg_round_trip` existe y está bien escrito, pero corre sólo si `APAP_INTEGRATION_ENABLED=1`. En CI por defecto está desactivado. Resultado: cualquier regresión en SQL Postgres-specific (`JSONB`, `ON CONFLICT`, partial indexes) pasa el gate silencioso hasta que se descubre en staging.

- **Por qué está así hoy:** la suite Postgres e2e requiere un Postgres disponible en el runner self-hosted `a2w`. El runner actual no expone Postgres efímero todavía (issue #248 abierta).
- **Mitigación presente:** `check_mutation.py` corre semanalmente contra la suite (Categorías 1+2+3) y detecta asserts vacíos aunque SQL real no se ejecute. La mutation score ratchet cubre el riesgo principal (lógica auth, reglas de mapping legacy).
- **Recomendación:** cuando se cierre #248, mover el gate `APAP_INTEGRATION_ENABLED` a un job separado `quality-postgres` en `ci.yml`, paralelo a `quality` (no en serie; suma ~6 min al budget del runner).

### Gap G2 — Sin test de carrera / concurrencia (severidad: baja)

No hay tests para race conditions reales en `bootstrap_global_admins` (admite arranque concurrente del primer container) ni en `track_presence` (múltiples sesiones en paralelo). Ambos tienen tests unitarios que demuestran el caso secuencial.

- **Recomendación:** Categoría 5 con `pytest-asyncio` + `asyncio.gather` contra Postgres efímero. Esperar a G1 — sin Postgres real, este test es teatro.

### Gap G3 — Test de audit log append atómico (severidad: baja)

DA-11 declara que `AuditLogPort.append` debe ser atómico respecto al commit de la transacción que lo origina. El unit test (`test_audit_append_stamps_correlation_id`) verifica el contrato funcional; la atomicidad transaccional requiere Categoría 5 contra Postgres (mismo bloqueo que G1).

## Lo que sustituye un test que no existe

Aquí se aplica la regla «no testeamos lo que no se rompe». Cuando una pieza está ausente de la suite, queda documentada la **sustitución** (gates mecánicos, revisión manual o convención):

| Pieza sin test | Sustitución activa | Documentada en |
|---|---|---|
| Contrato del container DI | `tests/lanzadera/di/test_lanzadera_container.py` (sí existe — ejemplo de lo que sí testeamos). | — |
| Capas hexagonal purity | `scripts/check_layers.py` (QC-2, QC-9). | `docs/calidad-de-codigo-y-ci.md` §Hexagonal layer gate. |
| Naming de branches | `scripts/check_branch_name.py` (QC-6). | `docs/calidad-de-codigo-y-ci.md` §Los 12 check_*.py. |
| Tamaño de PR | `scripts/check_pr_size.py` (QC-6). | `docs/calidad-de-codigo-y-ci.md` §Los 12 check_*.py. |
| Workflow YAML drift | `scripts/check_workflows.py` (QC-9). | `docs/calidad-de-codigo-y-ci.md` §Los 12 check_*.py. |
| Migration Alembic contra Postgres real | Test Categoría 5 gated por `APAP_INTEGRATION_ENABLED`. | `docs/03-aplicaciones/lanzadera/migration/` (D112). |
| Visual UI / accesibilidad | Revisión manual en staging; planes UAT. | `docs/05-capacidades/expedientes/uat-cutover-legacy-retirement.md` §UAT. |

## Mutation semanal

`check_mutation.py` corre semanalmente (no por PR). Detecta:

- Lógica auth (hash, verify, reset, bootstrap) — debe mantener score alto por DA-2 + DA-13.
- Reglas de mapping legacy (DA-12) — la cardinalidad exhaustiva de `SinAcceso` exclusivo.
- Servicios cross-cutting (CachePort invalidate, AuditLogPort append atómico, DA-11).

Cuando la mutation score cae por debajo del umbral, `security-deep.yml` falla y crea issue automático.

## Pieza 4 — Propuesta de gate `check_test_classification.py`

**Estado:** propuesta, **no instalado en CI**. Aplicar sólo si la taxonomía de las 5 categorías se rompe en PRs.

### Qué enforzaría

| Regla | Detección |
|---|---|
| Cada test nuevo bajo `tests/lanzadera/<capa>/test_*.py` importa sólo fixtures coherentes con su capa. | AST walk + import audit. |
| `application/` no importa `unittest.mock.MagicMock`. | grep AST. |
| `delivery/test_*.py` que toca `admin_routes` incluye fixture `auth_bypass` o marca `@pytest.mark.requires_auth`. | Análisis de fixtures declaradas vs. funciones llamadas. |
| Test Categoría 5 (Postgres e2e) está gated por `APAP_INTEGRATION_ENABLED`. | Detección de `@pytest.mark.skip` con condicional. |
| Nombre del archivo respeta la capa (HR-8 de la skill `lanzadera-testing-strategy`). | Regex sobre el path. |

### Por qué no se instala todavía

- La suite actual está **bien clasificada** (la disciplina se mantiene manualmente). El gate tiene costo de mantenimiento (los falsos positivos surgen cuando alguien añade una capa nueva) y el valor marginal es bajo hasta que el equipo crezca.
- El gate debe instalarse **junto con la skill `lanzadera-testing-strategy`** (Pieza 2 de la epic, fuera del scope de este doc). Sin skill, los contribuidores no sabrían qué clase de fix aplicar cuando el gate falle.
- Cuando se cierre #248 (Postgres efímero en runner), una de las reglas (Categorías 5 deben estar gated) puede endurecerse a un gate obligatorio por PR.

### Plan de rollout

1. Aprobar este doc y la skill (Pieza 2).
2. Implementar `scripts/check_test_classification.py` + `tests/test_check_test_classification.py` en un PR aparte.
3. Dejar el gate en `--dry-run` durante un sprint; reportar falsos positivos.
4. Activar en CI sólo cuando la tasa de falsos positivos baje a cero.

## What this is

| Es | Evidence in this repo |
|---|---|
| Taxonomía de los 5 tipos de test que este repo necesita. | Tabla §Las 5 categorías, con archivo ancla por categoría. |
| Reglas de selección por capa (domain / application / adapters / delivery). | Árbol de decisión §Cómo elegir el tipo correcto. |
| Mapeo de gaps abiertos a su mitigación actual. | Sección §Gaps abiertos, una fila por gap con severidad y plan. |
| Puente entre el CodeGuide y `calidad-de-codigo-y-ci.md`. | §Cross-references al pie, con paths verificables. |

## What this is not

| No es | Límite |
|---|---|
| Una copia del `apap-testing-strategy` de `ardelperal/APAP_WEB`. | Este repo tiene 50 archivos en 5 categorías, no 265 en 6; la estructura se acorta pero no se duplica. |
| Réplica operativa de `docs/calidad-de-codigo-y-ci.md`. | Ese doc cubre los 12 `check_*.py` y los 5 workflows (operativo). Este cubre qué probar y por qué (estratégico). |
| Lista exhaustiva de los 50 archivos. | La lista vive en el árbol de `tests/lanzadera/`; este doc la resume por capa. |
| Manual de pytest, FastAPI TestClient ni Alembic. | Esos son frameworks; este doc los asume y prescribe su uso por capa. |

## Cross-references

- [`docs/calidad-de-codigo-y-ci.md`](../calidad-de-codigo-y-ci.md) — los 12 gates mecánicos y los 5 workflows. **Operativo**.
- [`docs/architecture.md`](../architecture.md) §CI gates — decisiones arquitectónicas detrás de cada gate. **Decisión**.
- [`CODEBASE-GUIDE.md`](../../CODEBASE-GUIDE.md) — ownership, flujos y guardarraíles del monorepo. **Orientación**.
- [`openspec/changes/lanzadera-mvp/design.md`](../../openspec/changes/lanzadera-mvp/design.md) §Pipeline de calidad — el contrato completo de los gates en el SDD original.
- [`skills/documentation-alan-style/SKILL.md`](../../skills/documentation-alan-style/SKILL.md) — el patrón que este doc sigue.
- [`skills/lanzadera-testing-strategy/SKILL.md`](../../skills/lanzadera-testing-strategy/SKILL.md) — la skill operativa que codifica las reglas de este doc (Pieza 2 de la epic, fuera del scope aquí).

## Verificación

Tras editar este doc, regenere la tabla §Evidencia:

```bash
find tests/lanzadera -name 'test_*.py' -not -path '*/fixtures/*' | wc -l   # debe dar 50 ±1
git ls-files docs/testing/testing-strategy.md                                # debe estar trackeado
```

## Core invariants

Reglas que un cambio no debe romper. Citable en un review:

- **Domain sin dependencias**: los tests de `tests/lanzadera/domain/` no importan `_fakes` ni sesiones de DB. Si un test de dominio necesita repositorio, está testeando la capa equivocada.
- **Use cases con fakes, no mocks**: `tests/lanzadera/application/` sustituye dependencias por `FakeFixtures`; `unittest.mock.MagicMock` queda prohibido porque pierde la señal del contrato del puerto.
- **Auth bypass explícito en integration HTTP**: los tests de `delivery/` que tocan `admin_routes` declaran la fixture `auth_bypass` con un mensaje que cita el gate real (M02 / A01–A03).
- **Categorías 5 son gated, no skipped**: un test que requiere Postgres real se salta sólo si `APAP_INTEGRATION_ENABLED` está definido y es falsy. Un skip permanente con TODO es deuda.
- **Mutation score no cae**: cuando un nuevo test reduce la mutation score por debajo del umbral, el PR no se mergea; la causa se diagnostica antes de añadir más código.

## Contributor checklist

- [ ] El doc vive en `docs/testing/` (taxonomía `documentation-alan-style` §Taxonomía de docs).
- [ ] Un único `H1`; ningún `H4` (HR-3 de `documentation-alan-style`).
- [ ] Cero coincidencias de `vos`, `ejecutá`, `corré`, `che`, `dale`, `listo` fuera de backticks.
- [ ] Toda afirmación tiene ruta de archivo o path concreto en la columna «Evidence» (regla falsable).
- [ ] Las 5 categorías están actualizadas si se añade una capa nueva o se cierra un gap.
- [ ] Las referencias cruzadas siguen apuntando a archivos que existen en `main`.
- [ ] La tabla §Evidencia refleja el conteo real (`find tests/lanzadera -name 'test_*.py' | wc -l`).
- [ ] El path del test ancla la taxonomía (ver §Frescura en la skill `documentation-alan-style`).

## Navigation

Previous: [← Back to CODEBASE-GUIDE.md](../../CODEBASE-GUIDE.md) · Next: [docs/calidad-de-codigo-y-ci.md →](../calidad-de-codigo-y-ci.md)
