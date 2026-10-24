# 801-expedientes-lifecycle-vertical

- Repository: `DysTelefonica/access2web-blueprint` · Tracking issue: #801 (delivery HTTP + adapter PostgreSQL + DDL del ciclo de vida de Expedientes); épica #802.
- Locator: `odd/tasks/801-expedientes-lifecycle-vertical.md` · Engram mirror topic `odd/801-expedientes-lifecycle-vertical/tasks`.
- Worktree: `../access2web-blueprint-worktrees/801-expedientes-ddl` · base `origin/main` `f47cd1f3`.
- Decisión de producto (2026-10-24, dueño): la primera migración vertical pertenece a **#801** y usa **claves UUID**; #803 queda para las verticales de related-data, idempotencia y E2E. El criterio 8 de #803 («BIGSERIAL», `ID*`) sigue vigente para los catálogos legacy de `0002`, no para las tablas web nuevas, cuyo contrato es el de los puertos y el dominio (UUID en los 18 puertos; PK UUID en el esquema moderno de `0001`).

## Specs

S1. «Alcance: las migraciones DDL de las tablas de las verticales (related-data, E2E, cambios y económicos), con el mismo contrato que `0002`: una migración por vertical y `downgrade()` reversible en una transacción» — issue #803.
S2. «`alembic upgrade head` desde una base vacía crea el esquema de expedientes sin error» y «`alembic downgrade` revierte cada migración nueva sin residuos, en una transacción» — issue #803; D82.
S3. «El esquema no duplica ninguna de las 14 tablas de `0002`» — issue #803.
S4. «Las claves y los nombres siguen el criterio ya desplegado (`BIGSERIAL`, `ID*`)» — issue #803, re-leída por la decisión de producto: aplica a los catálogos legacy de `0002`; las tablas web nuevas usan UUID (S5).
S5. Columnas de `TbExpedientes` derivadas del agregado y sus invariantes (`app/src/modules/expedientes/domain/expediente.py`): `id` UUID PK, `tipo` ENUM (`AM`, `LOTE`, `BASED`), `estado` ENUM (`BORRADOR`, `ADJUDICADO`, `FORMALIZADO`, `ARCHIVADO`), `version` INTEGER NOT NULL, `id_expediente_padre` UUID NULL, `fecha_inicio_contrato` DATE NULL, `fecha_fin_contrato` DATE NULL, `created_at` TIMESTAMPTZ NULL, `updated_at` TIMESTAMPTZ NULL.
S6. Columnas de `TbCambios` derivadas de `ChangeRecord` (`app/src/modules/expedientes/ports/audit_log.py`): `id` UUID PK, `nombre_tabla` TEXT NOT NULL, `id_expediente` UUID NOT NULL, `nombre_campo` TEXT NULL, `valor_inicial` TEXT NULL, `valor_final` TEXT NULL, `fecha_cambio` TIMESTAMPTZ NOT NULL, `id_usuario_cambio` UUID NOT NULL, `accion` TEXT NOT NULL.
S7. Columnas de `TbUltimoCambio` (marcador de último cambio): `id` UUID PK, `id_expediente` UUID NOT NULL, `fecha_cambio` TIMESTAMPTZ NOT NULL, `id_usuario_cambio` UUID NOT NULL.
S8. La revisión `0009` encadena sobre `0008` (cabeza real: `0001`→`0002`→`0003`→`0004`→`0005`→`0006`→`0007`→`0008`); la `0010` encadena sobre `0009`. Una migración por vertical, sin ramas.
S9. Sin claves foráneas: el contrato del repo en estas migraciones es tabla + columnas (`0002`, `0008`) y una fila de auditoría debe sobrevivir a la baja del expediente.
S10. Prueba de integración con PostgreSQL real (nunca SQLite), siguiendo `tests/lanzadera/migrations/test_migration_0001.py`: base efímera por `CREATE DATABASE`/`DROP DATABASE`, `alembic` por subprocess con `DATABASE_URL` reescrito, aserciones sobre `information_schema` y `pg_enum`, y `APAP_INTEGRATION_ENABLED=1` como puerta fail-closed (issue #882).
S11. PR ≤ 400 líneas por unidad; sin `size:exception`; commits convencionales en inglés sin atribución de IA; merge de `origin/main` en la rama, nunca rebase ni force-push.

## Tasks

- T1 (S10) test(expedientes): add the shared PostgreSQL harness for the Expedientes migration tests — `tests/expedientes/migrations/_pg.py` + `conftest.py` + este tracker. Eslabón 1/3 de la cadena. Ruta: inline (sin autor de escritura independiente). Commit: pending.
- T2 (S2, S3, S4, S5, S8) feat(expedientes): create the Expedientes header table — migración `0009_expedientes_header.py` (`TbExpedientes`) + `test_migration_0009.py`. Eslabón 2/3. Ruta: inline. Commit: pending.
- T3 (S2, S6, S7, S8) feat(expedientes): create the Expedientes audit tables — migración `0010_expedientes_audit.py` (`TbCambios`, `TbUltimoCambio`) + test. Eslabón 3/3. Ruta: inline. Commit: pending.
- T4 (S1, #804) feat(expedientes): implement `ExpedienteRepositoryPg` sobre el schema. Ruta: pendiente de issue aprobada propia.
- T5 (S1, #801) feat(expedientes): delivery HTTP del ciclo de vida (201/400/403/409 en una transacción). Ruta: pendiente.
- T6 (#810) feat(expedientes): montar el módulo en el composition root con adapters reales y readiness. Ruta: pendiente.
- T7 (#803) feat(expedientes): verticales restantes (related-data, idempotencia/autosave, E2E). Ruta: pendiente.

Estado: T1 y T2 en curso en este turno; T3 preparado como eslabón 3.

## Log

L1. Petición original (verbatim, 2026-10-24): «sincronizate con el repo para estar alinados con todo lo que haya podido hacer andres asique bajate lo ultimo y despues vamos a volver hacer una planificacion de todo, y para poder ver como esta todo actualemnte y poder reportarloa mis jefes».
L2. Decisión de vía (verbatim): «Vertical de producto (Recomendado)» — implementar #803 → #804 → #801 → #810 con worktree y rama propios.
L3. Decisión de diseño (verbatim): «#801 con claves UUID (Recomendado)» — la cabecera y las tablas de auditoría de la vertical de ciclo de vida con PK UUID; #803 queda para related-data, idempotencia y E2E.
L4. Mapa previo (2026-10-24): la cabeza real de Alembic es `0008`; `0002` crea 14 catálogos legacy y remite las verticales a `0003..0049`; `design.md:61` advierte de no cerrar columnas finales sin revalidar las 49 tablas legacy, por lo que el DDL se dimensiona desde el contrato de código (dominio + puertos + specs) y la fidelidad legacy queda para el mapeo del extractor (#813).
L5. Troceo por presupuesto: la unidad completa (migración + test + arnés + tracker) suma 488 líneas, por encima del techo de 400 del repo. Se encadena en tres eslabones con cada cambio de producción acompañado de su prueba.
L6. RED observado (2026-10-24, PostgreSQL 16 local): sin la migración `0009` fallan los 4 tests (`7 failed, 1 error` en la primera versión; `4 failed` en la versión troceada). GREEN con la migración: `4 passed`.
L7. Los 13 rojos de la suite unitaria (`test_change_scope`, `test_check_deadcode`, `test_check_suites_populated`) también fallan en la base sin estos cambios: son falta de `vulture` e `import-linter` en el entorno local, no una regresión.
L8. Riesgo asumido: `TbDatosEconomicosExpedientes` no tiene requisito en `specs/` ni superficie de código, así que no entra en el plan de #803 sin decisión del dueño.
