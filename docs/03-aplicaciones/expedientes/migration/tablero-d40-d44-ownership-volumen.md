# Tablero D40–D44 — Ownership, volumen y uso de las 49 tablas

Tablero ejecutivo que consolida el estado de las 49 tablas del backend legacy `Expedientes_datos.accdb` en lo relativo a **ownership funcional, volumen cosechado y patrón de uso**. Documento de referencia para reabrir las preguntas que solo el runtime puede responder, y que el plan `expedientes-web-migration` marca como bloqueantes de F04 y M01.

**Design**: `openspec/changes/expedientes-web-migration/design.md` (D-EXP-1..9; gap "49 tablas y campos no revalidados").
**Tasks**: `openspec/changes/expedientes-web-migration/tasks.md` (D40–D44 dictionary tranches).
**Diccionarios fuente**: `docs/03-aplicaciones/expedientes/migration/source-dictionary-{01..05}.md`.
**Modelo físico**: `docs/03-aplicaciones/expedientes/data-model.md`.
**Schema DDL**: `docs/03-aplicaciones/expedientes/ERD/schema.sql`.

## Resumen ejecutivo

| Categoría | Tablas | Cardinalidad cosechada | Pendiente de runtime |
|---|---:|---:|---:|
| Diccionario principal (01–49, sin re-entries) | 39 | ~5.250 filas documentadas | 6 tablas con cardinalidad `—` |
| Re-entries documentados (41–49) | 9 | Alias de ordinales 02–13 | 0 |
| Esquema total cosechado | 48 entradas únicas | 14 tablas con cardinalidad explícita | Pendiente: 6 marcadas `—` en los diccionarios |

### Estado por ordinal (49 entradas, sin duplicar re-entries)

La columna **Volumen** refleja las filas cosechadas por `dysflow count_rows` el 2026-08-05 (cuando aplica). La columna **Ownership** combina lo declarado en cada tranche y la asignación implícita por capacidad destino. La columna **Patrón de uso** es la pregunta abierta D40–D44 que solo el runtime puede responder; se marca `[ABIERTO]` cuando no se ha documentado.

| Ord | Tabla legacy | Capacidad destino | Volumen | Ownership | Patrón de uso | Acción | Tranche |
|---:|---|---|---:|---|---|---|---|
| 01 | `Copia de TbExpedientes` | — | 0 | legacy | — | Rechazar | 01 |
| 02 | `Copia de TbExpedientesConEntidades` | — | 365 | legacy | — | Rechazar | 01 |
| 03 | `ListaPrevia` | — | 99 | legacy | — | Rechazar | 01 |
| 04 | `TbAusExpPostAGEDO` | EXP-CAP-051 | 203 | Integraciones (a confirmar) | `[ABIERTO]` | Cuarentena | 01 |
| 05 | `TbAuxEstadosMartina` | — | 143 | legacy | — | Rechazar | 01 |
| 06 | `TbAuxNemotecnico` | EXP-CAP-007 | 0 | legacy | — | Transformar | 01 |
| 07 | `TbCambios` | EXP-CAP-046 | 0 | Auditoría | `[ABIERTO]` | Migrar | 01 |
| 08 | `TbComerciales` | EXP-CAP-015 | 26 | Catálogos | `[ABIERTO]` | Migrar | 01 |
| 09 | `TbConfMostrarEstado` | EXP-CAP-005 | 23 | Preferencias UI (a confirmar) | `[ABIERTO]` | Transformar | 01 |
| 10 | `TbCPV` | EXP-CAP-016 | 73 | Catálogos | `[ABIERTO]` | Migrar | 01 |
| 11 | `TbRACS` | EXP-CAP-018 | 37 | Catálogos | `[ABIERTO]` | Migrar | 02 |
| 12 | `TbResponsablesPorRol` | EXP-CAP-019 | 11 | Auth (a confirmar) | `[ABIERTO]` | Migrar | 02 |
| 13 | `TbSuministradores` | EXP-CAP-015 | 72 | Catálogos | `[ABIERTO]` | Migrar | 02 |
| 14 | `TbExpedientes` | EXP-CAP-001 | 453 | **Expediente principal** | `[ABIERTO — necesita runtime]` | Migrar | 02 |
| 15 | `TbExpedientesConEntidades` | EXP-CAP-001 (vista) | 451 | Read-models | `[ABIERTO — necesita runtime]` | VIEW | 02 |
| 16 | `TbExpedientesAnualidades` | EXP-CAP-008 | 174 | Financiero | `[ABIERTO]` | Migrar | 02 |
| 17 | `TbExpedientesAnexos` | EXP-CAP-009 | 712 | Anexos | lectura/escritura mixta | Migrar | 02 |
| 18 | `TbExpedientesCadenaContratacion` | EXP-CAP-010 | — | Contratación | `[ABIERTO]` | Migrar | 02 |
| 19 | `TbExpedientesCodigoCompras` | EXP-CAP-011 | — | Contratación | `[ABIERTO]` | Migrar | 02 |
| 20 | `TbExpedientesComerciales` | EXP-CAP-012 | 333 | Catálogos/Expediente | `[ABIERTO]` | Migrar | 02 |
| 21 | `TbExpedientesCPVs` | EXP-CAP-013 | 429 | Catálogos/Expediente | `[ABIERTO]` | Migrar | 03 |
| 22 | `TbExpedientesE2E` | EXP-CAP-050 | — | E2E | `[ABIERTO]` | Migrar | 03 |
| 23 | `TbExpedientesHitos` | EXP-CAP-014 | 46 | Calendario | `[ABIERTO]` | Migrar | 03 |
| 24 | `TbExpedientesJefaturas` | EXP-CAP-006 | — | Jefaturas | `[ABIERTO]` | Migrar | 03 |
| 25 | `TbExpedientesJuridicas` | EXP-CAP-007 | 417 | Jurídicas | `[ABIERTO]` | Migrar | 03 |
| 26 | `TbExpedientesLugaresEjecucion` | EXP-CAP-006 | 194 | Lugares | `[ABIERTO]` | Migrar | 03 |
| 27 | `TbExpedientesModificados` | EXP-CAP-046 | 37 | Auditoría | `[ABIERTO]` | Migrar | 03 |
| 28 | `TbExpedientesPECAL` | EXP-CAP-018 | 366 | Catálogos/Expediente | `[ABIERTO]` | Migrar | 03 |
| 29 | `TbExpedientesRACS` | EXP-CAP-018 | — | RACS | `[ABIERTO]` | Migrar | 03 |
| 30 | `TbExpedientesResponsables` | EXP-CAP-019 | — | Responsables | `[ABIERTO]` | Migrar | 03 |
| 31 | `TbExpedientesSuministradores` | EXP-CAP-015 | 72 | Suministradores | `[ABIERTO]` | Migrar | 04 |
| 32 | `TbCambios` (re-entry) | EXP-CAP-046 | — | Auditoría | `[ABIERTO — necesita runtime]` | Bloqueado | 04 |
| 33 | `TbUltimoCambio` | EXP-CAP-046 | — | Auditoría (a confirmar) | `[ABIERTO]` | Cuarentena | 04 |
| 34 | `TbConfMostrarEstado` (re-entry) | EXP-CAP-005 | 23 | Preferencias UI | `[ABIERTO]` | Cuarentena | 04 |
| 35 | `TbE2EExportBatch` | EXP-CAP-050 | — | Operación E2E | `[ABIERTO]` | Cuarentena | 04 |
| 36 | `TbE2EExportBatchDetalle` | EXP-CAP-050 | — | Operación E2E | `[ABIERTO]` | Cuarentena | 04 |
| 37 | `TbE2EExportSeleccionTemp` | EXP-CAP-050 | — | Operación E2E (temporal) | `[ABIERTO]` | Rechazar | 04 |
| 38 | `TbE2EJsonDestinationUserConfig` | EXP-CAP-050 | — | Configuración E2E | `[ABIERTO]` | Cuarentena | 04 |
| 39 | `TbExpedientesE2E` (re-entry ordinal) | — | — | — | — | RE-ENTRY (→ 22) | 04 |
| 40 | `TbDatosEconomicosExpedientes` | EXP-CAP-001 | 91 | **Financiero / Expediente principal** | `[ABIERTO — necesita runtime]` | Bloqueado | 04 |
| 41 | `Comerciales` (re-entry) | — | — | — | — | RE-ENTRY (→ 02) | 05 |
| 42 | `CPVs` (re-entry) | — | — | — | — | RE-ENTRY (→ 03) | 05 |
| 43 | `RACS` (re-entry) | — | — | — | — | RE-ENTRY (→ 11) | 05 |
| 44 | `PECAL` (re-entry) | — | — | — | — | RE-ENTRY (→ 10) | 05 |
| 45 | `LugaresEjecucion` (re-entry) | — | — | — | — | RE-ENTRY (→ 07) | 05 |
| 46 | `Responsables` (re-entry) | — | — | — | — | RE-ENTRY (→ 12) | 05 |
| 47 | `Suministradores` (re-entry) | — | — | — | — | RE-ENTRY (→ 13) | 05 |
| 48 | `Jefaturas` (re-entry) | — | — | — | — | RE-ENTRY (→ 05) | 05 |
| 49 | `Juridicas` (re-entry) | — | — | — | — | RE-ENTRY (→ 06) | 05 |

> Las tablas **41–49** son aliases del legacy Access que `data-model.md` ya re-mapea contra las posiciones 02–13 (Tranche 05). El ordinal mapping queda cerrado y no requieren decisiones adicionales.

## Hallazgos ya documentados (no requieren runtime)

Los siguientes descubrimientos críticos del modelo físico **ya están cerrados** en `data-model.md` §Hallazgos críticos del esquema de Expedientes y se conservan aquí para que el tablero los muestre en contexto:

1. **D102 cross-cutting**: veinte o más columnas almacenan `Sí/No` como `Text 2)` en lugar de `Boolean`. La normalización a `BOOLEAN` en PostgreSQL aplica al extracto de staging (M01) y al schema DDL (F04). Aplica a 8 tablas del tablero: 05, 09, 17*(parcial)*, 13, 31, 14, 18, 25.
2. **D113 (Expedientes)**: `TbExpedientesConEntidades` (ordinal 15) es una vista desnormalizada con 13 columnas `Cadena*` separadas por comas. El traslado es a una `VIEW` PostgreSQL generada a partir de las tablas normalizadas, no a una tabla.
3. **D114 (Expedientes)**: múltiples sistemas E2E con hash conviven (`TbExpedientes.HashActual`, `TbExpedientes.HashUltimaExportacion`, `TbExpedientesE2E.HashPayload`). La consolidación en una sola tabla `expediente_sincronizacion_e2e` es parte del spec del vertical E2E.
4. **Multi-impuestos** (IVA, IPSI, IGIC) en `TbExpedientesAnualidades` (ordinal 16). Migrar con `NUMERIC(15,2)` para conservar precisión regulatoria.
5. **Typo legacy**: `Descripcon` en `TbExpedientesSuministradores` (ordinal 31). El extracto de staging debe corregirlo a `descripcion` antes del mapeo.

## Preguntas ABIERTAS — solo el runtime puede responder

Las preguntas pendientes requieren acceso al backend autoritativo (`C:\00repos\datos\Expedientes_datos.accdb`) y se mapean a las decisiones D40–D44 del plan. Cada pregunta se cierra al re-ejecutar `dysflow map_form_behavior` o equivalentes y al volcar el resultado en `data-model.md` y el tranche correspondiente.

### D40 — Ownership funcional

**Pregunta**: ¿Quién es el dueño funcional de cada tabla a día de hoy? El extracto de staging indica el equipo actual, pero el ownership **del módulo web** debe asignarse al equipo que mantiene la capacidad destino.

| Tabla | Pregunta concreta |
|---|---|
| 04 — `TbAusExpPostAGEDO` | ¿Equipo de Integraciones mantiene el contrato AGEDYS, o la mantiene el equipo de Expedientes? |
| 09 — `TbConfMostrarEstado` | ¿La preferencia UI es decisión de producto (UX) o de configuración legacy sin dueño? |
| 12 — `TbResponsablesPorRol` | ¿La mantiene el equipo de Auth (Lanzadera) o el de Expedientes? Hay FK a usuarios del sistema Lanzadera. |
| 33 — `TbUltimoCambio` | ¿Sigue siendo relevante en el nuevo esquema (event sourcing) o se reemplaza por el log de auditoría `AuditLogPort` (D-EXP-3)? |

### D41 — Volumen real

**Pregunta**: ¿Cuál es la cardinalidad real de las 6 tablas marcadas `—` en el tablero? La métrica exigida es **filas totales hoy, no al 2026-08-05**.

| Tabla | Pregunta concreta |
|---|---|
| 18 — `TbExpedientesCadenaContratacion` | Filas totales. Cardinalidad del join con `TbExpedientes` (1-N, N-N, jerárquica). |
| 19 — `TbExpedientesCodigoCompras` | Filas totales. ¿Catálogo de códigos o join con `TbExpedientes`? |
| 22 — `TbExpedientesE2E` | Filas totales. ¿Una fila por expediente o por evento E2E? |
| 24 — `TbExpedientesJefaturas` | Filas totales. |
| 29 — `TbExpedientesRACS` | Filas totales. |
| 30 — `TbExpedientesResponsables` | Filas totales. |

### D42 — Patrón de uso (lectura / escritura / frecuencia)

**Pregunta**: ¿Cuál es el patrón real de acceso por tabla? La asignación de `Migrar` o `Cuarentena` cambia si una tabla de "escribir una vez y leer mil veces" se trata distinto a "escribir y leer igual de frecuente".

| Tabla | Pregunta concreta |
|---|---|
| 07 — `TbCambios` | ¿Cuántas escrituras por minuto en horario pico? (Impacto: el extracto debe respetar watermarks.) |
| 13 — `TbSuministradores` | Lectura/escritura. ¿Actualizaciones masivas o fila a fila? |
| 17 — `TbExpedientesAnexos` | Frecuencia de subida (escritura) frente a descarga (lectura). Decide caché y CDN. |
| 27 — `TbExpedientesModificados` | Tasa de modificación por expediente (¿1.1 vs 1.5?). Define el costo de versionado en F04. |
| 35..38 — `TbE2E*` | Operación batch. ¿Frecuencia manual, nocturna, evento? Decide el ciclo del batch E2E (gap #6). |

### D43 — Tablas huérfanas y cardinalidad agregada

**Pregunta**: ¿Hay tablas con cardinalidad `0` que están en desuso real (no solo en la快照 `Expedientes_datos.accdb`)? Las marcadas `0` pueden ser herencia o estar activas en producción real.

| Tabla | Pregunta concreta |
|---|---|
| 01, 06 | ¿`Copia de TbExpedientes` y `TbAuxNemotecnico` están vacías en producción o solo en la copia local? |
| 07 | `TbCambios` declarada con 0 filas el 2026-08-05. ¿Auditoría real o fallback apagado? |

### D44 — Transformaciones pendientes

**Pregunta**: Para las tablas marcadas `Transformar` (09, 06) o `Bloqueado` (32, 40), ¿cuál es la regla de transformación exacta?

| Tabla | Pregunta concreta |
|---|---|
| 09 — `TbConfMostrarEstado` | ¿Cada preferencia UI migra a `preferencias_usuario` o se descarta por defecto? |
| 32 — `TbCambios` | ¿Se renombra a `expediente_cambios` o se consolida con el `AuditLogPort`? |
| 40 — `TbDatosEconomicosExpedientes` | ¿Se une a `TbExpedientes` (columnas nuevas) o queda como tabla aparte? |

## Cómo cerrar las preguntas

Cada pregunta cierra cuando se ejecuta `dysflow map_form_behavior` (o `count_rows` para cardinalidad) contra `C:\00repos\datos\Expedientes_datos.accdb` y el resultado se vuelca al tranche correspondiente del dictionary (`source-dictionary-{01..05}.md`) en formato `Ownership: <equipo>` / `Volumen: <filas>` / `Patrón: <lectura|escritura|mixta>`. Hasta entonces, las columnas marcadas `[ABIERTO]` se mantienen explícitamente vacías.

El tablero se regenera cuando se cierra al menos uno de los seis gaps enumerados arriba; la versión vigente vive en este archivo y se referencia desde `design.md` §Decisiones pendientes.

## Auditoría de origen

- **Fuente legacy**: `C:\00repos\datos\Expedientes_datos.accdb` (no versionada en este repositorio).
- **Extracto cosechado**: `docs/03-aplicaciones/expedientes/ERD/schema.sql` (PR #393, ejecutado con Jackcess 2026-08-18).
- **Diccionario**: `docs/03-aplicaciones/expedientes/data-model.md` (49 entradas, 20 detalladas).
- **Diccionarios por tranches**: `docs/03-aplicaciones/expedientes/migration/source-dictionary-{01..05}.md` (5 archivos, 1.073 líneas).
- **Cardinalidad**: `dysflow count_rows` ejecutado el 2026-08-05 y volcado en `data-model.md` + cada tranche.
- **Próxima cosecha sugerida**: tras cerrar las 6 preguntas, re-ejecutar `dysflow` y actualizar este tablero en un PR dedicado.
