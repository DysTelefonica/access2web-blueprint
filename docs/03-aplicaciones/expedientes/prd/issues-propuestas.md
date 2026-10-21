# Issues propuestas — Expedientes

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Issue:** #1073 · **Fase:** 5 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Unidades de trabajo que se derivan del PRD. **No se ha creado ninguna issue**: son propuestas para que el mantenedor las apruebe, las fusione con las existentes o las descarte. No se abre ninguna issue en el repositorio legacy de Expedientes. Cada fila debe publicarse con etiqueta `app/expedientes`; las marcadas «transversal» además con `cross-cutting` y coordinadas con las demás aplicaciones, para no duplicar capacidades compartidas.

## Relación con el backlog existente

El [backlog de migración](../issue-backlog.md) ya cubre la implementación web de las 63 capacidades (issues #172 a #283) y la entrega funcional abierta (#801 a #813). Las propuestas siguientes **no las duplican**: nacen de lo que RC1 aporta de nuevo (hallazgos, coexistencia, tipos de datos, ERD de origen). Donde hay solapamiento parcial se indica la issue existente.

## Propuestas

| Id | Propuesta | Tipo | Alcance transversal | Depende de | Solapa con |
|---|---|---|---|---|---|
| P-EXP-1 | Extraer del backend de RC1 los metadatos de solo lectura (tablas, claves, índices, relaciones, nulabilidad) | verificación | no | entorno con `dysflow` y acceso al backend | #813 (extractor de solo lectura) |
| P-EXP-2 | Verificar en el binario los hallazgos de impacto alto del legacy | verificación | no | RC1 compilado y entorno aislado | — |
| P-EXP-3 | Decidir cómo accede Access a los datos durante la coexistencia (`getdb()` frente a vínculos) | decisión | sí | P-EXP-1 | — |
| P-EXP-4 | Definir una fuente única de identificadores compartida por Access y la web | decisión | sí | P-EXP-3 | #804 |
| P-EXP-5 | Aprobar los tipos objetivo (importes, fechas, `Sí`/`No`, nulos) | decisión | sí | P-EXP-1 | — |
| P-EXP-6 | Política de conflictos de edición concurrente con versión de fila | diseño e implementación | no | P-EXP-3 | #801 (edición) |
| P-EXP-7 | Replicar en la web los efectos laterales de `Registrar` | implementación | no | P-EXP-5 | #801, #804 |
| P-EXP-8 | Convertir reglas de unicidad y de «en uso» en restricciones declaradas | diseño | no | P-EXP-1 | #803 |
| P-EXP-9 | Unificar la identidad de usuario en auditoría y tablas E2E | decisión | sí | — | #811 |
| P-EXP-10 | Plan y pruebas de aceptación bidireccionales Access–web | pruebas | no | P-EXP-3 | #282, #283 |
| P-EXP-11 | Incorporar al ledger las candidatas sin entrada propia | documentación | no | — | #807, #814 |
| P-EXP-12 | Decidir la disposición de las tablas históricas y de las E2E del snapshot | decisión de negocio | no | P-EXP-1 | — |

## Detalle

### P-EXP-1 · Extracción de metadatos del backend de RC1

- **Evidencia:** el ERD de origen no pudo extraerse (D-EXP-4 del [README](README.md)); la [matriz](matriz-campos.md) usa columnas del snapshot sin claves ni nulabilidad.
- **Criterios de aceptación:** metadatos de las 49 tablas locales de RC1 (columnas, tipos, longitudes, nulabilidad, valores por defecto, claves, índices, relaciones declaradas) en un artefacto versionado; extracción solo de lectura, sin escrituras en el `.accdb`; diferencias frente a `ERD/schema.sql` listadas.
- **Riesgo:** el backend real no está en el clon de RC1; requiere acceso a un entorno aislado.

### P-EXP-2 · Verificación de hallazgos de impacto alto

- **Evidencia:** [hallazgos del README](README.md#hallazgos-del-legacy). Prioridad: `PersistirArbolDesdeMemoria` vacío (H-EXP-17), `ColHitos` sin persistir en altas, anexos de altas, botón de guardar del alta HPS deshabilitado, handler de guardado de grado de clasificación, comprobación de «en uso» de oficina de programa y de órgano de contratación.
- **Criterios de aceptación:** por hallazgo, resultado reproducido o descartado con captura o descripción del paso; los confirmados pasan al PRD como comportamiento actual o defecto confirmado; ninguna corrección en el legacy.

### P-EXP-3 · Acceso a datos durante la coexistencia

- **Evidencia:** `getdb()` abre el `.accdb` por ruta y suma 303 llamadas en 62 ficheros ([coexistencia](coexistencia.md#1-cómo-accede-hoy-expedientes-a-los-datos)). Hay tres vías, sin recomendación: sustituir `getdb()` por `CurrentDb` con vínculos; backend sincronizado (descartada por el roadmap); y la hipótesis del propietario, que mantiene `getdb()` y vincula por ODBC las tablas dentro del mismo `.accdb`, medida en el [apartado 11](coexistencia.md#11-p-exp-3-verificación-de-la-hipótesis-del-propietario). **Abierta y escalada al propietario.** Dirección preferida del propietario, en estudio y no cerrada: un PostgreSQL con esquemas modernos y un esquema legacy sincronizados por triggers (opción C del [apartado 12](coexistencia.md#12-p-exp-3-opciones-de-arquitectura-de-datos), junto con A y B y los requisitos del esquema legacy).
- **Criterios de aceptación:** decisión escrita entre sustituir `getdb()` por `CurrentDb` sobre tablas vinculadas u otra vía, con coste, riesgo de transacciones y plan de reversión; coordinada con las demás aplicaciones que comparten el patrón.

### P-EXP-4 y P-EXP-5 · Identificadores y tipos

- **Evidencia:** `DameID` (`Max + 1`, 30 llamadas); importes `DOUBLE`; fechas sin zona; `Sí`/`No` como texto en 369 líneas ([coexistencia](coexistencia.md#3-identidad-y-generación-de-claves)).
- **Criterios de aceptación:** documento de decisión por cada punto; la [matriz](matriz-campos.md) se actualiza con los tipos aprobados.

### P-EXP-6 y P-EXP-7 · Concurrencia y efectos laterales

- **Evidencia:** sin bloqueo ni versión de fila; `Registrar` recalcula `TbExpedientesConEntidades`, tareas, último cambio y encola E2E fuera de la transacción ([coexistencia](coexistencia.md#6-reglas-que-hoy-solo-aplica-vba)).
- **Criterios de aceptación:** guardado obsoleto rechazado con mensaje; mismo resultado observable de `Registrar` tras un guardado desde la web; pruebas por efecto. Cada unidad ≤ 400 líneas.

### P-EXP-8 · Restricciones declaradas

- **Evidencia:** unicidad y «en uso» solo comprobadas por consultas previas, con hallazgos de comprobaciones que no se ejecutan (oficina de programa, órgano de contratación).
- **Criterios de aceptación:** lista de restricciones propuestas por tabla con su origen en `src` y su efecto sobre las altas que hoy no persisten; aprobadas antes de generar DDL.

### P-EXP-9 · Identidad de usuario

- **Evidencia:** `IDUsuarioCreacion` es `TEXT(255)` en `TbExpedientes` y las tablas E2E usan `LONG` (H-EXP-80). Debe alinearse con el principal que expone Lanzadera.

### P-EXP-10 · Aceptación bidireccional

- **Criterios de aceptación:** las nueve pruebas de la [coexistencia](coexistencia.md#9-pruebas-de-aceptación-bidireccionales-propuestas) automatizadas o con guion manual, en entorno aislado y nunca sobre producción.

### P-EXP-11 · Ledger

- **Criterios de aceptación:** decisión sobre cada candidata de [capacidades](capacidades.md#capacidades-de-rc1-sin-entrada-propia-en-el-ledger): nueva entrada o subsunción en una existente.

### P-EXP-12 · Tablas históricas y E2E del snapshot

- **Evidencia:** seis tablas copia, auxiliares o históricas sin cita en `src` y cuatro tablas E2E del snapshot sin uso en RC1 ([resumen de la matriz](matriz-campos.md#resumen-por-tabla)).
- **Criterios de aceptación:** disposición aprobada por negocio para cada una (archivar, migrar o descartar), sin borrar datos.

## Escalación fuera del backlog

La clave y la configuración de un proyecto Firebase aparecen en claro en `Form_Formulario1` ([ficha](formularios/Form_Formulario1.md)). Conviene tratarlas como expuestas y rotarlas; es una decisión del propietario del legacy, no una issue de producto en Expedientes. Las fichas no copian ningún valor.

## Pendientes

- Aprobar, fusionar o descartar cada propuesta antes de crear issues.
- Comprobar que ninguna propuesta transversal duplica trabajo de otra aplicación.
