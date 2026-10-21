# Planes de mitigación y de contingencia — matrices campo a campo

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Issue:** #1079 · Fase 3 · **Matriz:** [matriz-campos.md](../matriz-campos.md)

6 tablas. Cada fila está en estado **propuesta**: la topología PostgreSQL global y el ERD de origen siguen abiertos (D-GR-4 del [README](../README.md)).

### TbRiesgosPlanContingenciaDetalle

- Columnas: **11**. Filas (snapshot): **1817**.
- Evidencia en `src`: **4** fichero(s) de producción — `Edicion.cls`, `PC.cls`, `PCAccion.cls`, `Constructor.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDAccionContingencia | LONG(4) | — |
| IDContingencia | LONG(4) | — |
| CodAccion | TEXT(255) | — |
| Accion | MEMO | — |
| ResponsableAccion | TEXT(255) | personal |
| FechaInicio | SHORT_DATE_TIME(8) | — |
| FechaFinPrevista | SHORT_DATE_TIME(8) | — |
| FechaFinReal | SHORT_DATE_TIME(8) | — |
| EsUltimaAccion | TEXT(2) | — |
| Estado | TEXT(255) | — |
| NombreIcono | TEXT(255) | — |

### TbRiesgosPlanContingenciaDetalleReversa

- Columnas: **6**. Filas (snapshot): **27**.
- Evidencia en `src`: **3** fichero(s) de producción — `Edicion.cls`, `PCAccionReversa.cls`, `Constructor.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDAccionContingencia | LONG(4) | — |
| IDEdicion | LONG(4) | — |
| FechaInicio | SHORT_DATE_TIME(8) | — |
| FechaFinPrevista | SHORT_DATE_TIME(8) | — |
| FechaFinReal | SHORT_DATE_TIME(8) | — |
| Estado | TEXT(255) | — |

### TbRiesgosPlanContingenciaPpal

- Columnas: **8**. Filas (snapshot): **1028**.
- Evidencia en `src`: **4** fichero(s) de producción — `Edicion.cls`, `PC.cls`, `Constructor.bas`, `PublicabilidadRiesgo.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDContingencia | LONG(4) | — |
| IDRiesgo | LONG(4) | — |
| CodContingencia | TEXT(255) | — |
| DisparadorDelPlan | MEMO | — |
| FechaDeActivacion | SHORT_DATE_TIME(8) | — |
| FechaDesactivacion | SHORT_DATE_TIME(8) | — |
| Estado | TEXT(255) | — |
| NombreIcono | TEXT(255) | — |

### TbRiesgosPlanMitigacionDetalle

- Columnas: **11**. Filas (snapshot): **4221**.
- Evidencia en `src`: **4** fichero(s) de producción — `Edicion.cls`, `PM.cls`, `PMAccion.cls`, `Constructor.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDAccionMitigacion | LONG(4) | — |
| IDMitigacion | LONG(4) | — |
| CodAccion | TEXT(255) | — |
| Accion | MEMO | — |
| ResponsableAccion | TEXT(255) | personal |
| FechaInicio | SHORT_DATE_TIME(8) | — |
| FechaFinPrevista | SHORT_DATE_TIME(8) | — |
| FechaFinReal | SHORT_DATE_TIME(8) | — |
| EsUltimaAccion | TEXT(2) | — |
| Estado | TEXT(255) | — |
| NombreIcono | TEXT(255) | — |

### TbRiesgosPlanMitigacionDetalleReversa

- Columnas: **6**. Filas (snapshot): **41**.
- Evidencia en `src`: **3** fichero(s) de producción — `Edicion.cls`, `PMAccionReversa.cls`, `Constructor.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDAccionMitigacion | LONG(4) | — |
| IDEdicion | LONG(4) | — |
| FechaInicio | SHORT_DATE_TIME(8) | — |
| FechaFinPrevista | SHORT_DATE_TIME(8) | — |
| FechaFinReal | SHORT_DATE_TIME(8) | — |
| Estado | TEXT(255) | — |

### TbRiesgosPlanMitigacionPpal

- Columnas: **8**. Filas (snapshot): **2843**.
- Evidencia en `src`: **5** fichero(s) de producción — `Edicion.cls`, `PC.cls`, `PM.cls`, `Constructor.bas`, `PublicabilidadRiesgo.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDMitigacion | LONG(4) | — |
| IDRiesgo | LONG(4) | — |
| CodMitigacion | TEXT(255) | — |
| DisparadorDelPlan | MEMO | — |
| FechaDeActivacion | SHORT_DATE_TIME(8) | — |
| FechaDesactivacion | SHORT_DATE_TIME(8) | — |
| Estado | TEXT(255) | — |
| NombreIcono | TEXT(255) | — |
