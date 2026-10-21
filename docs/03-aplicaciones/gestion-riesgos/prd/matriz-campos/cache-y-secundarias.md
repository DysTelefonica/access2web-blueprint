# Cachés y tablas auxiliares — matrices campo a campo

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Issue:** #1079 · Fase 3 · **Matriz:** [matriz-campos.md](../matriz-campos.md)

7 tablas. Cada fila está en estado **propuesta**: la topología PostgreSQL global y el ERD de origen siguen abiertos (D-GR-4 del [README](../README.md)).

### TbArbolCacheMeta

- Columnas: **5**. Filas (snapshot): **0**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDEdicion | LONG(4) | — |
| CacheSetIdActivo | LONG(4) | — |
| BuiltAt | SHORT_DATE_TIME(8) | — |
| BuildMs | LONG(4) | — |
| Estado | TEXT(20) | — |

### TbAuxSumExp

- Columnas: **2**. Filas (snapshot): **6**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IdSumAux | LONG(4) | — |
| IDExp | LONG(4) | — |

### TbCacheArbolRiesgosMeta

- Columnas: **3**. Filas (snapshot): **31**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDEdicion | LONG(4) | — |
| ActiveBuildId | LONG(4) | — |
| UpdatedAt | SHORT_DATE_TIME(8) | — |

### TbCacheArbolRiesgosNodo

- Columnas: **16**. Filas (snapshot): **2089**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDEdicion | LONG(4) | — |
| BuildId | LONG(4) | — |
| NodeKey | TEXT(255) | — |
| ParentKey | TEXT(255) | — |
| NodeType | TEXT(30) | — |
| IDRiesgo | LONG(4) | — |
| IDMitigacion | LONG(4) | — |
| IDContingencia | LONG(4) | — |
| IDAccion | LONG(4) | — |
| EsVisibleSinRetirados | BOOLEAN(1) | — |
| TextConDescripcion | MEMO | — |
| TextSinDescripcion | MEMO | — |
| IconName | TEXT(255) | — |
| ForeColor | LONG(4) | — |
| SortIndex | LONG(4) | — |
| Depth | LONG(4) | — |

### TbCacheControlCambiosMeta

- Columnas: **5**. Filas (snapshot): **50**.
- Evidencia en `src`: **1** fichero(s) de producción — `ControlCambios.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDProyecto | LONG(4) | — |
| Edicion | LONG(4) | — |
| ActiveBuildId | LONG(4) | — |
| UpdatedAt | SHORT_DATE_TIME(8) | — |
| CacheVersion | TEXT(20) | — |

### TbCacheControlCambiosRow

- Columnas: **7**. Filas (snapshot): **321**.
- Evidencia en `src`: **1** fichero(s) de producción — `ControlCambios.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDProyecto | LONG(4) | — |
| Edicion | LONG(4) | — |
| BuildId | LONG(4) | — |
| CodigoRiesgo | TEXT(50) | — |
| EstadoHtml | MEMO | — |
| MitigacionHtml | MEMO | — |
| ContingenciaHtml | MEMO | — |

### TbCachePublicabilidadEdicion

- Columnas: **10**. Filas (snapshot): **0**.
- Evidencia en `src`: **1** fichero(s) de producción — `PublicabilidadEdicion.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDEdicion | LONG(4) | — |
| Tipo | TEXT(10) | — |
| IDRiesgo | LONG(4) | — |
| Publicable | BOOLEAN(1) | — |
| Veredicto | LONG(4) | — |
| CodigoRiesgo | TEXT(50) | — |
| Descripcion | MEMO | — |
| ChecksJson | MEMO | — |
| AlgoritmoVersion | LONG(4) | — |
| UpdatedAt | SHORT_DATE_TIME(8) | — |
