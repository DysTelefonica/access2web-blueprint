# Catálogos y valores posibles — matrices campo a campo

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Issue:** #1079 · Fase 3 · **Matriz:** [matriz-campos.md](../matriz-campos.md)

9 tablas. Cada fila está en estado **propuesta**: la topología PostgreSQL global y el ERD de origen siguen abiertos (D-GR-4 del [README](../README.md)).

### TbConfiguracionVisionRiesgos

- Columnas: **4**. Filas (snapshot): **35**.
- Evidencia en `src`: **1** fichero(s) de producción — `Entorno.cls`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDConfiguracion | LONG(4) | — |
| Usuario | TEXT(255) | personal |
| VerSoloNoRetirados | TEXT(2) | — |
| VerDescripcion | TEXT(2) | — |

### TbFelix

- Columnas: **2**. Filas (snapshot): **47**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDriesgo | LONG(4) | — |
| Prioridad | LONG(4) | — |

### TbHerramientaDocAyuda

- Columnas: **2**. Filas (snapshot): **26**.
- Evidencia en `src`: **1** fichero(s) de producción — `Constructor.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| NombreFormulario | TEXT(255) | — |
| NombreArchivoAyuda | TEXT(255) | — |

### TbTablas

- Columnas: **9**. Filas (snapshot): **41**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| Nombre | TEXT(255) | — |
| Orden | LONG(4) | — |
| Origen | TEXT(255) | — |
| NombreTablaOrigenID | TEXT(255) | — |
| NombreID1 | TEXT(255) | — |
| NombreID2 | TEXT(255) | — |
| TablaID | TEXT(255) | — |
| EsTablaID | TEXT(2) | — |
| RequiereID | TEXT(2) | — |

### TbValoresPosiblesContingencia

- Columnas: **1**. Filas (snapshot): **4**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| Valores | TEXT(255) | — |

### TbValoresPosiblesEstadoRiesgo

- Columnas: **1**. Filas (snapshot): **7**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| Valores | TEXT(255) | — |

### TbValoresPosiblesMitigacion

- Columnas: **2**. Filas (snapshot): **4**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| Valores | TEXT(255) | — |
| Explicacion | MEMO | — |

### TbValoresPosiblesPlazoCalidadCosteVulnerabilidad

- Columnas: **1**. Filas (snapshot): **5**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| Valores | TEXT(255) | — |

### TbValoresPosiblesValoracion

- Columnas: **1**. Filas (snapshot): **4**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| Valores | TEXT(255) | — |
