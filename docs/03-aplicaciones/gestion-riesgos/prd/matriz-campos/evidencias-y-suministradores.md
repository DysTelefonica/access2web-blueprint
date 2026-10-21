# Evidencias, anexos y suministradores — matrices campo a campo

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Issue:** #1079 · Fase 3 · **Matriz:** [matriz-campos.md](../matriz-campos.md)

6 tablas. Cada fila está en estado **propuesta**: la topología PostgreSQL global y el ERD de origen siguen abiertos (D-GR-4 del [README](../README.md)).

### TbAnexos

- Columnas: **9**. Filas (snapshot): **484**.
- Evidencia en `src`: **2** fichero(s) de producción — `Anexo.cls`, `Constructor.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDAnexo | LONG(4) | — |
| IDProyecto | LONG(4) | — |
| IDEdicion | LONG(4) | — |
| IDRiesgo | LONG(4) | — |
| FechaAnexo | SHORT_DATE_TIME(8) | — |
| Titulo | TEXT(255) | — |
| NombreArchivo | TEXT(255) | — |
| EvidenciaUTE | TEXT(2) | — |
| EvidenciaSuministrador | TEXT(2) | — |

### TbAnexosAntigua

- Columnas: **9**. Filas (snapshot): **2**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDAnexo | LONG(4) | — |
| IDProyecto | LONG(4) | — |
| IDEdicion | LONG(4) | — |
| FechaAnexo | SHORT_DATE_TIME(8) | — |
| CodigoUnico | TEXT(255) | — |
| Titulo | TEXT(255) | — |
| Descripcion | MEMO | — |
| NombreArchivo | TEXT(255) | — |
| EvidenciaUTE | TEXT(2) | — |

### TbProyectosEdicionesSuministradores

- Columnas: **4**. Filas (snapshot): **20**.
- Evidencia en `src`: **2** fichero(s) de producción — `EdicionSuministrador.cls`, `Constructor.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| ID | LONG(4) | — |
| IDEdicion | LONG(4) | — |
| IDSuministrador | LONG(4) | — |
| IDAnexo | LONG(4) | — |

### TbProyectosEdicionesSuministradores1

- Columnas: **4**. Filas: sin recuento en el esquema disponible.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| ID | N4 | — |
| IDEdicion | N4 | — |
| Suministrador | T10 | — |
| IDAnexo | N4 | — |

### TbProyectosSuministradores

- Columnas: **4**. Filas (snapshot): **0**.
- Evidencia en `src`: **2** fichero(s) de producción — `ProyectoSuministrador.cls`, `Constructor.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| ID | LONG(4) | — |
| IDProyecto | LONG(4) | — |
| IDSuministrador | LONG(4) | — |
| GestionCalidad | TEXT(2) | personal |

### TbProyectosSuministradores1

- Columnas: **4**. Filas: sin recuento en el esquema disponible.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| ID | N4 | — |
| IDProyecto | N4 | — |
| Suministrador | T10 | — |
| GestionCalidad | T10 | personal |
