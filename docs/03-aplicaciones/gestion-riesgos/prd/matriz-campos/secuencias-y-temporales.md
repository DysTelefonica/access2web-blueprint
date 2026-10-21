# Secuencias de identificadores y temporales — matrices campo a campo

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Issue:** #1079 · Fase 3 · **Matriz:** [matriz-campos.md](../matriz-campos.md)

16 tablas. Cada fila está en estado **propuesta**: la topología PostgreSQL global y el ERD de origen siguen abiertos (D-GR-4 del [README](../README.md)).

### TbIDAccionPlanContingencia

- Columnas: **5**. Filas (snapshot): **251**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDAccionContingencia | LONG(4) | — |
| UsuarioCrea | TEXT(255) | personal |
| FechaCreacion | SHORT_DATE_TIME(8) | — |
| FechaEliminacion | SHORT_DATE_TIME(8) | — |
| UsuarioElimina | TEXT(255) | personal |

### TbIDAccionPlanMitigacion

- Columnas: **5**. Filas (snapshot): **1074**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDAccionMitigacion | LONG(4) | — |
| UsuarioCrea | TEXT(255) | personal |
| FechaCreacion | SHORT_DATE_TIME(8) | — |
| FechaEliminacion | SHORT_DATE_TIME(8) | — |
| UsuarioElimina | TEXT(255) | personal |

### TbIDAnexos

- Columnas: **5**. Filas (snapshot): **107**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDAnexo | LONG(4) | — |
| UsuarioCrea | TEXT(255) | personal |
| FechaCreacion | SHORT_DATE_TIME(8) | — |
| FechaEliminacion | SHORT_DATE_TIME(8) | — |
| UsuarioElimina | TEXT(255) | personal |

### TbIDBibliotecaRiesgos

- Columnas: **5**. Filas (snapshot): **51**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDRiesgoTipo | LONG(4) | — |
| UsuarioCrea | TEXT(255) | personal |
| FechaCreacion | SHORT_DATE_TIME(8) | — |
| FechaEliminacion | SHORT_DATE_TIME(8) | — |
| UsuarioElimina | TEXT(255) | personal |

### TbIDCorreosEnviados

- Columnas: **5**. Filas (snapshot): **264**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDCorreo | LONG(4) | personal |
| UsuarioCrea | TEXT(255) | personal |
| FechaCreacion | SHORT_DATE_TIME(8) | — |
| FechaEliminacion | SHORT_DATE_TIME(8) | — |
| UsuarioElimina | TEXT(255) | personal |

### TbIDDocumentos

- Columnas: **5**. Filas (snapshot): **27**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDDocumento | LONG(4) | — |
| UsuarioCrea | TEXT(255) | personal |
| FechaCreacion | SHORT_DATE_TIME(8) | — |
| FechaEliminacion | SHORT_DATE_TIME(8) | — |
| UsuarioElimina | TEXT(255) | personal |

### TbIDEdiciones

- Columnas: **5**. Filas (snapshot): **209**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDEdicion | LONG(4) | — |
| UsuarioCrea | TEXT(255) | personal |
| FechaCreacion | SHORT_DATE_TIME(8) | — |
| FechaEliminacion | SHORT_DATE_TIME(8) | — |
| UsuarioElimina | TEXT(255) | personal |

### TbIDLog

- Columnas: **5**. Filas (snapshot): **1077**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDLog | LONG(4) | — |
| UsuarioCrea | TEXT(255) | personal |
| FechaCreacion | SHORT_DATE_TIME(8) | — |
| FechaEliminacion | SHORT_DATE_TIME(8) | — |
| UsuarioElimina | TEXT(255) | personal |

### TbIDPlanContingencia

- Columnas: **5**. Filas (snapshot): **154**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDContingencia | LONG(4) | — |
| UsuarioCrea | TEXT(255) | personal |
| FechaCreacion | SHORT_DATE_TIME(8) | — |
| FechaEliminacion | SHORT_DATE_TIME(8) | — |
| UsuarioElimina | TEXT(255) | personal |

### TbIDPlanMitigacion

- Columnas: **5**. Filas (snapshot): **640**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDMitigacion | LONG(4) | — |
| UsuarioCrea | TEXT(255) | personal |
| FechaCreacion | SHORT_DATE_TIME(8) | — |
| FechaEliminacion | SHORT_DATE_TIME(8) | — |
| UsuarioElimina | TEXT(255) | personal |

### TbIDProyectos

- Columnas: **5**. Filas (snapshot): **71**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDProyecto | LONG(4) | — |
| UsuarioCrea | TEXT(255) | personal |
| FechaCreacion | SHORT_DATE_TIME(8) | — |
| FechaEliminacion | SHORT_DATE_TIME(8) | — |
| UsuarioElimina | TEXT(255) | personal |

### TbIDRiesgos

- Columnas: **5**. Filas (snapshot): **889**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDRiesgo | LONG(4) | — |
| UsuarioCrea | TEXT(255) | personal |
| FechaCreacion | SHORT_DATE_TIME(8) | — |
| FechaEliminacion | SHORT_DATE_TIME(8) | — |
| UsuarioElimina | TEXT(255) | personal |

### TbIDRiesgosAIntegrar

- Columnas: **5**. Filas (snapshot): **58**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDRiesgoExt | LONG(4) | — |
| UsuarioCrea | TEXT(255) | personal |
| FechaCreacion | SHORT_DATE_TIME(8) | — |
| FechaEliminacion | SHORT_DATE_TIME(8) | — |
| UsuarioElimina | TEXT(255) | personal |

### TbIDTareas

- Columnas: **5**. Filas (snapshot): **53**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDtarea | LONG(4) | — |
| UsuarioCrea | TEXT(255) | personal |
| FechaCreacion | SHORT_DATE_TIME(8) | — |
| FechaEliminacion | SHORT_DATE_TIME(8) | — |
| UsuarioElimina | TEXT(255) | personal |

### ~TMPCLP576781

- Columnas: **4**. Filas (snapshot): **0**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| ID | LONG(4) | — |
| IDEdicion | LONG(4) | — |
| Suministrador | TEXT(255) | — |
| IDAnexo | LONG(4) | — |

### ~TMPCLP603661

- Columnas: **4**. Filas (snapshot): **0**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| ID | LONG(4) | — |
| IDProyecto | LONG(4) | — |
| Suministrador | TEXT(255) | — |
| GestionCalidad | TEXT(2) | personal |
