# Publicación, control de cambios y tareas — matrices campo a campo

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Issue:** #1079 · Fase 3 · **Matriz:** [matriz-campos.md](../matriz-campos.md)

9 tablas. Cada fila está en estado **propuesta**: la topología PostgreSQL global y el ERD de origen siguen abiertos (D-GR-4 del [README](../README.md)).

### TbCambiosExplicacion

- Columnas: **2**. Filas (snapshot): **5**.
- Evidencia en `src`: **1** fichero(s) de producción — `Constructor.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| Apartado | TEXT(255) | — |
| Explicacion | MEMO | — |

### TbExplicacionCarencias

- Columnas: **3**. Filas (snapshot): **6**.
- Evidencia en `src`: **2** fichero(s) de producción — `CarenciasExplicacion.cls`, `Constructor.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| ID | LONG(4) | — |
| Apartado | TEXT(255) | — |
| Explicacion | MEMO | — |

### TbLog

- Columnas: **12**. Filas (snapshot): **5817**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDLog | LONG(4) | — |
| IDProyecto | LONG(4) | — |
| IDEdicion | LONG(4) | — |
| IDRiesgo | LONG(4) | — |
| IDMitigacion | LONG(4) | — |
| IDContingencia | LONG(4) | — |
| IDAccionMitigacion | LONG(4) | — |
| IDAccionContingencia | LONG(4) | — |
| Usuario | TEXT(255) | personal |
| Fecha | SHORT_DATE_TIME(8) | — |
| Titulo | TEXT(255) | — |
| Linea | MEMO | — |

### TbLogPublicaciones

- Columnas: **11**. Filas (snapshot): **329**.
- Evidencia en `src`: **2** fichero(s) de producción — `PublicacionLog.cls`, `Constructor.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| ID | LONG(4) | — |
| IDEdicion | LONG(4) | — |
| FechaPreparadaParaPublicar | SHORT_DATE_TIME(8) | — |
| UsuarioFechaPreparadaParaPublicar | TEXT(255) | personal |
| FechaPreparadaParaPublicarQuitar | SHORT_DATE_TIME(8) | — |
| UsuarioFechaPreparadaParaPublicarQuitar | TEXT(255) | personal |
| FechaPublicacion | SHORT_DATE_TIME(8) | — |
| UsuarioFechaPublicacion | TEXT(255) | personal |
| PropuestaRechazadaPorCalidadFecha | SHORT_DATE_TIME(8) | personal |
| PropuestaRechazadaPorCalidadMotivo | MEMO | personal |
| FechaRegistro | SHORT_DATE_TIME(8) | — |

### TbTareas

- Columnas: **11**. Filas (snapshot): **36**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDtarea | LONG(4) | — |
| IDProyecto | LONG(4) | — |
| IDRiesgo | LONG(4) | — |
| CodRiesgo | TEXT(255) | — |
| TipoTarea | TEXT(255) | — |
| EstadoTarea | TEXT(255) | — |
| FechaAccion | SHORT_DATE_TIME(8) | — |
| FechaJustificacion | SHORT_DATE_TIME(8) | — |
| Justificacion | MEMO | — |
| FechaVisado | SHORT_DATE_TIME(8) | — |
| FechaRechazo | SHORT_DATE_TIME(8) | — |

### TbTareasExplicaciones

- Columnas: **3**. Filas (snapshot): **10**.
- Evidencia en `src`: **1** fichero(s) de producción — `Constructor.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| NodoTarea | TEXT(255) | — |
| TituloTarea | TEXT(255) | — |
| Explicacion | MEMO | — |

### tbCambios

- Columnas: **8**. Filas (snapshot): **6725**.
- Evidencia en `src`: **2** fichero(s) de producción — `Cambio.cls`, `Constructor.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDCambio | LONG(4) | — |
| IDProyecto | LONG(4) | — |
| EdicionInicial | INT(2) | — |
| EdicionFinal | INT(2) | — |
| Riesgo | TEXT(255) | — |
| NombreCampo | TEXT(255) | — |
| Descripcion | MEMO | — |
| FechaRegistro | SHORT_DATE_TIME(8) | — |

### tbCambios

- Columnas: **8**. Filas (snapshot): **6725**.
- Evidencia en `src`: **2** fichero(s) de producción — `Cambio.cls`, `Constructor.bas`

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDCambio | LONG(4) | — |
| IDProyecto | LONG(4) | — |
| EdicionInicial | INT(2) | — |
| EdicionFinal | INT(2) | — |
| Riesgo | TEXT(255) | — |
| NombreCampo | TEXT(255) | — |
| Descripcion | MEMO | — |
| FechaRegistro | SHORT_DATE_TIME(8) | — |

### tbCambiosParaPublicacion

- Columnas: **8**. Filas (snapshot): **68**.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
| IDCambio | LONG(4) | — |
| IDProyecto | LONG(4) | — |
| EdicionInicial | INT(2) | — |
| EdicionFinal | INT(2) | — |
| Riesgo | TEXT(255) | — |
| NombreCampo | TEXT(255) | — |
| Descripcion | MEMO | — |
| FechaRegistro | SHORT_DATE_TIME(8) | — |
