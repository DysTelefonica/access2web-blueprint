# catalogo — catálogo y soporte

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Issue:** #1076 · **Matriz:** [índice](../matriz-campos.md)

| Tabla | Columnas | Filas (snapshot) | Ficheros de `src` |
|---|---|---|---|
| TbAuxCursos | 1 | 90 | 0 |
| TbConsultas | 2 | 8 | 1 |

Leyenda de la columna «Evidencia en `src`»: `F` = la columna aparece como `Fields("columna")`; `S` = el nombre aparece en un fichero que cita la tabla sin `Fields(...)`, coincidencia poco fiable; `—` = sin cita. «Clase de dato» es una heurística por nombre, para priorizar la revisión de privacidad (RGPD). Ninguna fila está aprobada.

## TbAuxCursos

Columnas: 1 · Filas en el snapshot: 90 · Ficheros de `src` que la citan: 0

| Columna | Tipo de origen | Tipo objetivo propuesto | Clase de dato | Evidencia | Fichero |
|---|---|---|---|---|---|
| `ID` | LONG(4) | INTEGER | — | — | — |

## TbConsultas

Columnas: 2 · Filas en el snapshot: 8 · Ficheros de `src` que la citan: 1

| Columna | Tipo de origen | Tipo objetivo propuesto | Clase de dato | Evidencia | Fichero |
|---|---|---|---|---|---|
| `IDConsulta` | LONG(4) | INTEGER | — | S | `modules/Constructor.bas` |
| `Nombre` | TEXT(255) | VARCHAR | personal | F | `modules/Constructor.bas` |
