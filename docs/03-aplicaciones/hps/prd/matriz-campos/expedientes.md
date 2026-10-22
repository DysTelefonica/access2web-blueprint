# expedientes — expedientes y catalogación de contratación

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Issue:** #1076 · **Matriz:** [índice](../matriz-campos.md)

| Tabla | Columnas | Filas (snapshot) | Ficheros de `src` |
|---|---|---|---|
| Copia de TbExpedienteLugares | 4 | 87 | 0 |
| Copia de TbExpedientes | 7 | 70 | 0 |
| TbJuridicasContratacion | 1 | 3 | 0 |

Leyenda de la columna «Evidencia en `src`»: `F` = la columna aparece como `Fields("columna")`; `S` = el nombre aparece en un fichero que cita la tabla sin `Fields(...)`, coincidencia poco fiable; `—` = sin cita. «Clase de dato» es una heurística por nombre, para priorizar la revisión de privacidad (RGPD). Ninguna fila está aprobada.

## Copia de TbExpedienteLugares

Columnas: 4 · Filas en el snapshot: 87 · Ficheros de `src` que la citan: 0

| Columna | Tipo de origen | Tipo objetivo propuesto | Clase de dato | Evidencia | Fichero |
|---|---|---|---|---|---|
| `IDExpLugar` | LONG(4) | INTEGER | — | — | — |
| `IDExpediente` | LONG(4) | INTEGER | — | — | — |
| `Lugar` | TEXT(255) | VARCHAR | — | — | — |
| `Observaciones` | MEMO | TEXT | — | — | — |

## Copia de TbExpedientes

Columnas: 7 · Filas en el snapshot: 70 · Ficheros de `src` que la citan: 0

| Columna | Tipo de origen | Tipo objetivo propuesto | Clase de dato | Evidencia | Fichero |
|---|---|---|---|---|---|
| `IDExpediente` | LONG(4) | INTEGER | — | — | — |
| `Cod_Exp` | TEXT(255) | VARCHAR | — | — | — |
| `OC` | TEXT(255) | VARCHAR | — | — | — |
| `Grado` | TEXT(255) | VARCHAR | — | — | — |
| `Expediente` | TEXT(255) | VARCHAR | — | — | — |
| `IDJuridicaContrato` | LONG(4) | INTEGER | — | — | — |
| `IDExpedienteNuevo` | LONG(4) | INTEGER | — | — | — |

## TbJuridicasContratacion

Columnas: 1 · Filas en el snapshot: 3 · Ficheros de `src` que la citan: 0

| Columna | Tipo de origen | Tipo objetivo propuesto | Clase de dato | Evidencia | Fichero |
|---|---|---|---|---|---|
| `Juridica` | TEXT(255) | VARCHAR | — | — | — |
