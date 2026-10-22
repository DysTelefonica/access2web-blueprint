# hps — los cuatro tipos de HPS

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Issue:** #1076 · **Matriz:** [índice](../matriz-campos.md)

| Tabla | Columnas | Filas (snapshot) | Ficheros de `src` |
|---|---|---|---|
| TbHPS | 5 | 1360 | 8 |
| TbHPSEquivalencia | 4 | 3 | 0 |
| TbHPSGrado | 2 | 13 | 1 |
| TbMotivoHPS | 1 | 58 | 2 |

Leyenda de la columna «Evidencia en `src`»: `F` = la columna aparece como `Fields("columna")`; `S` = el nombre aparece en un fichero que cita la tabla sin `Fields(...)`, coincidencia poco fiable; `—` = sin cita. «Clase de dato» es una heurística por nombre, para priorizar la revisión de privacidad (RGPD). Ninguna fila está aprobada.

## TbHPS

Columnas: 5 · Filas en el snapshot: 1360 · Ficheros de `src` que la citan: 8

| Columna | Tipo de origen | Tipo objetivo propuesto | Clase de dato | Evidencia | Fichero |
|---|---|---|---|---|---|
| `IDUsuario` | LONG(4) | INTEGER | personal | F | `classes/HPS.cls` |
| `TipoHPS` | TEXT(255) | VARCHAR | — | F | `classes/HPS.cls` |
| `Grado` | TEXT(255) | VARCHAR | — | F | `classes/HPS.cls` |
| `Renovacion` | TEXT(2) | VARCHAR | — | F | `classes/HPS.cls` |
| `Especialidad` | TEXT(255) | VARCHAR | — | F | `classes/HPS.cls` |

## TbHPSEquivalencia

Columnas: 4 · Filas en el snapshot: 3 · Ficheros de `src` que la citan: 0

| Columna | Tipo de origen | Tipo objetivo propuesto | Clase de dato | Evidencia | Fichero |
|---|---|---|---|---|---|
| `GradoNacional` | TEXT(255) | VARCHAR | — | — | — |
| `GradoOTAN` | TEXT(255) | VARCHAR | — | — | — |
| `GradoESA` | TEXT(255) | VARCHAR | — | — | — |
| `GradoUE` | TEXT(255) | VARCHAR | — | — | — |

## TbHPSGrado

Columnas: 2 · Filas en el snapshot: 13 · Ficheros de `src` que la citan: 1

| Columna | Tipo de origen | Tipo objetivo propuesto | Clase de dato | Evidencia | Fichero |
|---|---|---|---|---|---|
| `TipoHPS` | TEXT(255) | VARCHAR | — | F | `modules/Constructor.bas` |
| `Grado` | TEXT(255) | VARCHAR | — | F | `modules/Constructor.bas` |

## TbMotivoHPS

Columnas: 1 · Filas en el snapshot: 58 · Ficheros de `src` que la citan: 2

| Columna | Tipo de origen | Tipo objetivo propuesto | Clase de dato | Evidencia | Fichero |
|---|---|---|---|---|---|
| `MotivoHPS` | TEXT(255) | VARCHAR | — | F | `modules/Constructor.bas` |
