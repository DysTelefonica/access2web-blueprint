# ExpedienteBusquedaTecnica

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/classes/ExpedienteBusquedaTecnica.cls` (27 líneas) · **PRD:** [README](../README.md)

Parámetros de la búsqueda técnica, sin lógica: cinco campos públicos de texto y `Error` (`:14-22`).

| Campo | Qué guarda |
|---|---|
| `PalabraClave`, `ESTADO`, `JURIDICA`, `CodExp`, `jp` | El valor de cada filtro de [`FormExpedientesGestionTecnica`](../formularios/Form_FormExpedientesGestionTecnica.md) |

La instancia activa vive en la global `m_ObjExpBusquedaTecnicaActiva` (`Variables Globales.bas:270`). Se crea en `FUNCIONES UTILES.bas:1272-1276`, la reciben las funciones de `:785`, `:808`, `:835` y `:1072`, y `FormExpedientesGestionTecnica` la rellena con `setExpBusquedaTecnica` (`Form_FormExpedientesGestionTecnica.cls:22`).

## Comportamiento que la migración debe conservar

Los cinco filtros de la búsqueda técnica y su combinación. Son parámetros de una consulta: en web viajan como parámetros de la petición, no como objeto global.

## Pendientes

- Semántica exacta de cada filtro en `getWhereBusqueda`: se describe en la ficha del formulario.
