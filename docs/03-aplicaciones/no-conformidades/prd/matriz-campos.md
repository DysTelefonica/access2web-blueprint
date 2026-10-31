# Matriz campo a campo: No Conformidades

> **Estado:** borrador para revisión · parte de la cadena del PRD de [#1080](README.md). Línea base: `main` del repositorio legacy en el commit `1d7a2f826a1e3c582874e4e29886a6183efb9067`.

Una fila por columna del esquema previo ([ERD/schema.sql](../ERD/schema.sql), volcado del backend `NoConformidades_Datos.accdb` del snapshot de 2026-08-06) con su destino propuesto en PostgreSQL. El esquema previo solo trae nombres y tipos: **no hay claves, índices, relaciones ni nulabilidad declaradas**. Las claves y relaciones que aparecen en las partes son *inferidas por código* (con cita) hasta que el propietario extraiga el ERD de origen (véase [erd-origen.md](erd-origen.md)).

## Reglas comunes

- **Destino:** `no_conformidades.<tabla>.<columna>` es una propuesta. No se fija hasta aprobar el ERD global; no se genera ni ejecuta DDL.
- **Estado:** `propuesta` (la tabla tiene uso en `src` de `main` y la columna un destino candidato) o `pendiente` con motivo. «Pendiente: sin uso en main» significa que una búsqueda literal sobre `src/forms`, `src/modules`, `src/classes` y `src/queries` no encuentra la tabla; la decisión de migrarla o descartarla es del propietario.
- **Transacciones:** `src` no contiene `BeginTrans`; las escrituras multi-paso no son atómicas (véase [coexistencia.md](coexistencia.md)). Los identificadores se calculan con máximo + 1 en cliente.
- **Datos:** no se reproducen datos de negocio ni recuentos de filas.

## Partes

| Parte | Tablas | Filas | Propuesta | Pendiente | A verificar |
|---|---|---|---|---|---|
| [a-anexos-auditoria](matriz-campos/a-anexos-auditoria.md) | 7 | 36 | 9 | 27 | 0 |
| [b-cache-indicadores](matriz-campos/b-cache-indicadores.md) | 6 | 100 | 0 | 100 | 0 |
| [c-configuracion-logs](matriz-campos/c-configuracion-logs.md) | 11 | 89 | 42 | 44 | 3 |
| [d-acciones](matriz-campos/d-acciones.md) | 8 | 68 | 50 | 18 | 0 |
| [e-no-conformidades](matriz-campos/e-no-conformidades.md) | 2 | 72 | 62 | 10 | 0 |
| [f-lotes-replanificaciones](matriz-campos/f-lotes-replanificaciones.md) | 8 | 66 | 18 | 48 | 0 |
| **Total** | **42** | **431** | **181** | **247** | **3** |

Las 428 columnas del DDL tienen fila. Las 3 filas «a verificar» son columnas que `InicializadorCache.bas` usa y que no figuran en el DDL (`FechaUltimoUso`, `HitsConsultas`, `TamanioBytes` de la caché por NC), un indicio de desfase entre el volcado y el backend actual.

## Hallazgos transversales

- 247 de 431 filas quedan **pendientes**, sobre todo por tablas sin uso en `main`: la caché de indicadores y de listados (retirada frente al estudio previo, véase [Objetos retirados](README.md#objetos-retirados-frente-al-estudio-previo)), los lotes de ingreso y varias tablas auxiliares. Son decisión del propietario.
- `TbNoConformidades` y `TbNoConformidadesAuditoria` concentran las columnas con uso demostrado; varias columnas de ambas no se leen ni escriben en `main`.
- Cuatro uniones de `constructor.bas` enlazan `TbAuditorias.IDAuditoria` con `TbNoConformidadesAuditoria.ID` en lugar de `IDAuditoria`; queda «a verificar».
- El volcado declara 49 tablas y el DDL contiene 42 (D-NC-4 en el [README](README.md#discrepancias-abiertas)).
