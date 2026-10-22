# ExpedienteCPV

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/classes/ExpedienteCPV.cls` (148 líneas) · **PRD:** [README](../README.md)

Objeto de datos de una fila de `TbExpedientesCPVs`: relaciona un expediente con un código CPV del catálogo ([Form_FormCPV](../formularios/Form_FormCPV.md)). Igual que [ExpedientePECAL](ExpedientePECAL.md): sin lógica de negocio, `ColCampos` con `IDCPVExpediente`, `IDCPV` e `IDExpediente` (`:82-84`), acceso por nombre y siguiente identificador libre con `DameID("TbExpedientesCPVs", "IDCPVExpediente", getdb())` (`:56`), que devuelve `Max + 1` o `1` con la tabla vacía (`FUNCIONES UTILES.bas:2851-2864`). Las citas son `ExpedienteCPV.cls:línea` salvo indicación.

## Dónde se usa

- **Lectura:** `constructor.getExpedienteCPV` (`constructor.bas:4741-4770`), con `SELECT * FROM TbExpedientesCPVs` (`:4754`).
- **Búsqueda por CPV:** `INNER JOIN` con `TbExpedientes` filtrado por `IDCPV` (`constructor.bas:2734-2735`).
- **Copia entre expedientes:** `CopiarCPVs` (`FUNCIONES UTILES.bas:3473-3534`), dentro de [Copiar y extender entidades](../modulos/funciones-utiles-copiar-entidades.md).

## Comportamiento que la migración debe conservar

La relación expediente-CPV con identificador propio y el filtrado de expedientes por CPV. El identificador `Max + 1` no está protegido contra escrituras concurrentes que se vean en el código (hipótesis); en web hace falta una secuencia o una restricción de unicidad.

## Pendientes

- Confirmar los tipos de columna con un volcado del ERD de RC1; las columnas salen del código.
