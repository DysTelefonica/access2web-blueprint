# ExpedienteRAC

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/classes/ExpedienteRAC.cls` (148 líneas) · **PRD:** [README](../README.md)

Objeto de datos de una fila de `TbExpedientesRACS`: relaciona un expediente con un elemento de catálogo. Sin lógica de negocio propia: campos públicos, `ColCampos` (los nombres de columna), acceso por nombre (`getPropiedad`, `SetPropiedad`; un nombre desconocido deja «Propiedad no reconocida en ExpedienteRAC.getPropiedad» y lanza el error interno 1000) y el siguiente identificador libre. Las citas son `ExpedienteRAC.cls:línea` salvo indicación.

## Miembros

| Miembro | Tipo | Significado |
|---|---|---|
| `IDRacExpediente` | `String` | Identificador de la relación |
| `IDExpediente` | `String` | Expediente |
| `IDRAC` | `String` | Elemento RAC del catálogo |
| `IDRacExpedienteCalculado` | `Property Get` | Siguiente identificador libre (`:55`) |

`IDRacExpedienteCalculado` llama a `DameID("TbExpedientesRACS", "IDRacExpediente", getdb())` en cada lectura, que devuelve `Max + 1` o `1` con la tabla vacía (`FUNCIONES UTILES.bas:2851-2864`); no reserva el valor.

## Dónde se usa

- **Lectura:** `constructor.getExpedienteRAC` (`constructor.bas:5133`): `SELECT *` sobre `TbExpedientesRACS` filtrado por identificador concatenado, y copia cada columna de `ColCampos` con `Nz(…, "")`. Sin filas devuelve `Nothing`.
- **Copia entre expedientes:** `CopiarRACs` (`FUNCIONES UTILES.bas:3663-3724`), dentro de [Copiar y extender entidades](../modulos/funciones-utiles-copiar-entidades.md).

## Comportamiento que la migración debe conservar

La relación expediente-elemento con identificador propio y el alta por `INSERT` con identificador calculado en el momento. El identificador `Max + 1` no está protegido contra escrituras concurrentes que se vean en el código (hipótesis); en web hace falta una secuencia o una restricción de unicidad.

## Pendientes

- Confirmar el tipo de cada columna con un volcado del ERD de RC1; las columnas salen del código, no del DDL.
