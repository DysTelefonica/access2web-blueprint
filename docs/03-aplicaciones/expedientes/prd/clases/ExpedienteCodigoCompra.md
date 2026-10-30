# ExpedienteCodigoCompra

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/classes/ExpedienteCodigoCompra.cls` (122 líneas) · **PRD:** [README](../README.md)

Objeto de datos de una fila de `TbExpedientesCodigoCompras`: relaciona un expediente con un elemento de catálogo. Sin lógica de negocio propia: campos públicos, `ColCampos` (los nombres de columna), acceso por nombre (`getPropiedad`, `SetPropiedad`; un nombre desconocido deja «Propiedad no reconocida en ExpedienteCodigoCompra.getPropiedad» y lanza el error interno 1000) y el siguiente identificador libre. Las citas son `ExpedienteCodigoCompra.cls:línea` salvo indicación.

## Miembros

| Miembro | Tipo | Significado |
|---|---|---|
| `id` | `String` | Identificador de la fila (`:14`) |
| `IDExpediente` | `String` | Expediente (`:15`) |
| `CodCompras` | `String` | Código de compra (`:16`) |
| `IDCalculado` | `Property Get` | Siguiente identificador libre (`:23-42`) |

`IDCalculado` llama a `DameID("TbExpedientesCodigoCompras", "ID", getdb())` en cada lectura, que devuelve `Max + 1` o `1` con la tabla vacía (`FUNCIONES UTILES.bas:2851-2864`); no reserva el valor.

## Dónde se usa

- **Lectura:** `constructor.getExpedienteCodigoCompra y getExpedienteCodigoCompras` (`constructor.bas:4544-4582, 4584`): `SELECT *` sobre `TbExpedientesCodigoCompras` filtrado por identificador concatenado, y copia cada columna de `ColCampos` con `Nz(…, "")`. Sin filas devuelve `Nothing`.
- **Copia entre expedientes:** `—` (`no hay `Copiar*` para esta relación en `FUNCIONES UTILES.bas`: los códigos de compra no se heredan al extender entidades`), dentro de [Copiar y extender entidades](../modulos/funciones-utiles-copiar-entidades.md).

## Comportamiento que la migración debe conservar

La relación expediente-elemento con identificador propio y el alta por `INSERT` con identificador calculado en el momento. El identificador `Max + 1` no está protegido contra escrituras concurrentes que se vean en el código (hipótesis); en web hace falta una secuencia o una restricción de unicidad.

## Pendientes

- Confirmar el tipo de cada columna con un volcado del ERD de RC1; las columnas salen del código, no del DDL.
