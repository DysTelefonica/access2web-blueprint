# modResponsableFiltroHelper

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/modules/modResponsableFiltroHelper.bas` (248 líneas) · **PRD:** [README](../README.md)

Resuelve el contrato del filtro de responsables de Calidad y Seguridad de `FormExpedientesGestion`. Se usa en `FUNCIONES UTILES.bas:1195-1196` (lee el filtro de los combos) y `:1557` (lo asigna). Las citas son `modResponsableFiltroHelper.bas:línea`.

**Contrato** (cabecera, `:1-60`):
1. Los combos `IDResponsableCalidad` e `IDResponsableSeguridad` tienen dos columnas y `BoundColumn = 1`: su `.Value` es el **ID**; la columna visible es el nombre.
2. La búsqueda se hace por **nombre**, no por ID, porque `TbExpedientesConEntidades` guarda el nombre del responsable y no expone IDs.
3. `getExpBusqueda` lee `Column(1, ListIndex)` (el nombre) como filtro; `setExpBusqueda` busca la fila por nombre y asigna al combo el ID de `Column(0)`, o `0` si el filtro es «Todos» o vacío.

| Función | Regla | Evidencia |
|---|---|---|
| `NormalizarFiltroResponsable` | `""`, `"0"` y `"Todos"` (sin distinguir mayúsculas) ⇒ «Todos»; el resto, recortado | `:62-71` |
| `BuscarIDPorNombre` | Devuelve el ID de la fila cuyo nombre coincide (sin mayúsculas ni espacios) o `"0"` | `:82-126` |
| `FiltroResponsableDeCombo`, `LeerResponsableDeCombo`, `AsignarResponsableACombo`, `SeleccionarComboPorNombre`, `ValorParaAsignarResponsable`, `ValorLeidoDeResponsable`, `FilasDeCombo`, `FiltroResponsableDeFila` | Envoltorios para los combos del formulario | `:127-248` |

## Comportamiento que la migración debe conservar

El filtro por responsable de Calidad o Seguridad con «Todos» como valor neutro, y la correspondencia nombre ⇒ ID. En web el filtro debe ir por **ID**, no por nombre: dos responsables con el mismo nombre hoy se confunden (hipótesis, sin verificar).
