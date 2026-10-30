# E2EFamilyExportHelper

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/modules/E2EFamilyExportHelper.bas` (346 líneas) · **PRD:** [README](../README.md)

Prepara una exportación E2E por familias. Las citas son `E2EFamilyExportHelper.bas:línea`. Lo usa `ExpedienteJsonExporter.bas` (`:323`, `:330`, `:378`, `:385`) y se cita desde `E2EEnlacesService.bas:21`; es código vivo.

| Función | Qué hace | Evidencia |
|---|---|---|
| `E2EFamilyExport_ExpandCsvToFamilyCsv(db, ids)` | Para cada ID seleccionado busca su expediente raíz (`FindRootExpedienteId`) y recorre sus descendientes (`ExpandDescendantsFromRoot`); devuelve el CSV de IDs sin repetir. Falla con «lista de IDs vacía o inválida» o «ningún ID corresponde a un expediente existente» | `:13-58`, `:164-266` |
| `E2EFamilyExport_AssignMissingOrdinals(db, ids)` | A los expedientes de la lista con `OrdinalE2E` nulo les asigna números correlativos desde `Max(OrdinalE2E) + 1`, en orden de `IDExpediente`; cada `UPDATE` va con parámetros | `:60-131` |
| `E2EFamilyExport_SetOrdinalPreparationFailureForTest` | Interruptor solo de pruebas que fuerza un fallo al asignar | `:9-11` |

Exportar un expediente exporta **toda su familia** (raíz y descendientes), no solo el seleccionado, y numera a los que aún no tienen ordinal. Capacidades relacionadas: [CAP-035](../../capabilities/CAP-035-batch-e2e.md) y [CAP-042](../../capabilities/CAP-042-ordinal-e2e.md).

## Posibles defectos

- La asignación recorre los expedientes uno a uno (`:103-115`) y el helper no abre transacción, pero `ExpedienteJsonExporter.bas` la llama dentro de `BeginTrans` (`:327-330`): un fallo a mitad se deshace entero. Quien llame al helper sin ese marco pierde la atomicidad.
- `Max(OrdinalE2E) + 1` leído antes del bucle (`:87`): sin un bloqueo visible, dos exportaciones simultáneas podrían asignar el mismo ordinal (hipótesis).

## Comportamiento que la migración debe conservar

Expansión a la familia completa y ordinales correlativos estables, asignados una sola vez y dentro de una transacción. En web, además, con una restricción de unicidad.
