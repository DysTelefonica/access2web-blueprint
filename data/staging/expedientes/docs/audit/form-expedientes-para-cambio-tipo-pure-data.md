# Auditoría pure-data — Form_FormExpedientesParaCambioTipo

## Resumen

- Estado: refactor parcial seguro en fuente para Phase 3.4 / PR41.
- Helper creado: `modExpedientesParaCambioTipoHelper`.
- Tests creados: `Test_ExpedientesParaCambioTipoHelper` con 4 átomos `phase3.4` / `pr41` / `atomic`.

## Separación aplicada

- `Form_Open` obtiene la cabecera de lista desde `ExpedientesParaCambioTipo_ListHeader` y solo pinta el resultado.
- Se cubren reglas pure-data de término de búsqueda obligatorio, selección de lista y validación de tipo `AM`/`Lote`.

## Deuda restante

- `Filtrar`, `ComandoSeleccionar_Click` y `ListaFiltrados_Click` siguen resolviendo objetos `Expediente` en UI. Deben migrarse en una fase dedicada con seam de repositorio o DTO.

## Evidencia fuente

- Sin referencias UI prohibidas en helper/test nuevo.
