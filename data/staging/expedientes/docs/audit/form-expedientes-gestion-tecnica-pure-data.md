# Auditoría pure-data — Form_FormExpedientesGestionTecnica

## Resumen

- Estado: refactor parcial seguro en fuente para Phase 3.4 / PR41.
- Helper creado: `modExpedientesGestionTecnicaHelper`.
- Tests creados: `Test_ExpedientesGestionTecnicaHelper` con 4 átomos `phase3.4` / `pr41` / `atomic`.

## Separación aplicada

- `Filtrar` delega la cabecera/estado inicial de la lista en `ExpedientesGestionTecnica_ListHeader` y conserva únicamente el renderizado en el formulario.
- Se añaden decisiones pure-data para reinicio de campos, refresco por firma de búsqueda y estado de selección.

## Deuda restante

- La carga de expedientes técnicos y la apertura de detalle siguen acopladas a constructor/globales. Se dejan fuera para evitar una PR no revisable.

## Evidencia fuente

- Sin referencias UI prohibidas en helper/test nuevo.
