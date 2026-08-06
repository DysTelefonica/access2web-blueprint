# Auditoría pure-data — Form_FormExpedientesGestion

## Resumen

- Estado: refactor parcial seguro en fuente para Phase 3.4 / PR41.
- Helper creado: `modExpedientesGestionHelper`.
- Tests creados: `Test_ExpedientesGestionHelper` con 4 átomos `phase3.4` / `pr41` / `atomic`.

## Separación aplicada

- `EstablecerDatos` ya no decide en línea la visibilidad/texto de `lblUltimaModificacion`; delega esa decisión en `ExpedientesGestion_FormLoadState` y el formulario solo renderiza el payload.
- Se documentan helpers pure-data adicionales para reinicio de campos, estado de selección y cabeceras de listas.

## Deuda restante

- El formulario sigue siendo de alta complejidad y conserva handlers con apertura de formularios, listas y operaciones DAO. Esta PR evita mover esas rutas por el bloqueo global de verificación binaria.
- Siguiente fase recomendada: separar por subfamilias (`alta`, `cambio tipo`, `detalle`, `exportación`) en PRs menores.

## Evidencia fuente

- Sin `Forms(...)`, `Screen.ActiveForm`, `DoCmd.OpenForm`, `MsgBox`, `InputBox` ni `Debug.Print` en el helper/test nuevo.
