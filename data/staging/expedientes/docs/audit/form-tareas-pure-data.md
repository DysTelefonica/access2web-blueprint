# Auditoría pure-data — Form_FormTareas

## Resumen

- Estado: refactor parcial seguro en fuente para Phase 3.4 / PR41.
- Helper creado: `modTareasHelper`.
- Tests creados: `Test_TareasHelper` con 4 átomos `phase3.4` / `pr41` / `atomic`.

## Separación aplicada

- `Filtrar` delega la cabecera y el texto de lista vacía en `Tareas_ListHeader`.
- Se cubren reglas pure-data de tipo de tarea válido, acción de informe y habilitación de acciones al seleccionar expediente.

## Deuda restante

- La carga desde `m_ObjEntorno` y la navegación al detalle/informe siguen en el formulario. Deben partirse por rutas en una PR posterior.

## Evidencia fuente

- Sin referencias UI prohibidas en helper/test nuevo.
