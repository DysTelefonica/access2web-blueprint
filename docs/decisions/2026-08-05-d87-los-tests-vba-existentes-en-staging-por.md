# D87 — Los tests VBA existentes en staging (`Test_*` por cada `Helper_*` + `tests.vba.json` +…

- estado: adoptada
- fecha: 2026-08-05
- decide: usuario (dueño del producto)
- motivo: El blueprint hexagonal de la plataforma nueva (D8) no es un invento: el legado ya tenía esa forma. La correspondencia es uno a uno y preserva la intención de negocio.
- origen: `blueprint/staging-expedientes-hexagonal-shape`

## Decisión

Los tests VBA existentes en staging (`Test_*` por cada `Helper_*` + `tests.vba.json` + `tests.vba.responsable-71.json`) son **evidencia de comportamiento** que se preserva como referencia para los nuevos tests pytest. Ningún `Test_*` se descarta sin trazabilidad
