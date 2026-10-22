# D86 — La forma hexagonal del legacy (dominio en clases, helpers por dominio, transaccionalidad…

- estado: adoptada
- fecha: 2026-08-05
- decide: usuario (dueño del producto)
- motivo: El blueprint hexagonal de la plataforma nueva (D8) no es un invento: el legado ya tenía esa forma. La correspondencia es uno a uno y preserva la intención de negocio.
- origen: `blueprint/staging-expedientes-hexagonal-shape`

## Decisión

La forma hexagonal del legacy (dominio en clases, helpers por dominio, transaccionalidad en `ExpedienteOperaciones`, config detrás de `getdb()` + `TbConfiguracionBackends`) **se preserva como referencia** para el mapeo a la nueva plataforma. El blueprint hexagonal NO es invención nueva: el legacy ya tenía esta forma
