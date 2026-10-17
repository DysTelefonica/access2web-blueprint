# D90 — Reset flow con tokens one-time de 24 h vía adapter de notificación con cola por tabla

- estado: adoptada
- fecha: 2026-08-08
- decide: usuario (dueño del producto)
- motivo: Con los hashes heredados descartados (D89), el usuario migrado necesita un camino explícito para fijar su contraseña.
- origen: `product/lanzadera-reset-flow`

## Decisión

Reset flow con tokens one-time de 24 h vía adapter de notificación con cola por tabla. `issue_reset_token` y `consume_reset_token` atómicos y single-use
