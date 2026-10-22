# D78 — Versionado semántico por módulo Y plataforma base (`modulo/vX.Y.Z-rc.n` / `modulo/vX.Y.Z`…

- estado: adoptada
- fecha: 2026-08-05
- decide: usuario (dueño del producto)
- motivo: Los despliegues coexistentes cuestan complejidad operativa (un enrutador por usuario, bases de datos separadas, diferenciación visual clara) incompatible con el principio de simplicidad para 2 desarrolladores. Expand and Contract es la estrategia correcta para cambios de base de datos sin parada.
- origen: `external-prompt-review/section-5-resolution`

## Decisión

Versionado semántico por módulo Y plataforma base (`modulo/vX.Y.Z-rc.n` / `modulo/vX.Y.Z` y `platform/vX.Y.Z`)
