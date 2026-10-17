# D68 — Estructura del repositorio: monorepo `access2web-blueprint/` con monolito modular…

- estado: adoptada
- fecha: 2026-08-05
- decide: usuario (dueño del producto)
- motivo: La sección 2.3 cierra la pregunta P8 del blueprint (el stack estaba ABIERTO). HTMX con renderizado en servidor es la elección más coherente con la sección 1.3 (simplicidad para 2 desarrolladores) y evita la complejidad de una SPA. La arquitectura hexagonal ya estaba aprobada (D8).
- origen: `external-prompt-review/section-2-resolution`

## Decisión

Estructura del repositorio: monorepo `access2web-blueprint/` con monolito modular, límites por paquete y ports por módulo
