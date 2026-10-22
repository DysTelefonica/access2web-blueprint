# D69 — Actualización de contadores pendientes vía polling HTMX (`hx-trigger="every 30s"`) +…

- estado: adoptada
- fecha: 2026-08-05
- decide: usuario (dueño del producto)
- motivo: El requisito real del usuario es «contador de campana fiable, que se pueda consultar periódicamente, con botón de refresco manual», no propagación de eventos en tiempo real.
- origen: `external-prompt-review/section-1-resolution`

## Decisión

Actualización de contadores pendientes vía polling HTMX (`hx-trigger="every 30s"`) + botón de refresh manual. Sin SSE, sin WebSockets
