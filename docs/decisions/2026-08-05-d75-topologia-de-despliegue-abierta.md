# D75 — Topología de despliegue ABIERTA (consistente con P7)

- estado: adoptada
- fecha: 2026-08-05
- decide: usuario (dueño del producto)
- motivo: La sección 4 original introducía InsForge (BaaS) como backend, lo que rompe la propiedad hexagonal. También proponía OpenShift + GCP para 200 usuarios concurrentes, un dimensionamiento diez veces excesivo. Las notas sobre persistencia de SSE están obsoletas porque SSE se rechazó en la sección 1.
- origen: `external-prompt-review/section-4-resolution`

## Decisión

Topología de despliegue ABIERTA (consistente con P7). Decisión de cloud y orquestador queda pendiente de métricas + IT corporativa
