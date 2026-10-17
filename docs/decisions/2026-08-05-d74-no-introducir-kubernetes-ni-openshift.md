# D74 — No introducir Kubernetes ni OpenShift prematuramente

- estado: adoptada
- fecha: 2026-08-05
- decide: usuario (dueño del producto)
- motivo: La sección 4 original introducía InsForge (BaaS) como backend, lo que rompe la propiedad hexagonal. También proponía OpenShift + GCP para 200 usuarios concurrentes, un dimensionamiento diez veces excesivo. Las notas sobre persistencia de SSE están obsoletas porque SSE se rechazó en la sección 1.
- origen: `external-prompt-review/section-4-resolution`

## Decisión

No introducir Kubernetes ni OpenShift prematuramente. Para 200 usuarios concurrentes, una instancia de FastAPI + PostgreSQL es suficiente
