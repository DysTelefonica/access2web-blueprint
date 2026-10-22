# D83 — HPS_Solicitudes se documenta como aplicación independiente de HPS: `TbAplicaciones.ID =…

- estado: adoptada
- fecha: 2026-08-05
- decide: usuario (dueño del producto)
- motivo: Estas tres preguntas bloqueaban el Lote 2 del descubrimiento legado. La resolución aporta D83 (independencia de HPS_Solicitudes), D84 (baselines de Condor y Brass) y D85 (baseline del catálogo con 8 IDs).
- origen: `blueprint/p1-p2-p5-resolved-aug2026`

## Decisión

HPS_Solicitudes se documenta como aplicación **independiente** de HPS: `TbAplicaciones.ID = 22` ("Solicitudes HPS"), con su propio checkout, frontend (`Solicitudes_HPS.accdb`) y backend (`Solicitudes_HPS_datos.accdb`). Comparte nombre conceptual con HPS (ID 17) por razones históricas pero son productos distintos
