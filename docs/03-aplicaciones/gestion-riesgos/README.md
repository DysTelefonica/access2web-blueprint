# 03 · Gestion_Riesgos

## Propósito

Carpeta de descubrimiento de **Gestion_Riesgos**: gestión de riesgos, proyectos, tareas, correos y su dependencia documentada de Lanzadera y Expedientes.

## Estado

- **Mapeo:** resuelto.
- **Dysflow:** `valid`; frontend y backend resueltos.
- **Inventario `list_objects`:** 193 elementos resumidos (cifra preliminar).
- **Backend observado:** 68 tablas; frontend contiene `TbAuxPriorizacion`.
- **Documentación local ya disponible:** `00_GESTION_RIESGOS/00_main/docs/DISCOVERY_MAP.md` y ERD en el checkout principal.

## Lote asociado

Lote 3 de `exploration.md` — Gestion_Riesgos (previo a este, Lote 1 y Lote 2).

## Entregable previsto

1. Capacidades de proyectos y riesgos: alta, edición, priorización, cierre.
2. Formularios por capacidad (no listado plano de forms).
3. Tablas locales y vinculadas; naturaleza de `TbExpedientes1` (copia, vista o contrato histórico).
4. Permisos, informes, automatizaciones batch.
5. Dependencias con Lanzadera (`getdbLanzadera`, `UsuarioAplicacionPermisos`) y Expedientes (`getdbExpedientes`).

## Fuentes de autoridad

1. `C:\00repos\documentacion\OPENSPEC\00_GESTION_RIESGOS`.
2. `C:\00repos\codigo\00_GESTION_RIESGOS\00_main` (incluye `docs/DISCOVERY_MAP.md` y ERD).
3. Inspección Dysflow solo lectura.
4. CodeGraph-VBA ya presente en este `00_main`.
5. Engram solo como contexto histórico.

## Reglas de evidencia

- `TbExpedientes1` se documenta como `pendiente de naturaleza` hasta verificación dirigida.
- Las llamadas `getdbLanzadera` y `getdbExpedientes` se trazan con su símbolo y callers.
- Las automatizaciones se citan con su script y su planificador, sin extrapolar su alcance.

## Checklist

- [ ] No se eleva `TbExpedientes1` a sinónimo de `TbExpedientes` sin prueba.
- [ ] Cada automatización lleva ruta a su `.bat`/script y a su evidencia.
- [ ] APAP y APAP_WEB no aparecen.

## Siguiente paso

Cruzar resultados con `02-topologia-ecosistema/matriz-dependencias.md` al cierre del lote.
