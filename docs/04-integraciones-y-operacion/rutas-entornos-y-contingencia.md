# 04 · Integraciones y operación — Rutas, entornos y contingencia

## Propósito

Describe cómo se localizan los binarios y configuraciones en cada entorno (desarrollo, staging, producción) y qué hacer cuando una ruta no resuelve. **No** introduce arquitectura web objetivo; solo documenta el estado legacy.

## Qué va aquí

- Tabla de rutas conocidas por entorno, por aplicación (cuando se conozcan).
- Procedimiento de contingencia cuando un backend no está accesible.
- Avisos heredados: `path-mismatch` de Expedientes, ausencia de `.dysflow/project.json` en Condor y Brass.
- Procedimiento de fallback a la documentación cuando el runtime no responde.

## Estado del contenido

**Placeholder.** Se poblará con los hallazgos de los Lotes 1 a 8 y se consolidará en el Lote 9.

## Fuentes de autoridad

1. `C:\00repos\codigo\<app>\00_main` por aplicación.
2. `C:\00repos\documentacion\OPENSPEC\` por aplicación.
3. Inspección Dysflow solo lectura.
4. Engram solo como contexto histórico.

## Reglas de evidencia

- Las rutas se citan tal cual aparecen; no se normalizan aquí.
- Los entornos se distinguen explícitamente; ningún dato cruza sin marcarlo.
- Las contingencias se describen a nivel operativo, no de diseño de migración.

## Checklist

- [ ] Cada ruta lleva su entorno y su fecha de observación.
- [ ] Los avisos heredados están enlazados a `08-decisiones-y-preguntas-abiertas.md`.
- [ ] APAP y APAP_WEB no aparecen.

## Siguiente paso

Cruzar con `02-topologia-ecosistema/topologia-frontends-backends.md` para evitar duplicación de rutas.
