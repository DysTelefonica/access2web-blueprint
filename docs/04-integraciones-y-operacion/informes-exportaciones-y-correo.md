# 04 · Integraciones y operación — Informes, exportaciones y correo

## Propósito

Catálogo operativo de los informes, exportaciones (Excel, PDF, otros) y envíos de correo que las aplicaciones legacy producen de forma regular. No incluye capacidades funcionales (esas viven en `05-capacidades/`), solo los **puntos de salida** observables.

## Qué va aquí

- Informes Access por aplicación, con su ruta de plantilla y disparador.
- Exportaciones a Excel/CSV/PDF: consultas, macros y código VBA que las genera.
- Correo saliente: plantillas, adjuntos, remitentes técnicos y excepciones.

## Estado del contenido

**Placeholder.** Se poblará con los hallazgos de los Lotes 1 a 8 y se consolidará en el Lote 9.

## Fuentes de autoridad

1. `C:\00repos\codigo\<app>\00_main` por aplicación.
2. `C:\00repos\documentacion\OPENSPEC\` por aplicación.
3. Inspección Dysflow solo lectura.
4. Engram solo como contexto histórico.

## Reglas de evidencia

- Cada informe y exportación cita módulo, macro o consulta origen.
- Los envíos de correo llevan su plantilla y su cuenta remitente.
- Las automatizaciones se cruzan con `procesos-batch-y-automatizaciones.md`.

## Checklist

- [ ] Cada salida lleva su disparador (evento, batch, manual) y su evidencia.
- [ ] Los adjuntos de correo y los destinos se citan textualmente.
- [ ] APAP y APAP_WEB no aparecen.

## Siguiente paso

Cruzar con `05-capacidades/` para que cada capacidad declare sus salidas observables.
