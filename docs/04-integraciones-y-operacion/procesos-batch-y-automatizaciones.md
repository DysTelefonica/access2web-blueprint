# 04 · Integraciones y operación — Procesos batch y automatizaciones

## Propósito

Inventario de los procesos batch, scripts, tareas programadas y automatizaciones que dan soporte al ecosistema (sincronizaciones, cargas, depuraciones, reporting nocturno, etc.). Describe el qué, no diseña cómo migrarlo.

## Qué va aquí

- Scripts `.bat`, `.ps1`, `.vbs` u otros en los checkouts.
- Tareas programadas (orquestador o sistema) y su evidencia documental.
- Automatizaciones VBA invocadas desde eventos, timers o formularios de mantenimiento.
- Para Gestion_Riesgos: `00_Automatizaciones/docs/scripts/GestionRiesgos_bat.md` ya localizado.

## Estado del contenido

**Placeholder.** Se poblará con los hallazgos de los Lotes 1 a 8 y se consolidará en el Lote 9.

## Fuentes de autoridad

1. `C:\00repos\codigo\<app>\00_main` por aplicación.
2. `C:\00repos\documentacion\OPENSPEC\` por aplicación.
3. Inspección Dysflow solo lectura.
4. Engram solo como contexto histórico.

## Reglas de evidencia

- Cada automatización lleva su ruta exacta, su planificador y su propósito declarado.
- Los efectos sobre datos (escritura, borrado, archivo) se documentan con su tabla destino.
- Las automatizaciones entre aplicaciones se marcan como `cross-app` y se referencian desde ambas fichas.

## Checklist

- [ ] Cada script lleva ruta y propósito.
- [ ] Los efectos sobre datos están citados, no asumidos.
- [ ] APAP y APAP_WEB no aparecen.

## Siguiente paso

Cruzar con `tablas-vinculadas-y-backends.md` y con `informes-exportaciones-y-correo.md` para mantener la coherencia de rutas y salidas.
