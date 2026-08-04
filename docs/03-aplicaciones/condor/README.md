# 03 · Condor

## Propósito

Carpeta de descubrimiento de **Condor**: gestión de solicitudes, workflow, expediente obligatorio y no conformidades asociadas.

## Estado

- **Mapeo frontend:** resuelto.
- **Dysflow:** `missing`; no hay `.dysflow/project.json` en `00_main`. Inventario realizado con rutas explícitas de lectura.
- **Inventario `list_objects`:** 172 elementos resumidos (cifra preliminar).
- **Backend:** no presente en el checkout inspeccionado; ruta y nombre pendientes.
- **Documentación local ya disponible:** `00_CONDOR/docs/PRD/03_Gestion_Solicitudes.md` (entre otros).

## Lote asociado

Lote 6 de `exploration.md` — Condor.

## Entregable previsto

1. Capacidades de solicitudes y workflow; alta con expediente obligatorio.
2. Asociación con `TbNoConformidades` (CD_CA) y adjuntos.
3. Formularios, informes, batch y backend (una vez localizado).
4. Permisos y dependencias con Lanzadera, Expedientes y No_Conformidades.

## Fuentes de autoridad

1. `C:\00repos\documentacion\OPENSPEC\00_CONDOR`.
2. `C:\00repos\codigo\00_CONDOR\00_main` (PRDs y código).
3. Inspección Dysflow solo lectura (cuando se configure el target).
4. Engram solo como contexto histórico.

## Reglas de evidencia

- La FK `tbSolicitudes.idExpediente` se documenta como contrato de negocio, no como detalle de implementación.
- El backend se documenta **solo** cuando se localiza su ruta y nombre exactos.
- `.dysflow/project.json` no se crea ni se modifica en esta fase; queda como tarea del lote.

## Checklist

- [ ] Backend localizado y nombrado antes de cerrar el lote.
- [ ] Configuración Dysflow propuesta (sin aplicar) registrada en `08-decisiones-y-preguntas-abiertas.md`.
- [ ] APAP y APAP_WEB no aparecen.

## Siguiente paso

Resolver la ruta y nombre del backend vigente, y la configuración Dysflow de solo lectura, antes de poblar capacidades.
