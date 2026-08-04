# 03 · Brass

## Propósito

Carpeta de descubrimiento de **Brass**: gestión de eventos, actividades, materiales, equipos, SLA, informes y facturación, con identidad y permisos procedentes de Lanzadera.

## Estado

- **Mapeo frontend:** resuelto.
- **Dysflow:** `missing`; no hay `.dysflow/project.json` en `00_main`. Inventario realizado con rutas explícitas de lectura.
- **Inventario `list_objects`:** 220 elementos resumidos (cifra preliminar).
- **Backend:** no presente en el checkout inspeccionado; ruta y nombre pendientes.
- **Documentación local ya disponible:** `00_BRASS/docs/ERD/Lanzadera_Datos.md` y `00_No_Conformidades/docs/ERD/Estructura_Datos_Lanzadera.md` como evidencia de catálogo y permisos.

## Lote asociado

Lote 7 de `exploration.md` — Brass.

## Entregable previsto

1. Capacidades de eventos, actividades, equipos, materiales, SLA y facturación.
2. Permisos y código de usuario/permisos procedente de Lanzadera.
3. Rastreo específico de `IDExpediente` / `CodExp` (pendiente en esta pasada).
4. Informes, batch y backend (una vez localizado).

## Fuentes de autoridad

1. `C:\00repos\documentacion\OPENSPEC\00_BRASS`.
2. `C:\00repos\codigo\00_BRASS\00_main`.
3. Inspección Dysflow solo lectura (cuando se configure el target).
4. Engram solo como contexto histórico.

## Reglas de evidencia

- El vínculo actual con `IDExpediente`/`CodExp` se documenta como `pendiente`; no se asume.
- Las entidades Brass se citan textualmente; no se renombran aquí.
- `.dysflow/project.json` no se crea ni se modifica en esta fase.

## Checklist

- [ ] Backend localizado y nombrado antes de cerrar el lote.
- [ ] Rastreo de `IDExpediente`/`CodExp` ejecutado con su evidencia.
- [ ] APAP y APAP_WEB no aparecen.

## Siguiente paso

Resolver la ruta y nombre del backend vigente, y la configuración Dysflow de solo lectura, antes de poblar capacidades.
