# 03 · HPS

## Propósito

Carpeta de descubrimiento de **HPS**: solicitudes, usuarios HPS, expedientes, SICA, históricos, anexos e integraciones documentales con Expedientes y Lanzadera.

## Estado

- **Mapeo:** resuelto.
- **Dysflow:** `valid` como `00-hps-staging`; frontend y backend resueltos.
- **Inventario `list_objects`:** 130 elementos resumidos (cifra preliminar).
- **Backend observado:** 22 tablas; frontend 12 tablas.
- **Documentación local ya disponible:** specs, UAT y releases en `00_HPS`.

## Lote asociado

Lote 4 de `exploration.md` — HPS.

## Entregable previsto

1. Capacidades de solicitudes, usuarios HPS, entidades y SICA.
2. Relación con `TbExpedientes`, `IDExpediente`, `TbSuministradores` y `IDSolicitud`.
3. Históricos, anexos, exportaciones Excel y backends.
4. Permisos vía `UsuarioAplicacionPermisos` y acceso a Lanzadera.
5. Separación entre `HPST.accdb` actual y la consolidación histórica documentada.

## Fuentes de autoridad

1. `C:\00repos\documentacion\OPENSPEC\00_HPS`.
2. `C:\00repos\codigo\00_HPS\00_main`.
3. Inspección Dysflow solo lectura.
4. CodeGraph-VBA ya presente en este `00_main`.
5. Engram solo como contexto histórico.

## Reglas de evidencia

- Se distingue explícitamente el estado actual (`HPST.accdb`) de la consolidación histórica.
- `TbSuministradores` se documenta con su vínculo declarado a `Expedientes_datos.accdb` y con su semántica (vinculada, no dueña del dato).
- `IDSolicitud` se traza a lo largo de la cadena de solicitudes sin extrapolar.

## Checklist

- [ ] Estado actual y consolidación histórica separados en capítulos.
- [ ] Permisos de HPS y de Lanzadera referenciados sin duplicar el catálogo.
- [ ] APAP y APAP_WEB no aparecen.

## Siguiente paso

Cruzar con `02-topologia-ecosistema/expedientes-ciclo-de-vida.md` para fijar el contrato de `IDExpediente` desde la óptica de HPS.
