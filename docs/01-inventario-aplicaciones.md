# 01 · Inventario de aplicaciones

## Propósito

Registro vivo de las ocho aplicaciones dentro de alcance. Una fila por aplicación, con su documentación, checkout principal, binarios observables y estado del descubrimiento. **No** añade hallazgos, capacidades ni diagramas; solo identifica.

## Tabla de inventario

| Aplicación | Carpeta documental canónica | Checkout `00_main` | Frontend observable | Backend observable (raíz `C:\00repos\datos`) | Estado |
|---|---|---|---|---|---|
| Lanzadera | `C:\00repos\documentacion\OPENSPEC\00_LANZADERA` | `C:\00repos\codigo\00_LANZADERA\00_main` | `Lanzadera.accdb` | `Lanzadera_Datos.accdb` | Mapeo resuelto |
| Gestion_Riesgos | `C:\00repos\documentacion\OPENSPEC\00_GESTION_RIESGOS` | `C:\00repos\codigo\00_GESTION_RIESGOS\00_main` | `Gestion_Riesgos.accdb` | `Gestion_Riesgos_Datos.accdb` | Mapeo resuelto |
| No_Conformidades | `C:\00repos\documentacion\OPENSPEC\00_No_Conformidades` | `C:\00repos\codigo\00_NO_CONFORMIDADES\00_main` | `NoConformidades.accdb` | `NoConformidades_Datos.accdb` | Inventario Dysflow pendiente |
| Condor | `C:\00repos\documentacion\OPENSPEC\00_CONDOR` | `C:\00repos\codigo\00_CONDOR\00_main` | `CONDOR.accdb` | `condor_datos.accdb` | Backend en `C:\00repos\datos`; `.dysflow/project.json` pendiente |
| HPS_Solicitudes | `C:\00repos\documentacion\OPENSPEC\00_HPS_SOLICITUDES` | `C:\00repos\codigo\HPS_SOLICITUDES` (sin prefijo `00_`; fallback `main` por ausencia de checkout `00_HPS_SOLICITUDES/`/`staging/`) | Pendiente de inspección | `Solicitudes_HPS_datos.accdb` | Checkout localizado tras el Lote 1; `.dysflow/project.json` y CodeGraph pendientes |
| HPS | `C:\00repos\documentacion\OPENSPEC\00_HPS` | `C:\00repos\codigo\00_HPS\00_main` | `HPS.accdb` | `HPST.accdb` | Mapeo resuelto |
| Brass | `C:\00repos\documentacion\OPENSPEC\00_BRASS` | `C:\00repos\codigo\00_BRASS\00_main` | `Gestion_Brass_Gestion.accdb` | `Gestion_Brass_Gestion_Datos.accdb` | Backend en `C:\00repos\datos`; `.dysflow/project.json` pendiente |
| Expedientes | `C:\00repos\documentacion\OPENSPEC\00_EXPEDIENTES` | `C:\00repos\codigo\00_EXPEDIENTES\00_main` | `Expedientes.accdb` | `C:\00repos\datos\Expedientes_datos.accdb` | Descubrimiento Batch 2 completado; Dysflow read-only |

## Fuentes y reglas

- Esta tabla refleja el estado al cierre de `exploration.md`. Cualquier actualización debe citar el lote de descubrimiento y la fecha.
- Las celdas `—` indican vacío legítimo (no localizado), no error.
- Los nombres de fichero (`Lanzadera.accdb`, etc.) se citan textualmente; no se renombran aquí.

## Checklist

- [ ] Cada fila enlaza con su `README.md` en `03-aplicaciones/`.
- [ ] Los huecos se mantienen como `pendiente`; no se inventan rutas.
- [ ] APAP/APAP_WEB **no** aparecen en esta tabla.

## Actualización del Lote 2 · 2026-08-05

Expedientes tiene descubrimiento funcional, técnico y de migración documentado en `docs/03-aplicaciones/expedientes/`. El backend autoritativo queda confirmado bajo `C:\00repos\datos`; el cambio no modifica los checkouts ni sus configuraciones Dysflow.

## Actualización del Lote 1 · 2026-08-04

Lanzadera ya tiene inventario normalizado y documentación de descubrimiento en `docs/03-aplicaciones/lanzadera/`. El estado de las otras siete filas no cambia en este lote.

## Actualización posterior al Lote 1 · 2026-08-05

- **HPS_Solicitudes**: el checkout se localizó en `C:\00repos\codigo\HPS_SOLICITUDES` (sin prefijo `00_`). Se mantiene como referencia `main` por ausencia de checkout `00_HPS_SOLICITUDES/` o `staging/`. Quedan pendientes la creación del `.dysflow/project.json`, la inspección Dysflow y la inicialización del índice CodeGraph-VBA local.
- **Brass y Condor**: backends localizados en `C:\00repos\datos`. Persiste el aviso por ausencia de `.dysflow/project.json`; no se mutan configuraciones en esta fase.
- **Lanzadera**: disposiciones finales consolidadas en `03-aplicaciones/lanzadera/capabilities.md` y `02-topologia-ecosistema/lanzadera-identidad-permisos.md`.

## Baselines operativas

Para evitar drift entre bases de discovery:

- Los `.accdb` de uso viven bajo `C:\00repos\datos` (raíz única). Se citan por nombre; no se exponen aquí rutas UNC ni hosts.
- Cada aplicación inspeccionada tiene dos commits de referencia: `00_main` (publicado) y, cuando exista, `staging` (baseline funcional). El Lote 1 usa Lanzadera `staging` como baseline y `00_main` como comparación publicada.
- Las configuraciones Dysflow con `path-mismatch` o ausentes **no se modifican** durante la fase de discovery. Su corrección es decisión del lote de modernización correspondiente.
- La inspección CodeGraph-VBA solo se realiza cuando existe `.codegraph-vba` en la aplicación objetivo. En este repo solo Lanzadera `staging`, Lanzadera `00_main`, Gestion_Riesgos, HPS y Expedientes tienen índice disponible. Condor y Brass no se inspeccionan por CodeGraph-VBA hasta preparar su repo.

## Siguiente paso

Cruzar esta tabla con la matriz de dependencias en `02-topologia-ecosistema/matriz-dependencias.md` para detectar aplicaciones aún sin consumidor identificado.
