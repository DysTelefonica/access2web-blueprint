# 01 · Inventario de aplicaciones

## Propósito

Registro vivo de las ocho aplicaciones dentro de alcance. Una fila por aplicación, con su documentación, checkout principal, binarios observables y estado del descubrimiento. **No** añade hallazgos, capacidades ni diagramas; solo identifica.

## Tabla de inventario

| Aplicación | Carpeta documental canónica | Checkout `00_main` (READ-ONLY) | Source base del snapshot staging | Snapshot self-contained en este repo | Frontend observable | Backend observable | Estado |
|---|---|---|---|---|---|---|---|
| Lanzadera | `C:\00repos\documentacion\OPENSPEC\00_LANZADERA` | `C:\00repos\codigo\00_LANZADERA\00_main` | `00_LANZADERA\staging` | `data/staging/lanzaderas/` | `Lanzadera.accdb` | `Lanzadera_Datos.accdb` | Mapeo resuelto |
| Gestion_Riesgos | `C:\00repos\documentacion\OPENSPEC\00_GESTION_RIESGOS` | `C:\00repos\codigo\00_GESTION_RIESGOS\00_main` | `00_GESTION_RIESGOS\staging` | `data/staging/gestion-riesgos/` | `Gestion_Riesgos.accdb` | `Gestion_Riesgos_Datos.accdb` | Descubrimiento Batch 3 completado; queja del árbol caracterizada (D88); Dysflow read-only pendiente para segunda pasada |
| No_Conformidades | `C:\00repos\documentacion\OPENSPEC\00_No_Conformidades` | `C:\00repos\codigo\00_NO_CONFORMIDADES\00_main` | `00_NO_CONFORMIDADES\staging` | `data/staging/no-conformidades/` | `NoConformidades.accdb` | `NoConformidades_Datos.accdb` | Inventario Dysflow pendiente |
| Condor | `C:\00repos\documentacion\OPENSPEC\00_CONDOR` | `C:\00repos\codigo\00_CONDOR\00_main` | `00_CONDOR\staging` | `data/staging/condor/` (default Dysflow) | `CONDOR.accdb` | `condor_datos.accdb` | **Snapshot completo** + default del runtime Dysflow |
| HPS_Solicitudes | `C:\00repos\documentacion\OPENSPEC\00_HPS_SOLICITUDES` | `C:\00repos\codigo\HPS_SOLICITUDES` (sin prefijo `00_`; fallback `main` por ausencia de checkout `00_HPS_SOLICITUDES/`/`staging/`) | `HPS_SOLICITUDES` (raíz) | `data/staging/hps-solicitudes/` | `Solicitudes_HPS.accdb` | `Solicitudes_HPS_datos.accdb` | **Snapshot completo** |
| HPS | `C:\00repos\documentacion\OPENSPEC\00_HPS` | `C:\00repos\codigo\00_HPS\00_main` | `00_HPS\staging` | `data/staging/hps/` | `HPS.accdb` | `HPST.accdb` | Mapeo resuelto |
| Brass | `C:\00repos\documentacion\OPENSPEC\00_BRASS` | `C:\00repos\codigo\00_BRASS\00_main` | `00_BRASS\00_main` (no hay `staging/`) | `data/staging/brass/` | `Gestion_Brass_Gestion.accdb` | (mismo archivo) | **Snapshot completo**; BRASS no tiene backend separado |
| Expedientes | `C:\00repos\documentacion\OPENSPEC\00_EXPEDIENTES` | `C:\00repos\codigo\00_EXPEDIENTES\00_main` | `00_EXPEDIENTES\staging` | `data/staging/expedientes/` | `Expedientes.accdb` | `Expedientes_datos.accdb` | Descubrimiento Batch 2 completado; Dysflow read-only |

## Fuentes y reglas

- Esta tabla refleja el estado al cierre de `exploration.md`. Cualquier actualización debe citar el lote de descubrimiento y la fecha.
- Las celdas `—` indican vacío legítimo (no localizado), no error.
- Los nombres de fichero (`Lanzadera.accdb`, etc.) se citan textualmente; no se renombran aquí.

## Checklist

- [x] Cada fila enlaza con su `README.md` en `03-aplicaciones/`.
- [x] Los huecos se mantienen como `pendiente`; no se inventan rutas.
- [x] APAP/APAP_WEB **no** aparecen en esta tabla.
- [x] Snapshot self-contained presente en `data/staging/` para las 8 apps (foundation, 2026-08-06).

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

- Los `.accdb` de uso viven bajo `C:\00repos\datos` (raíz única) **o** en `data/staging/<app>/` (snapshot self-contained). Se citan por nombre; no se exponen aquí rutas UNC ni hosts.
- Cada aplicación inspeccionada tiene dos commits de referencia: `00_main` (publicado) y, cuando exista, `staging` (baseline funcional). El Lote 1 usa Lanzadera `staging` como baseline y `00_main` como comparación publicada.
- Las configuraciones Dysflow con `path-mismatch` o ausentes **no se modifican** durante la fase de discovery. Su corrección es decisión del lote de modernización correspondiente.
- La inspección CodeGraph-VBA se realiza sobre `data/staging/<app>/.codegraph-vba/` (foundation 2026-08-06). Cada app tiene su propio índice regenerado.

## Siguiente paso

Cruzar esta tabla con la matriz de dependencias en `02-topologia-ecosistema/matriz-dependencias.md` para detectar aplicaciones aún sin consumidor identificado.

## Actualización del Lote 3 · 2026-08-05

Gestion_Riesgos tiene descubrimiento funcional, técnico y de migración documentado en `docs/03-aplicaciones/gestion-riesgos/`. La queja histórica del árbol de ediciones/riesgos queda caracterizada en [data-model.md § Rendimiento del árbol](docs/03-aplicaciones/gestion-riesgos/data-model.md#rendimiento-del-árbol-de-riesgos--causa-raíz-y-opciones-de-implementación) con propuesta técnica **D88** (HTMX + CTE recursivo + lazy expansion por nivel, descartando MSComctlLib.TreeView y el doble modelo `nuevo/antiguo`). El acoplamiento declarado con Lanzadera vía `getdbLanzadera()` se incorpora a D86/D87 como punto de migración. La inspección Dysflow sobre el binario y el backend autoritativo queda pendiente para una segunda pasada por límite de tiempo en esta sesión.

## Foundation · 2026-08-06

Setup del snapshot self-contained en `data/staging/` para las 8 apps. Total: 404.45 MB commiteados (decisión explícita: repo privado, secretos legacy aceptados). `.dysflow/project.json` unificado con default = Condor. `.codegraph-vba/` regenerado por app. Las 7 apps no-default se acceden con `accessPath`/`backendPath` por llamada. Detalle por app en `data/staging/README.md` y `data/staging/<app>/README.md`.

| App | Frontend | Backend | Source base | src files | docs files |
|---|---|---|---|---|---|
| `condor` | `CONDOR.accdb` | `condor_datos.accdb` | `00_CONDOR\staging` | 221 | 20 |
| `hps` | `HPS.accdb` | `HPST.accdb` | `00_HPS\staging` | 139 | — |
| `hps-solicitudes` | `Solicitudes_HPS.accdb` | `Solicitudes_HPS_datos.accdb` | `HPS_SOLICITUDES\` | 108 | 2 |
| `brass` | `Gestion_Brass_Gestion.accdb` | (mismo archivo) | `00_BRASS\00_main` | 218 | — |
| `gestion-riesgos` | `Gestion_Riesgos.accdb` | `Gestion_Riesgos_Datos.accdb` | `00_GESTION_RIESGOS\staging` | 328 | 29 |
| `no-conformidades` | `NoConformidades.accdb` | `NoConformidades_Datos.accdb` | `00_NO_CONFORMIDADES\staging` | 249 | 51 |
| `expedientes` | `Expedientes.accdb` | `Expedientes_datos.accdb` | `00_EXPEDIENTES\staging` | 302 | 69 |
| `lanzaderas` | `Lanzadera.accdb` | `Lanzadera_Datos.accdb` | `00_LANZADERA\staging` | 89 | — |
