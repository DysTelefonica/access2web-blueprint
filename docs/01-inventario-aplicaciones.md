# 01 · Inventario de aplicaciones

## Propósito

Registro vivo de las ocho aplicaciones dentro de alcance. Una fila por aplicación, con su documentación, checkout principal, binarios observables y estado del descubrimiento. **No** añade hallazgos, capacidades ni diagramas; solo identifica.

## Tabla de inventario

| Aplicación | Carpeta documental canónica | Checkout `00_main` | Frontend observable | Backend observable | Estado |
|---|---|---|---|---|---|
| Lanzadera | `C:\00repos\documentacion\OPENSPEC\00_LANZADERA` | `C:\00repos\codigo\00_LANZADERA\00_main` | `Lanzadera.accdb` | `Lanzadera_Datos.accdb` | Mapeo resuelto |
| Gestion_Riesgos | `C:\00repos\documentacion\OPENSPEC\00_GESTION_RIESGOS` | `C:\00repos\codigo\00_GESTION_RIESGOS\00_main` | `Gestion_Riesgos.accdb` | `Gestion_Riesgos_Datos.accdb` | Mapeo resuelto |
| No_Conformidades | `C:\00repos\documentacion\OPENSPEC\00_No_Conformidades` | `C:\00repos\codigo\00_NO_CONFORMIDADES\00_main` | `NoConformidades.accdb` | `NoConformidades_Datos.accdb` | Inventario Dysflow pendiente |
| Condor | `C:\00repos\documentacion\OPENSPEC\00_CONDOR` | `C:\00repos\codigo\00_CONDOR\00_main` | `CONDOR.accdb` | No presente en `00_main` | Backend y `.dysflow/project.json` pendientes |
| HPS_Solicitudes | `C:\00repos\documentacion\OPENSPEC\00_HPS_SOLICITUDES` | No localizado | — | — | Identidad de checkout pendiente |
| HPS | `C:\00repos\documentacion\OPENSPEC\00_HPS` | `C:\00repos\codigo\00_HPS\00_main` | `HPS.accdb` | `HPST.accdb` | Mapeo resuelto |
| Brass | `C:\00repos\documentacion\OPENSPEC\00_BRASS` | `C:\00repos\codigo\00_BRASS\00_main` | `Gestion_Brass_Gestion.accdb` | No presente en `00_main` | Backend y `.dysflow/project.json` pendientes |
| Expedientes | `C:\00repos\documentacion\OPENSPEC\00_EXPEDIENTES` | `C:\00repos\codigo\00_EXPEDIENTES\00_main` | `Expedientes.accdb` | `Expedientes_datos.accdb` (observado en staging) | Aviso: `path-mismatch` en Dysflow |

## Fuentes y reglas

- Esta tabla refleja el estado al cierre de `exploration.md`. Cualquier actualización debe citar el lote de descubrimiento y la fecha.
- Las celdas `—` indican vacío legítimo (no localizado), no error.
- Los nombres de fichero (`Lanzadera.accdb`, etc.) se citan textualmente; no se renombran aquí.

## Checklist

- [ ] Cada fila enlaza con su `README.md` en `03-aplicaciones/`.
- [ ] Los huecos se mantienen como `pendiente`; no se inventan rutas.
- [ ] APAP/APAP_WEB **no** aparecen en esta tabla.

## Siguiente paso

Cruzar esta tabla con la matriz de dependencias en `02-topologia-ecosistema/matriz-dependencias.md` para detectar aplicaciones aún sin consumidor identificado.
