# 04 · Integraciones y operación — Rutas, entornos y contingencia

## Propósito

Describe cómo se localizan los binarios y configuraciones en cada entorno (desarrollo, staging, producción) y qué hacer cuando una ruta no resuelve. Distingue **inventario físico** (lo que está en disco) de **rutas operativas** (cómo se accede en cada entorno). No introduce arquitectura web objetivo; solo documenta el estado y la política de contingencia del legado.

## Reglas transversales

- **Backends compartidos bajo `C:\00repos\datos`**: los `.accdb` de uso viven en esa raíz. No se exponen rutas UNC, hosts ni cuentas en este documento.
- **Frontends versionados bajo `C:\00repos\codigo`**: cada aplicación reside en una carpeta propia (`00_<aplicacion>\00_main`, `00_<aplicacion>\staging` o variante local). Las `.dysflow/project.json` son la fuente de la ruta real de cada binario.
- **Backends por aplicación** (referencia, tomados del inventario):

| Aplicación | Backend observable |
|---|---|
| Lanzadera | `C:\00repos\datos\Lanzadera_Datos.accdb` |
| Gestion_Riesgos | `C:\00repos\datos\Gestion_Riesgos_Datos.accdb` |
| No_Conformidades | `C:\00repos\datos\NoConformidades_Datos.accdb` |
| Condor | `C:\00repos\datos\condor_datos.accdb` |
| HPS_Solicitudes | `C:\00repos\datos\Solicitudes_HPS_datos.accdb` |
| HPS | `C:\00repos\datos\HPST.accdb` |
| Brass | `C:\00repos\datos\Gestion_Brass_Gestion_Datos.accdb` |
| Expedientes | `C:\00repos\datos\Expedientes_datos.accdb` (confirmado como backend autoritativo en Batch 2, 2026-08-05) |

- **No mutar configuraciones Dysflow** durante la fase de discovery. Los avisos de `path-mismatch` o las configuraciones ausentes se resuelven en un lote aprobado de modernización.
- **APAP y APAP_WEB fuera de alcance**: este documento no las contempla.

## Estado actual por aplicación

| Aplicación | Checkout frontend | `.dysflow/project.json` | Backend en `C:\00repos\datos` | Avisos |
|---|---|---|---|---|
| Lanzadera | `C:\00repos\codigo\00_LANZADERA\00_main` (publicado) y `staging` (baseline funcional) | `valid`, declara backend relativo local | Localizado | Inspección Dysflow operativa en ambos worktrees; el backend compartido se selecciona explícitamente como ruta de lectura. |
| Gestion_Riesgos | `C:\00repos\codigo\00_GESTION_RIESGOS\00_main` | `valid` | Localizado | — |
| No_Conformidades | `C:\00repos\codigo\00_NO_CONFORMIDADES\00_main` | `valid`, pero `list_objects` falló | Localizado | Inventario Dysflow pendiente. |
| Condor | `C:\00repos\codigo\00_CONDOR\00_main` | `missing` | Localizado | Restaurar `.dysflow/project.json` antes de cerrar el lote. |
| HPS_Solicitudes | `C:\00repos\codigo\HPS_SOLICITUDES` (sin prefijo `00_`) | `missing` | Localizado | Sin staging local; se mantiene como fallback `main` por ausencia de `00_HPS_SOLICITUDES/`. |
| HPS | `C:\00repos\codigo\00_HPS\00_main` | `valid` | Localizado | Staging y main comparten commit. |
| Brass | `C:\00repos\codigo\00_BRASS\00_main` | `missing` | Localizado | Restaurar `.dysflow/project.json` antes de cerrar el lote. |
| Expedientes | `C:\00repos\codigo\00_EXPEDIENTES\00_main` | `path-mismatch` (`accessPath` absoluto heredado) | Observado en staging, no resoluble desde main | Aviso heredado; se respeta el estado actual. |

## Baselines operativas de discovery

- **`00_main`** se trata como referencia publicada; no se modifica durante la fase de discovery.
- **`staging`** es la baseline funcional cuando existe. El Lote 1 (Lanzadera) usa staging como baseline y main solo como comparación publicada.
- **CodeGraph-VBA** solo se consulta cuando existe `.codegraph-vba` en el worktree objetivo. En esta pasada los índices existentes son los de Lanzadera `staging`, Lanzadera `00_main`, Gestion_Riesgos, HPS y Expedientes. No se inicializan, sincronizan ni reconstruyen índices en esta pasada.
- **Dysflow** se ejecuta siempre en modo solo lectura sobre los worktrees autorizados.

## Procedimiento de contingencia

1. Si Dysflow falla por configuración (`valid` con error, `path-mismatch` o `missing`), documentar el aviso y continuar discovery con CodeGraph-VBA y lectura de código donde sea posible.
2. Si el backend no responde, no se consulta. No se infieren rutas desde artefactos no observables.
3. Si el frontend no está localizable, se documenta el motivo (carpeta inexistente, ruta renombrada, dependencia externa) y se pregunta al usuario antes de cualquier suposición.
4. Si la documentación describe una aplicación que no tiene binario localizado (caso HPS_Solicitudes antes del 2026-08-05), se etiqueta como **discrepancia crítica** y se resuelve su identidad antes de continuar con su lote.

## Avisos heredados (no resueltos aquí)

- Condor y Brass sin `.dysflow/project.json`.
- Expedientes con `accessPath` absoluto y `path-mismatch`.
- No_Conformidades con `list_objects` fallido pese a configuración `valid`.
- HPS_Solicitudes sin checkout bajo el patrón `00_*` y sin `.dysflow/project.json`.
- Backends ausentes en algunos `00_main` (los `.accdb` están en `C:\00repos\datos`, no dentro del repo): se documenta la separación entre versiones y binarios operativos.

## Fuentes de autoridad

1. `C:\00repos\codigo\<app>\00_main` por aplicación.
2. `C:\00repos\documentacion\OPENSPEC\` por aplicación.
3. Inspección Dysflow solo lectura.
4. CodeGraph-VBA para trazado de símbolos.
5. Engram como contexto histórico.

## Reglas de evidencia

- Las rutas se citan tal cual aparecen en disco; no se normalizan a una convención única en esta fase.
- Los entornos se distinguen explícitamente; ningún dato cruza sin marcarlo.
- APAP y APAP_WEB no aparecen.

## Checklist

- [ ] Cada ruta lleva su entorno (desarrollo, staging, producción) y su fecha de observación.
- [ ] Los avisos heredados están enlazados a `08-decisiones-y-preguntas-abiertas.md`.
- [ ] La baseline de discovery (staging vs main) está declarada para cada aplicación que la tenga.

## Siguiente paso

Cruzar con `02-topologia-ecosistema/topologia-frontends-backends.md` para evitar duplicación de rutas y mantener una única fuente de verdad.
