# NoConformidades — integraciones y automatización

## Contratos observados

| Sistema | Uso | Evidencia |
|---|---|---|
| **Lanzadera** (identidad y permisos) | identidad, permisos y entorno | `constructor.bas:119` (`getUsuario` consulta `TbUsuariosAplicaciones`; presumiblemente vía `getdbLanzadera` o acceso directo a la BD de Lanzadera) |
| **Expedientes** | vínculo por código de NC ↔ código de expediente | `Expediente.cls`, `ExpedienteResponsable.cls` (clases compartidas) |
| **Gestion_Riesgos** | selección de riesgos vinculados a NC de Proyecto | `Riesgo.cls`, `RiesgoServicio.cls`, `Form_formRiesgosSeleccion.cls` |
| HPS | presumiblemente pedidos | (presumido, no verificado en código) |
| AGEDYS | presumiblemente DPD/proyectos | (presumido, no verificado en código) |
| Correos | avisos de error al administrador, cierre de NC | `Correo.cls`, `Form_FormCorreo.cls` |
| SharePoint/ficheros | enlaces de documentación | referencias en `DocumentoAuditoria.cls`, `DocumentoProyecto.cls` |
| Informes | generación de informes (NC de Auditorías + otros) | `Informe.cls`, `InformeNCAuditorias.cls` |

## Acoplamiento declarado con Lanzadera

La aplicación accede a la identidad/permisos del usuario vía `constructor.bas:getUsuario` que consulta `TbUsuariosAplicaciones`. El patrón es idéntico al de Lanzadera/Expedientes/Gestion_Riesgos. La diferencia con Gestion_Riesgos es que NoConformidades **no muestra una llamada explícita a `getdbLanzadera`** en el código inspeccionado — probablemente accede vía tablas compartidas o un método equivalente.

Implicaciones para la nueva plataforma:

- La identidad debe resolverse vía el **adaptador unificado de autenticación** (D9-D10), no vía acceso directo a la BD de Lanzadera.
- Los permisos por aplicación se cargan con `IDAplicacion = "8"` (producción) o `"81"` (pruebas).

## Automatización y tareas

`SegTareasAuditoria` y `SegTareasProyecto` mantienen worklists con seguimiento de tareas (calendario, responsables, cierre). En la nueva plataforma:

- El patrón de actualización de contadores pendientes vía polling HTMX con `hx-trigger="every 30s"` (D69) aplica también aquí.
- Las tareas se traducen a queries indexadas con sus propios endpoints.
- El envío de avisos (correos) pasa por el servicio unificado de notificaciones (D11-D13).

## Capa de caché propia (referencia de diseño)

NoConformidades tiene la **capa de caché selectivo más rica del ecosistema**. Se detalla en [Matriz de migración § D91](migration-matrix.md#d91-caché-selectivo-maduro-como-referencia-del-puerto-de-caché). Resumen:

- **Tabla `TbCacheNCProyecto`** con snapshots JSON por NC (`DatosNC`, `DatosACs`, `DatosARs`), hits, tamaño, fecha de uso, versión, validez.
- **Comandos de mantenimiento** bien documentados y expuestos (invalidar, eliminar, regenerar, poblar masivamente, limpiar logs, diagnosticar, estadísticas).
- **Kill switch** (`Test_KillSwitch.bas`) para activar/desactivar la caché en runtime de forma atómica.
- **Logs de caché** (`TbLogCache`) con tipos de operación y errores.
- **Diagnóstico de integridad** (huérfanos, JSONs vacíos).

Esta capa es la **referencia principal** del blueprint para el puerto de caché de la nueva plataforma (D70-D71).

## Configuración y flags de operación

Flags activos hoy en producción (inferidos de `Variables Globales.bas`):

| Flag / Campo | Valor por defecto | Significado |
|---|---|---|
| `BackendActivo` | `PROD` (producción) / `LOCAL` / `SANDBOX` | Cuál backend se usa |
| `EnPruebas` | `Sí`/`No` (texto) | Si está en pruebas; cuando `"Sí"`, IDAplicacion pasa a `"81"` |
| `IDAplicacion` | `"8"` (producción) / `"81"` (pruebas) | ID en `TbAplicaciones` |
| `PasswordBackend` | texto | Contraseña del backend (hoy en `backends.json` ⚠️ D90) |
| `CacheHabilitada` (en `TbConfiguracion`) | Boolean | Habilita/deshabilita la caché (kill switch) |
| `BackendSandboxURL` / `BackendSandboxPassword` (testing mode) | — | Solo en `m_TestingMode = True` |

En la nueva plataforma estos flags se mueven a **configuración del módulo** (no TempVars, no `backends.json`):

- `BackendActivo`, `BackendSandboxURL`, etc. → config del puerto de persistencia.
- `EnPruebas` → variable de entorno del runner.
- `IDAplicacion` → config del módulo.
- `PasswordBackend` → **secret manager / env var** (nunca en repo).
- `CacheHabilitada` → config del módulo + kill switch runtime vía endpoint admin.

Para rutas UNC, hosts y nombres de máquina concretos: **NO se reproducen en este artefacto** (regla de evidencia). Las rutas operativas se mantienen en `04-integraciones-y-operacion/rutas-entornos-y-contingencia.md` con criterio de privacidad.