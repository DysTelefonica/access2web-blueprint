# Condor — integraciones y automatización

## Contratos observados

| Sistema | Uso | Evidencia |
|---|---|---|
| **Lanzadera** (identidad y permisos) | identidad, permisos, entorno (vía `getdbLanzadera()` en `FUNCIONES UTILES.bas:84`) | `src/modules/FUNCIONES UTILES.bas:84`; `src/modules/UsuarioRepositorio.bas:9`; `UsuarioAplicacionPermisos.cls` |
| **Expedientes** | `idExpediente` (FK conceptual) + búsqueda de expedientes | `Expediente.cls`, `ExpedienteServicio.cls`, `ExpedienteViewModel.cls` |
| **NoConformidades** | `idNCAsociada` (FK conceptual) + verificación antes de eliminar | `Form_frmGestionSolicitud.cls:1908`; `NoConformidad.cls`; `NoConformidadServicio.cls` |
| **Gestion_Riesgos** | presumiblemente vía código compartido (no inspeccionado en detalle) | — |
| **HPS** | presumiblemente vía código compartido (no inspeccionado en detalle) | — |
| AGEDYS | presumiblemente vía código compartido (no inspeccionado en detalle) | — |
| Correos | `CorreoRepositorio.bas` + `ModuloCoreo` | `src/modules/CorreoRepositorio.bas` |
| SharePoint/ficheros | presumiblemente vía adjuntos (`Adjunto.cls`) | `Adjunto.cls` |

## Edge WebView embebido (patrón único de Condor)

`Form_frmGestionSolicitud.cls:209` usa `Me.webInfo.Navigate rutaNavegacion` con un control WebView Edge embebido en Access. Esto es **único en el ecosistema** — las otras apps son Access puro.

Mecanismo:
1. `MostrarVista(vista)` genera HTML dinámico vía `WorkflowServicio`.
2. `WebVisorCacheServicio` cachea el HTML.
3. `SnapshotServicio` toma un snapshot del estado actual a fichero.
4. `Me.webInfo.Navigate rutaTemporal` carga el HTML en el control Edge.

Esto es un **puente Access → web** dentro del propio Access. La nueva plataforma web **absorbe** este patrón: las vistas son páginas nativas, no embebidas en Access.

## Acoplamiento declarado con Lanzadera

La aplicación accede a la identidad/permisos del usuario vía `getdbLanzadera()` en `FUNCIONES UTILES.bas:84`. Patrón idéntico al de las otras apps.

Implicaciones para la nueva plataforma:

- La identidad debe resolverse vía el **adaptador unificado de autenticación** (D9-D10), no vía acceso directo a la BD de Lanzadera.
- Los permisos por aplicación se cargan con `IDAplicacion = "23"` (producción).
- El acceso directo a `TbUsuariosAplicaciones` en la BD de Lanzadera debe **eliminarse** y sustituirse por el servicio de identidad/permisos expuesto por la plataforma.

## Testing sandbox seguro (patrón único de Condor)

`FUNCIONES UTILES.bas:84-120` implementa un patrón maduro de testing sandbox que se preserva como referencia del puerto de testing de la nueva plataforma:

- `m_TestingMode=True` enruta `getdb()` a `m_BackendSandboxURL`.
- `m_BackendSandboxPassword` se usa en lugar de `m_PasswordBackend` para evitar contaminar la cache de producción.
- Cache safety (Spec-008): si `g_dbCondor` apunta a una URL distinta del sandbox configurado, **se cierra y se reabre** automáticamente.
- Si `m_BackendSandboxURL` está vacío, **`Err.Raise 513 "TESTS BLOCKED"`** con mensaje explícito. **Fail-fast** (no fallback silencioso).
- Test session management: `m_TestOnly*` overrides + `BeginTestSession` / `EndTestSession` / `ResetTestSession`.

Este patrón **NO se rompe en la nueva plataforma**. Se traduce al **puerto de testing** de FastAPI con `pytest` + fixtures que aíslan la BD de pruebas.

## Automatización y notificaciones

- `NotificacionServicio.cls`: envío de notificaciones (correos, avisos in-app).
- `MockNotifServ.cls`: test double para notificaciones (confirma D87 — tests VBA como evidencia de comportamiento).
- `WorkflowServicio.cls`: motor de transiciones de estado del workflow.
- `SnapshotServicio.cls`: snapshots del estado para WebView embebido (se traduce a endpoints de snapshot en la nueva plataforma).
- `WebVisorCacheServicio.cls`: caché de HTML renderizado (se traduce a fragment caching con HTMX + ETag en la nueva plataforma, ver D72).
- `LogError.bas` + `tbLogErrores`: log de errores estructurado (coherente con D27).

## Configuración y flags de operación

Flags activos hoy en staging (inferidos de `Variables Globales.bas`):

| Flag / Campo | Valor por defecto | Significado |
|---|---|---|
| `BackendActivo` | `PROD` (producción) / `LOCAL` / `SANDBOX` | Cuál backend se usa |
| `IDAplicacion` | `"23"` (producción) | ID en `TbAplicaciones` |
| `EnDesarrollo` | flag TempVar | Modo desarrollo |
| `EnPruebas` | flag TempVar | Modo pruebas |
| `m_TestingMode` | `True/False` (módulo) | Modo testing → sandbox |
| `m_BackendSandboxURL` | ruta al sandbox | URL del backend de testing |
| `m_BackendSandboxPassword` | contraseña | Contraseña del sandbox (env var) |
| `PasswordBackend` | texto (⚠️ D93 hardcoded fallback) | Contraseña del backend producción |
| `m_TextoWin64` | `"(64 bits)"` | Info de versión |
| `rolUsuario` / `rolUsuarioReal` | enum `rol` | Rol del usuario conectado / rol real (impersonación) |
| `g_blnImpersonando` | `True/False` | Flag de impersonación activa |
| `UsarRenderizadoFichero` | `"Sí"` (TempVar) | Toggle renderizado en fichero (más estable, evita bloqueo) vs inyección en memoria (método antiguo) |

En la nueva plataforma estos flags se mueven a **configuración del módulo** (no TempVars, no código):

- `BackendActivo`, `m_BackendSandboxURL`, etc. → config del puerto de persistencia.
- `EnDesarrollo`, `EnPruebas` → variables de entorno del runner.
- `IDAplicacion` → config del módulo.
- `PasswordBackend` → **secret manager / env var** (nunca en repo, ver D93).
- `m_TestingMode` → flag de runtime del runner (no en código de aplicación).
- `UsarRenderizadoFichero` → decisión arquitectónica (en la nueva plataforma, siempre se usa renderizado nativo web).

Para rutas UNC, hosts y nombres de máquina concretos: **NO se reproducen en este artefacto** (regla de evidencia).