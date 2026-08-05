# HPS — integraciones y automatización

## Contratos observados

| Sistema | Uso | Evidencia |
|---|---|---|
| **Lanzadera** (identidad y permisos) | identidad, permisos, entorno (probablemente vía `getdbLanzadera()` o tablas compartidas) | `Constructor.bas:1415` (`getUsuario`); `UsuarioAplicacionPermisos.cls` |
| **Expedientes** (vinculación conceptual) | búsqueda de expediente por código (sin FK física) | `Form_FormExpedientesBusqueda.cls`, `Expediente.cls` (clase compartida con Lanzadera) |
| **SICA** (sistema externo) | `TbUsuariosSICA` + `TbAnexosUsuariosSICA` (vinculación con sistema SICA externo) | `UsuarioSICA.cls`, `AnexoUsuarioSICA.cls` |
| **HPS_Solicitudes** | `IDSolicitud` como FK conceptual (sin FK física) | `TbUsuarios.IDSolicitud` |
| AGEDYS | presumiblemente DPD/proyectos (vía código compartido) | `Expediente.cls` compartido con Lanzadera |
| Correos | avisos de error al administrador, notificación de cursos | `Correo.cls`, `modIndicadores.bas` |
| SharePoint/ficheros | enlaces de documentación | referencias en `TbAnexos*` |

## Acoplamiento declarado con Lanzadera

La aplicación accede a la identidad/permisos del usuario vía `Constructor.bas:getUsuario` que consulta `TbUsuariosAplicaciones`. El patrón es idéntico al de Lanzadera/Expedientes/Gestion_Riesgos/NoConformidades.

Implicaciones para la nueva plataforma:

- La identidad debe resolverse vía el **adaptador unificado de autenticación** (D9-D10), no vía acceso directo a la BD de Lanzadera.
- Los permisos por aplicación se cargan con `IDAplicacion = "17"` (producción) o `"51"` (pruebas).
- El acceso directo a `TbUsuariosAplicaciones` en la BD de Lanzadera debe **eliminarse** y sustituirse por el servicio de identidad/permisos expuesto por la plataforma.

## Automatización y tareas

HPS tiene **3+ módulos de caché** propios que coordinan el rendimiento:

- `cacheUsuario.bas`: caché de usuarios.
- `cacheSuministrador.bas`: caché de suministradores.
- `Mod_Cache_Core.bas`: núcleo del sistema de caché.
- `Mod_StartupCacheInitialization.bas`: inicialización al arranque.
- `CacheConsistencyAudit.bas`: auditoría de consistencia del caché.
- `Mod_Sincronizacion_Historico.bas`: sincronización histórico ↔ anexos.

La caché se controla vía `TbConfiguracion.CacheHabilitada` (kill switch), leído por `LeeConfiguracionLocal`. Patrón idéntico al de NoConformidades (`TbCacheNCProyecto` + `InicializadorCache.bas`).

**Confirmación cross-app de D91**: tanto HPS como NoConformidades implementan el mismo patrón de caché selectivo maduro. La nueva plataforma debe **preservar** este patrón como referencia del puerto de caché.

## Configuración y flags de operación

Flags activos hoy en producción (inferidos de `VariablesEntorno.bas`):

| Flag / Campo | Valor por defecto | Significado |
|---|---|---|
| `BackendActivo` | `PROD` (producción) / `LOCAL` / `SANDBOX` | Cuál backend se usa |
| `EnPruebas` | `Sí`/`No` (texto) | Si está en pruebas; cuando `"Sí"`, IDAplicacion pasa a `"51"` |
| `IDAplicacion` | `"17"` (producción) / `"51"` (pruebas) | ID en `TbAplicaciones` |
| `PasswordBackend` | texto | Contraseña del backend (⚠️ D92) |
| `CacheHabilitada` (en `TbConfiguracion`) | Boolean | Habilita/deshabilita la caché (kill switch) |
| `BackendSandboxURL` / `BackendSandboxPassword` (testing mode) | — | Solo en `m_TestingMode = True` |

En la nueva plataforma estos flags se mueven a **configuración del módulo** (no TempVars):

- `BackendActivo`, `BackendSandboxURL`, etc. → config del puerto de persistencia.
- `EnPruebas` → variable de entorno del runner.
- `IDAplicacion` → config del módulo.
- `PasswordBackend` → **secret manager / env var** (nunca en repo, ver D92).
- `CacheHabilitada` → config del módulo + kill switch runtime vía endpoint admin.

Para rutas UNC, hosts y nombres de máquina concretos: **NO se reproducen en este artefacto** (regla de evidencia).