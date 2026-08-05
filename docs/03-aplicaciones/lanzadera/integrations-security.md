# Lanzadera · Integraciones, seguridad y riesgos

## Baseline y procedencia

La inspección funcional usa `C:\00repos\codigo\00_LANZADERA\staging` (`63ba5e01617fdda857503d43151f06bb7bc11829`); `C:\00repos\codigo\00_LANZADERA\00_main` (`1474e8e8c2a8c352599ffa8b846223c6eb6e0f17`) se conserva como comparación publicada. El backend seleccionado explícitamente para las consultas agregadas fue `C:\00repos\datos\Lanzadera_Datos.accdb`. La configuración Dysflow de staging declara un backend relativo local, por lo que esa ruta compartida es una selección de lectura explícita y no una afirmación de que el `.dysflow/project.json` la persista.

## Contratos con el ecosistema

| Contrato | Evidencia |
|---|---|
| Aplicación | `TbAplicaciones.IDAplicacion`, `NombreEjecutable`, `NombreArchivoDatos`, `NombreCarpeta`, `NombreFuncionPublicacion`, `Comando`; catálogo consultado contiene IDs de Lanzadera, Expedientes (19), HPS (17), Solicitudes HPS (22), Condor (23), Riesgos (5), NC (8) y Brass (6). |
| Identidad de arranque | `VBA.Command`/`/cmd` y `CorreoUsuario`; `m_ObjUsuarioConectadoLogin`; clases `Usuario` y `UsuarioAplicacionPermisos`. |
| Permisos | `TbUsuariosAplicacionesPermisos` por `CorreoUsuario` + `IDAplicacion`; consumidores resuelven `ColAplicacionesPermisos`. |
| Backend | frontend `TbConfiguracionBackends`; `getdb()` elige backend activo y abre DAO; `src/backends.json` declara `backend_principal`. |
| Lanzamiento | `Funciones Genreales.Lanzar` → `Aplicacion.Lanzar` → `EjecutarShelllanzar`; copia de ejecutable/configuración/iconos y argumento `/cmd <correo>`. |
| Auditoría | `TbConexionesRegistro` y `TbAplicacionesAperturas`; el código registra conexión/apertura y máquina/ubicación. |
| Correo | `Correo.EnviarCorreo` encola en `TbUsuariosCorreosEnvio` (código referencia también `TbCorreosEnviados`); el dispatcher externo se invoca aproximadamente cada cinco minutos (ver `04-integraciones-y-operacion/informes-exportaciones-y-correo.md`). El código contiene un BCC fijo interno: no se reproduce aquí. **Verified-static + discovery `legacy-email-queue-flow`.** |

No hay informes Access inventariados y `export_queries` devolvió cero consultas frontend. No se demostró ejecución de batch, exportación Excel, macros, API externa ni lectura efectiva de archivos formativos; quedan como `Likely`/gap.

## Contratos objetivo que sustituyen los mecanismos legacy

Las tablas anteriores describen **cómo funciona Lanzadera hoy**. La plataforma web sustituye varios de esos mecanismos por servicios compartidos hexagonal-puros. Esta sub-tabla es la **dirección de sustitución**, no el diseño detallado (que vive en `docs/09-arquitectura-objetivo-y-principios.md`):

| Contrato legacy | Contrato objetivo | Disposición | Origen |
|---|---|---|---|
| `TbConfiguracionBackends` + `Config_BackendHelper.bas` + rutas hardcoded | Configuración técnica externalizada tras un adapter/port; secretos por proveedor reemplazable. | Retirar la administración de plataforma; los valores se migran a configuración de despliegue, no a UI. | `product/lanzadera-backend-config-disposition` |
| `Variables Globales.getdb()` / `getdbLanzadera` / `getdbExpedientes` | Persistencia por adapter (PostgreSQL por esquema); cada módulo detrás de su port. | Sustituir mecanismo; `getdb` deja de ser el camino de acceso. | `architecture/global-hexagonal-principle` + `architecture/target-database-topology` |
| `Funciones Genreales.Lanzar` + `Aplicacion.Lanzar` + `EjecutarShelllanzar` | Entradas permission-aware del menú global web; sin `Shell` ni `/cmd`. | Retirar mecanismo desktop. | `product/lanzadera-launcher-disposition` |
| `Form_FormAplicacionesOficina` / `FueraOficina` con `EjecucionEnOficina` | Visibilidad única por capacidades; diferenciación por entorno (UAT/prod) cuando aplique. | Retirar segmentación por ubicación. | `product/lanzadera-location-visibility` |
| `Correo.EnviarCorreo` + `TbUsuariosCorreosEnvio` | Servicio unificado de notificaciones detrás de un port; la cola por tabla queda como adaptador transitorio v1. | Sustituir mecanismo; conservar tabla solo como adapter. | `architecture/unified-notification-service` + `architecture/notification-delivery-adapter-v1` |
| `TbConfiguracionBackends.PasswordBackend`, `PassIncialPlana`, `TbUsuariosHistoricoContrasenias.PassAntigua` | Adaptador de credenciales con verificación legacy versionada y rehash transparente al primer login; contraseñas planas/histórico no se migran como secretos. | Retirar campos sensibles heredados; rehashing oportunista posterior. | `product/credential-migration` + `architecture/opportunistic-password-rehash` |
| `TbConexionesRegistro` (con SSID/ubicación/coordenadas) | Logs canónicos con correlación; módulo de auditoría propio. Telemetría de ubicación retirada. | Preservar evento de autenticación/apertura; retirar SSID/coordenadas/ubicación. | `product/lanzadera-audit-disposition` |

## Seguridad observada

- **Hash:** `Critografia.SHA256` usa `CALG_SHA_256` de `advapi32.dll` y devuelve hexadecimal en minúsculas; `Login` compara el hash de la contraseña introducida con `TbUsuariosAplicaciones.Password` (`Verified-static`). No hay evidencia de salt, factor adicional o algoritmo versionado.
- **Almacenamiento sensible:** `Password` es texto de 255 caracteres; `PassIncialPlana` existe en el esquema y en la clase. `TbUsuariosHistoricoContrasenias.PassAntigua` también es texto. No se han leído valores de datos ni hashes.
- **Autenticación:** correo + contraseña; administrador de máquina puede obtener estado correcto sin el mismo camino de contraseña; SSO usa archivo/usuario de máquina (`Verified-static`, revisar antes de migrar).
- **Ciclo de credencial:** `ResetearPass`/`CambioPass` escriben credenciales e histórico mediante DAO; el reset notifica mediante `Correo.EnviarCorreo`. Caducidad, flag de cambio, bloqueo por fecha/intentos y contraseña de un solo uso son `Verified-static`; no se ejecutaron pruebas.
- **Autorización:** UI y clases filtran por permisos; administrador global recibe todas las aplicaciones activas. El enforcement no está demostrado en un servicio/backend aislado.
- **Excepción pendiente:** Expedientes se fuerza visible/accesible en dos formularios; el significado operativo del ID 51 no queda resuelto por el grafo ni por el catálogo consultado.
- **Configuración/secrets:** `TbConfiguracionBackends.PasswordBackend` existe; `Config_BackendHelper.bas` mantiene rutas y constantes de entorno hardcodeadas, aunque la configuración actual también las persiste. Debe eliminarse la duplicidad antes de migrar.

## Uso observado y privacidad

Las consultas fueron `SELECT` agregados únicamente: catálogo por aplicación, aperturas por aplicación, conexiones, asignaciones de permisos, visionados, cola de correo, histórico de contraseñas, flags de tareas y perfiles. No se devolvieron filas crudas, identidades, destinatarios, asunto/cuerpo, credenciales, hashes, equipos, coordenadas ni secretos. Las cifras y fechas resumidas están en `capabilities.md`; no prueban valor futuro ni sustituyen la decisión con el usuario.

## Pain points y disposición

La modernización debe tratar como riesgo principal la duplicación entre formulario, clases de dominio, módulos globales y rutas/configuración. También hay SQL concatenado, `On Error Resume Next`, errores mostrados como UI, acoplamiento a ActiveX/Windows/UNC y contratos de texto `Sí/No`. Disposición semilla: **preservar** contratos de IDs y permisos durante transición; **reemplazar mecanismo** de Access/Shell/ActiveX; **fusionar** identidad, autorización, auditoría, correo y configuración en servicios compartidos; **abrir decisión** sobre vídeos, tareas y legado de aplicaciones inactivas. Los 76 callers internos de `getdb` confirman que no es un detalle local aislado.
