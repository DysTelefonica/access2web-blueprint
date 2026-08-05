# HPS_Solicitudes — capacidades observadas

## Resultado

La aplicación cubre un agregado de **gestión de solicitudes HPS**: alta de solicitud, renovación, cambio de tipo, traspaso a ONS, justificación, responsables, adjuntos Excel, plantillas HTML, correos automáticos, registro en HPS, exportación a Lanzadera. Es la **app más simple de las 8** en superficie de datos (11 tablas) pero **en uso activo** (245 solicitudes reales). La paridad futura debe incluir como mínimo las capacidades siguientes; ninguna se marca como retirada.

| Dominio | Capacidades evidenciadas | Evidencia principal |
|---|---|---|
| Arranque e identidad | `EVE` (en `Variables Globales.bas:231`) con `getUsuario` (en `constructor.bas:309`). IDAplicacion `22` (producción). | `src/modules/Variables Globales.bas:231` (EVE), `:458` (getdb); `src/modules/constructor.bas:309` (getUsuario) |
| Configuración de backend | `getConfiguracion` (en `constructor.bas:270`) lee `TbConfiguracion`. TempVars: `EnDesarrollo`, `DatosEnLocal`, `EnPruebas`, `ConCorreoCopiaGestor`, `ActivadoCorreoAutomatico`, `RegistroEnHPS`, `ExpedienteUnificado` | `src/modules/constructor.bas:270` (getConfiguracion) |
| Flags de operación | `ConCorreoCopiaGestor` ("Sí"), `ActivadoCorreoAutomatico` ("Sí"), `RegistroEnHPS` ("Sí"), `ExpedienteUnificado` ("No") | `src/modules/Variables Globales.bas:263-270` |
| Roles | dos niveles (Administrador / Técnico). Si no es Administrador, es Técnico (`EsUsuarioTecnicoCalculado`). Si no es Técnico, error 1000. | `src/modules/Variables Globales.bas:317-336` |
| **Solicitudes** | alta, edición, renovación, cambio de tipo, baja | `Solicitud.cls`, `SolicitudOperaciones.cls`, `SolicitudFechas.cls` |
| **Traspasos a ONS** | traspaso de solicitud al Organismo Notificador de Seguridad (ONS) con adjuntos | `Form_FormAdjuntaTraspasoONS.cls`, `Form_FormSolicitudAltaTraspaso.cls`, `Form_FormSolicitudesAltaTraspasoDatos.cls`, `Form_FormSolicitudesTraspasoFechas.cls`, columna `URLAdjuntoEnvioONS` en `TbSolicitudes` |
| **Adjuntos Excel** | el solicitante adjunta un Excel con datos | `Form_FormAdjuntarExcelSolicitante.cls` |
| **Plantillas HTML** | generación de vistas web con plantillas | `Form_FormPlantillasHTML.cls`, `Form_FormWeb.cls` |
| **Renovación** | workflow de renovación | `Form_FormSolicitudRenovacion.cls`, `Form_FormSolicitudRenovacionDatos.cls`, `Form_FormSolicitudRenovacionFechas.cls`, `Form_FormSolicitudesAltaRenovacionFechas.cls` |
| **Cambio de tipo** | workflow de cambio de tipo de solicitud | `Form_FormCambioTipo.cls` |
| **Justificaciones** | alta de justificaciones sobre la solicitud | `Justificacion.cls`, `JustificacionServicio.cls`, `Form_frmJustificacion.cls`, `Form_frmBuscarJustificacion.cls` |
| **Responsables** | gestión de responsables (vinculados a TbResponsables por email) | `Responsable.cls`, `ResponsableOperaciones.cls`, `Form_FormResponsable.cls`, `Form_FormResponsablesGestion.cls` |
| **Motivos HPS** | catálogo de motivos | `MotivoHPS.cls`, `Form_FormMotivoHPS.cls`, `Form_FormMotivosHPS.cls` |
| **Entidades (suministradores/expedientes)** | búsqueda de entidades relacionadas | `Form_FormEntidades.cls`, `Form_FormExpedientesBusqueda.cls`, `Form_FormExpedienteDetalle.cls` |
| **Tareas pendientes (tramitador)** | workflow de tareas del tramitador | `Form_FormTareasTramitadorPendientes.cls` |
| **Configuración** | gestión de la configuración | `Configuracion.cls`, `ConfiguracionOperaciones.cls`, `Form_FormConfiguracion.cls` |
| **Log general** | log de eventos | `LogGeneral.cls`, `LogGeneralOperaciones.cls` → `TbLogsGeneral` (2058 filas) |
| **Log específico** | log específico de eventos | (presumido en `TbLogs`) |
| **Último cambio** | tracking del último cambio en la solicitud | `UltimoCambio.cls`, `UltimoCambioOperaciones.cls` → `TbUltimoCambio` |
| **Correos** | envío de correos (plantillas + automáticos) | `Correo.cls`, `CorreoOperaciones.cls`, `CorreoServicio.cls`, `Form_FormCorreo.cls`, `Form_FormCorreosGestion.cls`, `Form_FormEnvioCorreo.cls`, `Form_FormCorreoPruebas.cls` → `TbCorreosEnviados` |
| **Registro en HPS** | alta/baja/renovación en HPS (`RegistroEnHPS = "Sí"`) | `IDUsuarioHPS` en `TbSolicitudes`, código que detecta la flag |
| **Expediente unificado** | `ExpedienteUnificado = "No"` (desactivado) | flag en `EVE` |
| **Búsqueda de usuarios HPS** | buscar usuarios HPS para asignar | `Form_FormUsuariosHPSBusqueda.cls` |
| **Automatización** | procesos automatizados (probablemente traspasos, renovaciones) | `AutomatizacionRepositorio.bas`, `Automiatizacion.bas` (typo en el nombre: "Automiatizacion" en lugar de "Automatizacion") |
| **Informes** | generación de informes | `InformesOperaciones.bas` |
| **Búsqueda general** | búsqueda de solicitudes | `Form_FormSolicitudesGestion.cls`, `Form_FormAltaDatosSolicitante.cls` |
| **Tests VBA** | 2 archivos: `Test.bas` (genérico), `TestParametrosParser.bas` (específico de `ParametrosParser.cls`) | confirma D87 con cobertura limitada |
| **Acoplamiento DAO** | `getdb()` con **89 callers** (intermedio) | `src/modules/Variables Globales.bas:458` (`getdb`) |
| **Acoplamiento con Lanzadera** | `getdbLanzadera()` para identidad (igual que el resto) | `src/modules/constructor.bas:358` (en `getUsuario`) |

## Reglas de conservación

- La **identidad se carga UNA vez en `EVE`** vía `getUsuario` y se pasa por `m_ObjUsuarioConectado`. Patrón análogo al resto del ecosistema.
- `IDAplicacion = "22"` se asigna en `EVE` (línea 274). En la nueva plataforma viene de la configuración (D9).
- **7 flags `TempVars` activos** deben traducirse a **configuración del módulo** (no TempVars, no código).
- **El sistema de traspasos a ONS** se traduce a una **integración con servicio externo ONS** vía adaptador (D16, object storage para adjuntos).
- **El sistema de plantillas HTML** se traduce a templates **Jinja2** server-side (D66) con auto-escape.
- **El sistema de adjuntos Excel** se traduce a un **endpoint de upload** con parsing server-side (probablemente `openpyxl` en Python).
- **Los datos personales de `TbSolicitudes` (DNI, nombres, fechas, email, teléfono)** requieren la misma política que HPS (D92).

## Evidencia previa

Se han cosechado PRD, Discovery Map, Architecture Overview, ERD, OpenSpec CAP-001..055, UAT y releases antes de inspeccionar staging. CodeGraph-VBA sobre el repo se consultó primero; el inventario Dysflow real se ejecutó después de `setup_project` + `register_worktree` + `accessPath` + `backendPath` absolutos: 11 tablas, 3 FKs, 245 solicitudes, 28 columnas en `TbSolicitudes`, 2058 filas en `TbLogsGeneral`. Las afirmaciones divergentes entre esos documentos y el código quedan abiertas, no resueltas por intención.