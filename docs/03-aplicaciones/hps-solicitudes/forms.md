# HPS_Solicitudes — formularios, navegación y call paths

## Navegación principal

```text
Form_frmSplash
  -> EVE (Variables Globales.bas:231)
     -> TbConfiguracion + TbConfiguracionBackends (TempVars + BackendActivo)
     -> getdbLanzadera().TbUsuariosAplicaciones (getUsuario)
  -> Form_Form0BDOpciones (opciones generales)
      ├─> Form_FormSolicitudesGestion (gestión/búsqueda de solicitudes)
      │    ├─> Form_FormSolicitudAlta (alta de solicitud)
      │    │    ├─> Form_FormSolicitudesAltaFechas (fechas de alta)
      │    │    ├─> Form_FormSolicitudesAltaRenovacionFechas (renovación)
      │    │    ├─> Form_FormAdjuntarExcelSolicitante (adjuntar Excel)
      │    │    └─> Form_FormMotivoHPS / Form_FormMotivosHPS (motivos)
      │    ├─> Form_FormSolicitudAltaTraspaso (alta con traspaso a ONS)
      │    │    ├─> Form_FormSolicitudesAltaTraspasoDatos (datos del traspaso)
      │    │    └─> Form_FormSolicitudesTraspasoFechas (fechas del traspaso)
      │    ├─> Form_FormSolicitudRenovacion (renovación)
      │    │    ├─> Form_FormSolicitudRenovacionDatos (datos de la renovación)
      │    │    └─> Form_FormSolicitudRenovacionFechas (fechas de la renovación)
      │    ├─> Form_FormCambioTipo (cambio de tipo de solicitud)
      │    └─> Form_FormAdjuntaTraspasoONS (adjuntar traspaso a ONS)
      ├─> Form_FormResponsablesGestion (gestión de responsables)
      │    └─> Form_FormResponsable (detalle de responsable)
      ├─> Form_FormEntidades (entidades relacionadas)
      │    └─> Form_FormExpedienteDetalle / Form_FormExpedientesBusqueda
      ├─> Form_FormTareasTramitadorPendientes (tareas del tramitador)
      ├─> Form_FormPlantillasHTML / Form_FormWeb (vistas web)
      ├─> Form_FormConfiguracion (configuración)
      └─> Form_FormCorreo / Form_FormEnvioCorreo / Form_FormCorreosGestion (correos)
```

## Call paths críticos

| Capacidad | Camino observado | Persistencia / efecto |
|---|---|---|
| Inicio | `frmSplash.Form_Timer → EVE → TbConfiguracion (TempVars) → getdbLanzadera().TbUsuariosAplicaciones (getUsuario) → m_ObjUsuarioConectado` | TempVars (7 flags), sesión, validación fail-fast |
| Búsqueda de solicitudes | `Form_FormSolicitudesGestion → TbSolicitudes (filtros)` | consulta SQL directa |
| Alta de solicitud | `Form_FormSolicitudAlta → Form_FormSolicitudesAltaFechas → TbSolicitudes (Insert) → TbSolicitudesFechas (Insert) → TbResponsables (vinculación)` | transacción DAO multi-tabla |
| Renovación | `Form_FormSolicitudRenovacion → Form_FormSolicitudRenovacionDatos → Form_FormSolicitudRenovacionFechas → TbSolicitudes (Update) → TbSolicitudesFechas (Insert)` | idem alta |
| Cambio de tipo | `Form_FormCambioTipo → TbSolicitudes (Update Tipo)` | cambio del campo `Tipo` |
| **Traspaso a ONS** | `Form_FormAdjuntaTraspasoONS → URLAdjuntoEnvioONS (Update) → envío` | integración con sistema externo ONS |
| Adjuntar Excel | `Form_FormAdjuntarExcelSolicitante → URLAdjunto (Update) → upload` | upload del fichero Excel del solicitante |
| Justificación | `Form_frmJustificacion → Form_frmBuscarJustificacion → Justificacion.Registrar → TbJustificaciones (Insert)` | log de justificación |
| Responsables | `Form_FormResponsablesGestion → Form_FormResponsable → Responsable.Registrar → TbResponsables (Insert) → vinculación por email` | gestión de responsables (⚠️ FK por email) |
| Tareas del tramitador | `Form_FormTareasTramitadorPendientes → TareasServicio (no inspeccionado)` | worklist |
| Correos automáticos | `Form_FormEnvioCorreo → CorreoServicio.Enviar → TbCorreosEnviados (Insert)` | envío de correos |
| Log | `LogGeneral.Registrar → TbLogsGeneral (Insert)` (2058 filas) | log de eventos |
| **Plantillas HTML** | `Form_FormPlantillasHTML → genera HTML → Form_FormWeb.Navigate` | vista web embebida (similar a Condor) |
| **Registro en HPS** | `EVE lee TempVar RegistroEnHPS = "Sí" → código que vincula TbSolicitudes con HPS via IDUsuarioHPS` | integración con HPS |
| Configuración | `Form_FormConfiguracion → TbConfiguracion (Update)` | gestión de config |

## Inventario normalizado

- **Formularios**: ~30-40 archivos `Form_*.cls` cada uno con `.form.txt` (estimación; listado completo arriba). Patrón `.cls + .form.txt` presente.
- **Clases** (26 en `src/classes/`):
  - **Domain**: `Solicitud`, `SolicitudFechas`, `Justificacion`, `Responsable`, `MotivoHPS`, `Expediente`, `Suministrador`, `Usuario`, `UsuarioHPS`, `UsuarioEntidad`, `Configuracion`, `Correo`, `LogGeneral`, `UltimoCambio`.
  - **Operaciones**: `SolicitudOperaciones`, `SolicitudFechasOperaciones`, `JustificacionServicio`, `ResponsableOperaciones`, `ConfiguracionOperaciones`, `CorreoOperaciones`, `LogGeneralOperaciones`, `UltimoCambioOperaciones`.
  - **Compartidas con Lanzadera**: `UsuarioAplicacionPermisos`, `Entorno`.
  - **Util**: `ParametrosParser`.
- **Módulos** (14 en `src/modules/`):
  - **Bootstrap/factory/DAO centralizado**: `Variables Globales.bas` (EVE + getdb), `constructor.bas` (getUsuario + getConfiguracion), `FUNCIONES UTILES.bas`, `Factoria.bas`, `RepositorioComun.bas`.
  - **Repositorios**: `AutomatizacionRepositorio.bas`, `CorreoRepositorio.bas`, `JustificacionRepositorio.bas`.
  - **Tests**: `Test.bas`, `TestParametrosParser.bas` (cobertura limitada, 2 archivos).
  - **Otros**: `Automiatizacion.bas` (typo en el nombre: "Automiatizacion"), `InformesOperaciones.bas`, `Instalador.bas`, `JsonConverter.bas`.
- **Tests VBA**: 2 archivos (cobertura básica).
- **Reports/macros/queries**: no inspeccionados en detalle en esta pasada.

## Nota de evidencia

CodeGraph-VBA se consultó primero sobre el repo y devolvió 44 símbolos en 2 archivos para la query inicial. Dysflow read-only se ejecutó después de `setup_project` (creó `.dysflow/project.json` en `HPS_SOLICITUDES/` con permiso del user) + `register_worktree` + `accessPath` + `backendPath` absolutos: 11 tablas, 3 FKs, 245 solicitudes, 2058 logs. La inspección de UI se mantiene read-only y no se han alterado formularios.