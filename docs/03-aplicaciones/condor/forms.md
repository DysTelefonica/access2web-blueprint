# Condor — formularios, navegación y call paths

## Navegación principal

```text
Form_frmSplash
  -> EVE (Variables Globales.bas:81)
     -> ValidacionCritica: conexión + usuario + entorno
  -> Form_frm0PpalTecnico (opciones generales para técnico)
     -> Form_frmBuscarSolicitudes (búsqueda de solicitudes)
        -> Form_frmFiltrosAvanzadosSolicitudes (filtros avanzados)
     -> Form_frmGestionSolicitud (detalle/edición de solicitud)
        -> Form_frmDatosPC | Form_frmDatosCDCA | Form_frmDatosCDCASUB | Form_frmDatosPCSUB (datos por tipo)
        -> Form_frmPlanPrincipal (workflow / plan)
     -> Form_frmFiltrosAvanzadosSolicitudes (filtros)
  -> Form_frmNotificaciones (notificaciones)
  -> Form_frmConfiguracion (configuración)
  -> Form_frmAyuda (ayuda)
```

## Call paths críticos

| Capacidad | Camino observado | Persistencia / efecto |
|---|---|---|
| Inicio | `frmSplash.Form_Timer → EVE → getdb() (validación) → userServ.getUsuarioConPermisos(usuario) → m_ObjUsuarioReal → m_ObjUsuarioActivo → m_ObjEntorno (validación)` | TempVars, sesión, validación fail-fast |
| Búsqueda de solicitudes | `Form_frmBuscarSolicitudes → FiltrosSolicitud → SolicitudServicio.Buscar → SolicitudRepositorio → SQL` | lee `tbSolicitudes` con filtros |
| Detalle de solicitud | `Form_frmGestionSolicitud.Form_Load → SolicitudServicio.getById(id) → SolicitudViewModel` | hidrata viewmodel con `tbSolicitudes` + relaciones |
| Cambio de estado | `Form_frmGestionSolicitud.cmdCambiarEstado_Click → WorkflowServicio.transicionar(estadoActual, estadoNuevo) → SolicitudServicio.guardar → tbSolicitudes` | transición + log en `tbLogEstados` |
| Datos por tipo (PC) | `Form_frmDatosPC → DatosPCViewModel → DatosPCServicio.GuardarDecisionFinal → tbDatosPC` | upsert por `idSolicitud` |
| Datos por tipo (CD_CA) | `Form_frmDatosCDCA → DatosCDCAViewModel → DatosCDCAServicio.GuardarDecisionFinal → tbDatosCDCA` | mismo patrón |
| Datos por tipo (CD_CA_SUB) | `Form_frmDatosCDCASUB → DatosCDCASUBViewModel → DatosCDCASUBServicio.GuardarDecisionFinal → tbDatosCDCASUB` | mismo patrón |
| Datos por tipo (PC_SUB) | `Form_frmDatosPCSUB → DatosPCSUBViewModel → DatosPCSUBServicio.GuardarDecisionFinal → tbDatosPCSUB` | mismo patrón |
| Adjuntos | `Form_frmAdjuntos → AdjuntosServicio → AdjuntoRepositorio → tbAdjuntos` | upload + persistencia |
| Validación de calidad | `Form_frmGestionSolicitud.cmdRevisionCalidad → ValidacionRevisionServicio → tbValidacionRevision` | log de validación |
| **Edge WebView embebido** | `Form_frmGestionSolicitud.MostrarVista(vista) → WorkflowServicio → HTML generado → WebVisorCacheServicio.cache → Me.webInfo.Navigate rutaTemporal` | **vistas web dentro de Access** |
| **Snapshot para WebView** | `Form_frmGestionSolicitud → SnapshotServicio.take(form, vista) → escribe fichero .html en disco → devuelve ruta` | crea snapshot del estado actual |
| Vinculación NC | `Form_frmGestionSolicitud (línea 1908): ncServ.getNoConformidadPorCodigoCondor(codigoSolicitud)` | verifica NC vinculada antes de eliminar |
| Recepción de rechazo | `Form_frmGestionSolicitud.cmdRechazar → RechazoServicio → tbRechazos + tbHistorialRechazos` | log de rechazo |
| Notificaciones | `Form_frmNotificaciones → NotificacionServicio → MockNotifServ (test)` | envío de avisos |
| Logs | `LogError.bas` → `tbLogErrores` | logging de errores |
| Búsqueda con filtros avanzados | `Form_frmFiltrosAvanzadosSolicitudes → FiltrosSolicitud (3 callers)` | filtros sobre `tbSolicitudes` |

## Inventario normalizado

- **Formularios**: ~30-40 archivos `Form_*.cls` cada uno con `.form.txt` (estimación; el listado completo se obtendrá en iteración posterior). El más relevante es `Form_frmGestionSolicitud` (gestión completa de una solicitud con Edge WebView).
- **Patrón `.cls + .form.txt`**: presente en staging.
- **Clases** (52+ en `src/classes/`):
  - **Domain**: `Solicitud`, `Estado`, `Adjunto`, `Rechazo`, `Expediente`, `Suministrador`, `NoConformidad`, `Usuario`, `Entorno`, `UsuarioAplicacionPermisos`, `LogCambio`, `LogError`, `LogEstado`, `MapeoCampos`, `ValidacionRevision`.
  - **Domain data** (4 tipos de Solicitud): `DatosCDCA`, `DatosCDCASUB`, `DatosPC`, `DatosPCSUB`.
  - **ViewModels**: `SolicitudViewModel`, `SolicitudBusquedaViewModel`, `AdjuntoViewModel`, `DatosCDCAViewModel`, `DatosCDCASUBViewModel`, `DatosPCViewModel`, `DatosPCSUBViewModel`, `ExpedienteViewModel`, `FiltrosSolicitud`.
  - **Servicios**: `SolicitudServicio`, `AdjuntosServicio`, `DatosCDCAServicio`, `DatosCDCASUBServicio`, `DatosPCServicio`, `DatosPCSUBServicio`, `DocumentoServicio`, `EstadoServicio`, `ExpedienteServicio`, `LogCambioServicio`, `NoConformidadServicio`, `NotificacionServicio`, `RechazoServicio`, `RevisionServicio`, `SuministradorServicio`, `UsuarioServicio`, `ValidacionRevisionServicio`, `WebVisorCacheServicio`, `WorkflowServicio`, `MapeoServicio`, `SnapshotServicio`.
  - **Coordinadores transaccionales**: probablemente presentes en `UsuarioServicio` (transacciones de lifecycle).
  - **Sandbox**: `SandboxConfig`, `SandboxGestor`, `SandboxValidator`.
  - **Mocks**: `MockNotifServ`.
  - **Errores**: `CondorError`.
- **Módulos** (muchos en `src/modules/`):
  - **Bootstrap/factory/DAO centralizado**: `Variables Globales.bas` (`EVE`), `UsuarioRepositorio.bas` (`getUsuario`), `FUNCIONES UTILES.bas` (`getdb` con testing sandbox seguro).
  - **Repositorios por entidad**: `AdjuntoRepositorio`, `AplicacionRepositorio`, `CorreoRepositorio`, `DatosCDCARepositorio`, `DatosCDCASUBRepositorio`, `DatosPCRepositorio`, `DatosPCSUBRepositorio`, `EstadoRepositorio`, `ExpedienteRepositorio`, `LogCambioRepositorio`, `LogErrorRepositorio`, `LogEstadoRepositorio`, `MapeoRepositorio`.
  - **Helpers**: `ChecklistHelper`, `DatosTipoSolicitudHelper`, `DecisionFinalHelper`, `DictamenRACDefaultsHelper`, `ErrorLogger`, `ErrorPresenter`, `FormulariosPadreAuxiliares`, `HashHelper`, `JerarquiaRolesHelper`.
  - **JSON**: `JsonConverter`, `JsonHelper`.
  - **Actualizaciones**: `modActualizaciones`.
  - **DevTools**: `modDevTools`.
  - **Batería canonical**: `modBattery_Canonical`.
  - **Patrón MVVM**: presente (ViewModels + Servicios + Repositorios) — análogo a MVVM pero en VBA.
- **Tests VBA**: en `src/tests/` (no inspeccionado en detalle en esta pasada; cobertura probable significativa dado el patrón TDD maduro).
- **Reports/macros/queries**: `src/queries/` (presumido) + `src/reports/` (presumido). Macros embebidas requieren revisión del binario.

## Nota de evidencia

CodeGraph-VBA se consultó primero sobre `staging` y devolvió 98 símbolos en 4 archivos para la query inicial (incluyendo `EVE`, `getUsuario`, `getdb`, `FiltrosSolicitud`). Dysflow read-only se ejecutó después de `register_worktree` + `accessPath` absoluto explícito: 15 tablas, 5 FKs desde `tbSolicitudes`. La inspección de UI se mantiene read-only y no se han alterado formularios.