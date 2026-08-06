# Condor — capacidades observadas

## Resultado

La aplicación cubre un agregado de **gestión de solicitudes de calidad** vinculadas a Expedientes, con 4 tipos (`PC`, `CD_CA`, `CD_CA_SUB`, `PC_SUB`), workflow con estados, validaciones de calidad, rechazos, adjuntos, snapshots para vistas web y notificaciones. Condor es la **app más activa y arquitectónicamente evolucionada** del ecosistema: tiene ViewModels, Servicios, Repositorios, Edge WebView embebido y un sistema de snapshots. La paridad futura debe incluir como mínimo las capacidades siguientes; ninguna se marca como retirada.

| Dominio | Capacidades evidenciadas | Evidencia principal |
|---|---|---|
| Arranque e identidad | `EVE` con validación crítica (`EVE.ValidacionCritica` líneas 144-168 de `Variables Globales.bas`): valida conexión backend + valida usuario + valida entorno. `getUsuario` (`UsuarioRepositorio.bas:9`) con `getUsuarioConPermisos` | `src/modules/Variables Globales.bas:81` (`EVE`), `:248` (`TbConfiguracionBackends`); `src/modules/UsuarioRepositorio.bas:9` (`getUsuario`); `src/classes/UsuarioServicio.cls:17` (`getUsuarioConPermisos`) |
| Roles | tres niveles (Administrador, Calidad, Técnico) via `rolUsuario`, `rolUsuarioReal` (preparación para impersonación). `JerarquiaRolesHelper.bas` | `src/modules/Variables Globales.bas:68-69`; `src/modules/JerarquiaRolesHelper.bas` |
| Configuración de backend | `TbConfiguracionBackends` con `BackendActivo` (PROD/LOCAL/SANDBOX), `IDAplicacion`, `EnPruebas`/`EnDesarrollo`, `PasswordBackend` (⚠️ D93) | `src/modules/Variables Globales.bas:248` |
| **Testing sandbox seguro** | `m_TestingMode=True` enruta `getdb()` a `m_BackendSandboxURL` con cache safety (Spec-008). Si `m_BackendSandboxURL` está vacío → `TESTS BLOCKED` con error explícito. Patrón maduro de referencia | `src/modules/FUNCIONES UTILES.bas:84-120` |
| Login flow | `EVE` con `m_Command` (lanzado desde `Shell` con correo) o `Wscript.Network.UserName` | `src/modules/Variables Globales.bas:158-168` |
| Validación crítica | `EVE.ValidacionCritica`: si falla conexión, `Application.Quit`. Si falla usuario, `Err.Raise 513`. Patrón fail-fast | `src/modules/Variables Globales.bas:144-168` |
| **Solicitudes** | alta, edición, baja, cambio de estado, rechazo, validación de calidad, snapshot para vista web | `Solicitud.cls`, `SolicitudServicio.cls`, `SolicitudViewModel.cls`, `SolicitudBusquedaViewModel.cls`, `Form_frmGestionSolicitud.cls` |
| 4 tipos de Solicitud | `PC`, `CD_CA`, `CD_CA_SUB`, `PC_SUB` con clases, servicios, repositorios y viewmodels propios | `DatosPC.cls`, `DatosCDCA.cls`, `DatosCDCASUB.cls`, `DatosPCSUB.cls` + sus servicios/repositorios/viewmodels |
| Estados | catálogo de 9 estados + transiciones registradas en `tbTransiciones` | `Estado.cls`, `EstadoServicio.cls`, `EstadoRepositorio.bas` |
| Validación de calidad | `revisionCalidadEstado` (campo) + `revisionCalidadComentarios` (memo) | `Solicitud.cls`, `ValidacionRevision.cls`, `ValidacionRevisionServicio.cls` |
| Recepción de rechazo | historial de rechazos con motivo | `Rechazo.cls`, `RechazoServicio.cls`, `tbHistorialRechazos` |
| Adjuntos | alta, baja, asociación a Solicitud | `Adjunto.cls`, `AdjuntosServicio.cls`, `AdjuntoViewModel.cls`, `AdjuntoRepositorio.bas` |
| **Edge WebView embebido** | `Form_frmGestionSolicitud.cls:209`: `Me.webInfo.Navigate rutaNavegacion` + `WebVisorCacheServicio` (caché de HTML) + `SnapshotServicio` (snapshot del estado). **Vistas web dentro de Access** | `src/forms/Form_frmGestionSolicitud.cls`, `WebVisorCacheServicio.cls`, `SnapshotServicio.cls` |
| Workflow | `WorkflowServicio.cls` (motor de transiciones de estado) | `WorkflowServicio.cls`, `tbTransiciones` |
| Vinculación con Expedientes | `idExpediente` (FK conceptual sin constraint) | `Expediente.cls`, `ExpedienteServicio.cls`, `ExpedienteViewModel.cls`, `ExpedienteRepositorio.bas` |
| Vinculación con NoConformidades | `idNCAsociada` (FK conceptual) + verificación `ncServ.getNoConformidadPorCodigoCondor(codigoSolicitud)` | `Form_frmGestionSolicitud.cls:1908`, `NoConformidad.cls`, `NoConformidadServicio.cls` |
| Suministradores | `Suministrador.cls`, `SuministradorServicio.cls` (vinculación a `ExpedienteSuministrador`) | mismas |
| Documentos | `DocumentoServicio.cls` (gestión de documentos anexos) | mismo |
| Notificaciones | `NotificacionServicio.cls` (envío de avisos) + `MockNotifServ.cls` (test double) | mismos |
| Logs estructurados | Eventos de observabilidad web-native (Sentry / OpenTelemetry / structured logs a Loki/CloudWatch). Las 3 tablas VBA `tbLogCambios`, `tbLogErrores`, `tbLogEstados` NO migran a PostgreSQL; sus llamadas se traducen a eventos web. Confirma D27. | confirma D27 (logs estructurados canónicos) |
| Mapeo de campos | `MapeoCampos.cls`, `MapeoServicio.cls`, `MapeoRepositorio.bas` → `tbMapeoCampos` (config de mapeo entre columnas legacy y modernas) | `src/classes/MapeoCampos.cls` |
| Errores | `CondorError.cls` (clase de error tipada con `.Create`, `.Raise`) | `src/classes/CondorError.cls` |
| Sandbox | `SandboxConfig.cls`, `SandboxGestor.cls`, `SandboxValidator.cls` (gestión del entorno de pruebas) | `src/classes/Sandbox*.cls` |
| Filtros de búsqueda | `FiltrosSolicitud.cls` con 3 callers en forms (`Form_frm0PpalTecnico`, `Form_frmBuscarSolicitudes`, `Form_frmFiltrosAvanzadosSolicitudes`) | `src/classes/FiltrosSolicitud.cls` |
| Hashing | `HashHelper.bas` (utilería de hashing, probablemente para snapshots) | `src/modules/HashHelper.bas` |
| **Tests VBA** | probablemente `tests/` tiene cobertura significativa. `src/tests/` existe. | confirma D87 |
| **Acoplamiento DAO** | `getdb()` con **1 caller** (centralizado en `FUNCIONES UTILES.bas`). `getUsuario()` con **2 callers**. Patrón **diferente** a las otras apps (más DAO-direct). | `src/modules/FUNCIONES UTILES.bas:84`, `src/modules/UsuarioRepositorio.bas:9` |

## Reglas de conservación

- La **identidad se carga UNA vez en `EVE`** y se pasa por `m_ObjUsuarioReal` / `m_ObjUsuarioActivo`. NO se repite en cada operación (análogo al resto del ecosistema).
- `IDAplicacion = "23"` (producción) / probablemente `"23"` (pruebas también, no hay TempVar `EnPruebas` en la rama principal de `EVE`). Mantener como config del módulo.
- **El patrón `m_TestingMode` con sandbox seguro se preserva como referencia** del puerto de caché y de testing de la nueva plataforma.
- **Los ViewModels (`SolicitudViewModel`, etc.) se traducen directamente a DTOs Pydantic** en la nueva plataforma.
- **Los Servicios (`SolicitudServicio`, etc.) se traducen directamente a casos de uso** Python.
- **Los Repositorios (`SolicitudRepositorio`, etc.) se traducen directamente a adaptadores de persistencia** detrás del puerto.
- **Edge WebView embebido** ya no es necesario en la nueva plataforma (la web es nativa, no embebida en Access).
- **El `SnapshotServicio` se traduce a endpoints de snapshot** en la nueva plataforma.
- **El `WorkflowServicio` se traduce a motor de workflow** server-side con transiciones explícitas.
- La **vinculación con NoConformidades** (`idNCAsociada`) se preserva como FK conceptual en PostgreSQL (decisión pendiente D95).

## Evidencia previa

Se han cosechado PRD, Discovery Map, Architecture Overview, ERD, OpenSpec CAP-001..055, UAT y releases antes de inspeccionar staging. CodeGraph-VBA sobre `staging` se consultó primero; el inventario Dysflow real se ejecutó después (`register_worktree` + `accessPath` absoluto explícito): 15 tablas, 5 FKs, 1 solicitud en staging. Las afirmaciones divergentes entre esos documentos y el código quedan abiertas, no resueltas por intención.